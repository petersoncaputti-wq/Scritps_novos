"""Streaming com retomada Range vinculada a identidade e validador HTTP."""
import json
import logging
import re
import time
from pathlib import Path
from urllib.parse import urlsplit


def transfer_file(session, url, target, size, progress=None, resume_key=None, stop=None):
    from browser_bulk import DownloadError, retry_delay
    partial = target.with_name(target.name + '.spdownload.part')
    metadata = partial.with_name(partial.name + '.json')
    identity = resume_key or url
    saved = {}
    if metadata.exists():
        try:
            saved = json.loads(metadata.read_text(encoding='utf-8'))
        except (ValueError, OSError):
            pass
    if partial.exists() and not saved:
        raise DownloadError('Parcial sem identificacao; preserve ou mova o arquivo .spdownload.part antes de repetir.')
    offset = partial.stat().st_size if partial.exists() else 0
    initial_offset = offset
    started = time.monotonic()
    if saved.get('identity') != identity or saved.get('size') != size or offset >= size:
        offset = 0
    validator = saved.get('etag')
    if not validator or validator.startswith('W/'):
        offset = 0
    headers = {'Accept-Encoding': 'identity'}
    if offset:
        headers.update({'Range': f'bytes={offset}-', 'If-Range': validator})
        logging.info('PARCIAL_RETOMADA_SOLICITADA local=%s bytes_existentes=%s total_bytes=%s', target, offset, size)
    elif initial_offset:
        logging.info('PARCIAL_REINICIO local=%s bytes_existentes=%s motivo=identidade_tamanho_ou_validador', target, initial_offset)
    with session.get(url, stream=True, timeout=(20, 60), headers=headers) as response:
        logging.info('DOWNLOAD_RESPOSTA local=%s http=%s', target, response.status_code)
        if response.status_code not in (200, 206):
            from http_diagnostics import error_details
            detail = error_details(response)
            logging.warning('SHAREPOINT_HTTP_ERRO local=%s http=%s detalhes=%s', target, response.status_code, json.dumps(detail, ensure_ascii=False))
            message = f'Download recusado: HTTP {response.status_code}.'
            if detail['codigo']:
                message += f" Codigo: {detail['codigo']}."
            if detail['mensagem']:
                message += f" Mensagem: {detail['mensagem']}"
            raise DownloadError(message, response.status_code,
                                retry_delay(response.headers.get('Retry-After')))
        if urlsplit(response.url).hostname != urlsplit(url).hostname:
            raise DownloadError('Download redirecionado para outro dominio; refaca o login.')
        if 'text/html' in response.headers.get('Content-Type', '').lower() and target.suffix.lower() not in ('.html', '.htm'):
            raise DownloadError('Download retornou HTML; a sessao pode ter expirado.')
        if response.status_code == 206:
            match = re.fullmatch(r'bytes (\d+)-(\d+)/(\d+)', response.headers.get('Content-Range', ''))
            if (not offset or not match or int(match[1]) != offset or int(match[3]) != size
                    or int(match[2]) < offset or int(match[2]) >= size
                    or response.headers.get('ETag') != validator):
                raise DownloadError('Resposta Range ou versao invalida; parcial preservado.')
            logging.info('PARCIAL_RETOMADA_CONFIRMADA local=%s offset=%s', target, offset)
        else:
            if offset:
                logging.info('PARCIAL_REINICIO local=%s motivo=servidor_retornou_200', target)
            offset = 0  # Server ignored Range or the remote version changed.
        metadata.write_text(json.dumps({'identity': identity, 'size': size,
                                        'etag': response.headers.get('ETag')}), encoding='utf-8')
        written = offset
        with partial.open('ab' if offset else 'wb') as output:
            if progress:
                progress(written)
            for chunk in response.iter_content(1024 * 1024):
                if stop and stop.is_set():
                    logging.warning('TRANSFERENCIA_CANCELADA local=%s bytes_gravados=%s parcial=%s', target, written, partial)
                    raise DownloadError('Transferencia interrompida; parcial preservado.')
                output.write(chunk)
                written += len(chunk)
                if progress:
                    progress(written)
    if written != size:
        # A complete but different-sized response cannot safely be resumed.
        metadata.write_text(json.dumps({'identity': identity, 'size': size, 'etag': None}), encoding='utf-8')
        raise DownloadError(f'Tamanho divergente: esperado {size:,} bytes; recebido {written:,} bytes. O arquivo pode ter mudado apos o inventario.')
    if target.exists():
        raise DownloadError('Destino criado durante o download; preservado sem sobrescrever.')
    partial.rename(target)
    logging.info('TRANSFERENCIA_VALIDADA local=%s bytes=%s bytes_reutilizados=%s duracao_segundos=%.2f', target, written, offset, time.monotonic()-started)
    metadata.unlink(missing_ok=True)

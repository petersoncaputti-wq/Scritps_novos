"""Comparacao de rotas de leitura com amostras limitadas, sem salvar documentos."""
import json
import logging
import re
import uuid
from urllib.parse import quote, urlsplit

import requests
from config import BASE
from browser_bulk import copy_cookies, safe_error
from browser_test import read_json
from http_diagnostics import error_details, sanitize
from sharepoint import parse_sharepoint_url


def choose_target(url='', remote=''):
    if not remote:
        logs = sorted((BASE / 'logs').glob('execucao_*.log'), key=lambda p: p.stat().st_mtime, reverse=True)
        for path in logs:
            with path.open(encoding='utf-8') as source:
                for line in source:
                    match = re.search(r'DOWNLOAD_FALHA_FINAL remoto=(.*?) local=', line)
                    if match:
                        remote = match[1]
                        break
            if remote:
                if not url:
                    with path.open(encoding='utf-8') as source:
                        for line in source:
                            match = re.search(r'ORIGEM hostname=([^ ]+)', line)
                            if match:
                                url = 'https://' + match[1].strip() + quote(remote.rsplit('/', 1)[0], safe='/')
                                break
                break
    if not remote or not remote.startswith('/') or remote.startswith('//'):
        raise ValueError('Nenhuma falha com caminho remoto encontrada. Informe --diagnose-file e --url.')
    if not url:
        url = input('URL da pasta SharePoint: ').strip()
    parsed = parse_sharepoint_url(url)
    if not remote.startswith(parsed.path.rstrip('/') + '/'):
        raise ValueError('O arquivo de diagnostico deve estar dentro da pasta informada.')
    print(f'Arquivo escolhido para diagnostico: {remote}')
    return url, remote


def probe(session, label, endpoint, metadata=False):
    result = {'teste': label}
    payload = None
    try:
        with session.get(endpoint, headers={'Accept': 'application/json;odata=nometadata' if metadata else '*/*',
                                            'Accept-Encoding': 'identity'},
                         stream=True, timeout=(20, 60), allow_redirects=False) as response:
            result['http'] = response.status_code
            result['formato'] = sanitize(response.headers.get('Content-Type', ''), 100)
            if response.status_code not in (200, 206):
                result['erro'] = error_details(response)
                result['redirecionou'] = response.is_redirect
            else:
                data = bytearray()
                for chunk in response.iter_content(4096):
                    data.extend(chunk[:32768-len(data)])
                    if len(data) >= 32768:
                        break
                result['bytes_amostra'] = len(data)
                result['assinatura_zip'] = data.startswith(b'PK\x03\x04')
                if metadata:
                    try:
                        body = json.loads(data)
                        payload = body.get('d', body)
                        result['metadados'] = {key: sanitize(payload[key], 400) for key in
                                               ('Name', 'Length', 'Exists', 'UniqueId', 'TimeLastModified') if key in payload}
                    except (ValueError, TypeError, AttributeError):
                        result['parse_metadados'] = 'resposta_nao_identificada'
    except (requests.RequestException, OSError) as error:
        result['erro_local'] = safe_error(error)
    logging.info('DIAGNOSTICO_RESULTADO detalhes=%s', json.dumps(result, ensure_ascii=False))
    print(f"{label}: HTTP {result.get('http', 'indisponivel')}; tipo={result.get('formato', '')}; amostra={result.get('bytes_amostra', 0)} bytes")
    return payload


def diagnose(page, web, remote):
    logging.info('DIAGNOSTICO_INICIO remoto=%s site=%s', remote, web)
    literal = quote(remote.replace("'", "''"), safe='/')
    file_api = web + "/_api/web/GetFileByServerRelativePath(decodedUrl='" + literal + "')"
    select = '?$select=Name,Length,Exists,UniqueId,TimeLastModified'
    with requests.Session() as session:
        copy_cookies(page, session)
        metadata = probe(session, 'metadados_terminal', file_api + select, metadata=True)
        browser_metadata = None
        try:
            browser_metadata = read_json(page, file_api + select)
            logging.info('DIAGNOSTICO_METADADOS_NAVEGADOR estado=ok nome=%s bytes=%s unique_id=%s',
                         sanitize(browser_metadata.get('Name')), sanitize(browser_metadata.get('Length')),
                         sanitize(browser_metadata.get('UniqueId')))
            print('metadados_navegador: leitura concluida')
        except RuntimeError:
            logging.warning('DIAGNOSTICO_METADADOS_NAVEGADOR estado=falha')
            print('metadados_navegador: leitura recusada')
        probe(session, 'conteudo_por_caminho', file_api + '/$value')
        unique_id = (browser_metadata or metadata or {}).get('UniqueId')
        if unique_id:
            try:
                unique_id = str(uuid.UUID(unique_id))
                probe(session, 'conteudo_por_id', web + f"/_api/web/GetFileById('{unique_id}')/$value")
            except (ValueError, AttributeError, TypeError):
                logging.warning('DIAGNOSTICO_POR_ID ignorado=identificador_invalido')
        else:
            logging.warning('DIAGNOSTICO_POR_ID ignorado=sem_identificador')
        origin = 'https://' + urlsplit(web).hostname
        probe(session, 'conteudo_url_direta', origin + quote(remote, safe='/'))
        probe(session, 'conteudo_download_aspx', web + '/_layouts/15/download.aspx?SourceUrl=' + quote(remote, safe=''))
    logging.info('DIAGNOSTICO_FIM remoto=%s validacao=amostra_apenas', remote)
    print('Diagnostico concluido. Nenhum documento foi salvo; consulte DIAGNOSTICO_RESULTADO no log.')
    return 0

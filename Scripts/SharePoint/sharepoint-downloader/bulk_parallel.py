"""Workers HTTP; Playwright, progresso e SQLite ficam na thread principal."""
import logging
import queue
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import quote, urlsplit

import requests


def transfer_with_fallback(session, web, path, target, size, progress, version, stop, transfer):
    from browser_bulk import DownloadError
    literal = quote(path.replace("'", "''"), safe='/')
    endpoints = [
        ('rest_caminho', web + "/_api/web/GetFileByServerRelativePath(decodedUrl='" + literal + "')/$value"),
        ('url_direta', 'https://' + urlsplit(web).hostname + quote(path, safe='/')),
        ('download_aspx', web + '/_layouts/15/download.aspx?SourceUrl=' + quote(path, safe='')),
    ]
    for index, (route, endpoint) in enumerate(endpoints):
        try:
            transfer(session, endpoint, target, size, progress, resume_key=path + '|' + version, stop=stop)
            logging.info('DOWNLOAD_ROTA_SUCESSO remoto=%s local=%s rota=%s', path, target, route)
            return
        except DownloadError as error:
            if error.status not in (400, 404) or index == len(endpoints) - 1:
                raise
            # Only switch route on a rejected request, never hide auth, network,
            # integrity, or disk failures. Full size validation stays in transfer.
            logging.warning('DOWNLOAD_ROTA_ALTERNATIVA remoto=%s rota_recusada=%s http=%s proxima_rota=%s',
                            path, route, error.status, endpoints[index + 1][0])


def run_parallel(db, page, web, root, progress, renew, workers, transfer):
    from browser_bulk import DownloadError, safe_error
    stop = threading.Event()
    updates = queue.Queue()
    active = {}
    jobs = {}
    retry = []
    downloaded = skipped = failed = 0
    limit = workers
    cooldown = 0
    rows = iter(db.execute('SELECT remote,local,size,version FROM inventory ORDER BY remote'))
    exhausted = False
    last_log = time.monotonic()

    def worker(row, cookies):
        path, text, size, version = row
        with requests.Session() as session:
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'],
                                    path=cookie['path'], secure=cookie.get('secure', True))
            transfer_with_fallback(session, web, path, Path(text), size,
                                   lambda n: updates.put((path, n)), version, stop, transfer)

    def prepare(row):
        nonlocal skipped, failed
        path, text, size, version = row
        target = Path(text)
        if not target.resolve().is_relative_to(root.resolve()):
            raise ValueError('Arquivo local fora do destino.')
        target.parent.mkdir(parents=True, exist_ok=True)
        record = db.execute('SELECT local,size,version,status FROM files WHERE remote=?', (path,)).fetchone()
        if target.exists():
            same_version = not record or record[2] == version
            if target.is_file() and target.stat().st_size == size and same_version:
                skipped += 1
                db.execute('INSERT OR REPLACE INTO files VALUES (?,?,?,?,?)', (path, text, size, version, 'existing'))
                db.commit()
                progress.message(f'Existente: {target.name}')
                logging.info('ARQUIVO_EXISTENTE remoto=%s local=%s bytes=%s verificacao=tamanho_e_versao_conhecida', path, text, size)
                progress.complete(size)
            else:
                failed += 1
                progress.failed = failed
                progress.message(f'Conflito: {target.name} (detalhes no log)')
                logging.warning('ARQUIVO_CONFLITO remoto=%s local=%s esperado_bytes=%s local_bytes=%s motivo=tamanho_ou_versao', path, text, size, target.stat().st_size if target.is_file() else 'nao_e_arquivo')
            return False
        db.execute('INSERT OR REPLACE INTO files VALUES (?,?,?,?,?)', (path, text, size, version, 'pending'))
        db.commit()
        return True

    executor = ThreadPoolExecutor(max_workers=workers)
    try:
        while not exhausted or jobs or retry:
            while len(jobs) < limit and time.monotonic() >= cooldown:
                now = time.monotonic()
                ready = next((i for i, entry in enumerate(retry) if entry[0] <= now), None)
                if ready is not None:
                    _, row, attempt = retry.pop(ready)
                elif not exhausted:
                    try:
                        row = next(rows)
                    except StopIteration:
                        exhausted = True
                        continue
                    attempt = 1
                    if not prepare(row):
                        continue
                else:
                    break
                cookies = page.context.cookies([web])
                progress.message(f'Baixando: {Path(row[1]).name}')
                logging.info('DOWNLOAD_INICIO remoto=%s local=%s bytes=%s tentativa=%s', row[0], row[1], row[2], attempt)
                jobs[executor.submit(worker, row, cookies)] = (row, attempt)
                active[row[0]] = 0
            page.wait_for_timeout(100)
            while not updates.empty():
                path, count = updates.get_nowait()
                if path in active:
                    active[path] = count
            progress.show(sum(active.values()))
            if time.monotonic() - last_log >= 30:
                last_log = time.monotonic()
                logging.info('LOTE_PROGRESSO concluidos=%s total_arquivos=%s bytes_concluidos=%s bytes_em_transferencia=%s total_bytes=%s ativos=%s falhas=%s', progress.completed_files, progress.total_files, progress.completed_bytes, sum(active.values()), progress.total_bytes, len(jobs), failed)
            for future in list(jobs):
                if not future.done():
                    continue
                row, attempt = jobs.pop(future)
                path, text, size, version = row
                active.pop(path, None)
                try:
                    future.result()
                    db.execute("UPDATE files SET status='done' WHERE remote=?", (path,))
                    db.commit()
                    downloaded += 1
                    progress.complete(size)
                    progress.message(f'Concluido: {Path(text).name}')
                    logging.info('DOWNLOAD_CONCLUIDO remoto=%s local=%s bytes=%s tentativa=%s', path, text, size, attempt)
                except (requests.RequestException, RuntimeError, OSError) as error:
                    reason = safe_error(error)
                    progress.message(f'Nova tentativa: {Path(text).name} (detalhes no log)')
                    logging.warning('DOWNLOAD_ERRO remoto=%s local=%s tentativa=%s motivo=%s', path, text, attempt, reason)
                    if attempt >= 3:
                        failed += 1
                        progress.failed = failed
                        progress.message(f'Falha: {Path(text).name} (detalhes no log)')
                        logging.error('DOWNLOAD_FALHA_FINAL remoto=%s local=%s tentativas=%s motivo=%s', path, text, attempt, reason)
                        continue
                    delay = getattr(error, 'retry_after', None)
                    delay = delay if delay is not None else attempt * 3
                    if isinstance(error, DownloadError) and error.status in (429, 503):
                        limit = 1
                        cooldown = max(cooldown, time.monotonic() + delay)
                        progress.message(f'Limitado pelo SharePoint; pausa de {delay:.0f}s')
                        logging.warning('LOTE_LIMITADO http=%s concorrencia=1 espera_segundos=%.2f', error.status, delay)
                    elif isinstance(error, DownloadError) and (error.status in (401, 403) or any(s in str(error) for s in ('HTML', 'login'))):
                        progress.close()
                        page = renew()
                    retry.append((time.monotonic() + delay, row, attempt + 1))
                    logging.info('DOWNLOAD_NOVA_TENTATIVA_AGENDADA remoto=%s tentativa=%s espera_segundos=%.2f', path, attempt + 1, delay)
        progress.show(force=True)
        progress.close()
        return downloaded, skipped, failed
    except BaseException as error:
        logging.warning('LOTE_PARADA tipo=%s baixados=%s existentes=%s falhas=%s ativos=%s tentativas_pendentes=%s', type(error).__name__, downloaded, skipped, failed, len(jobs), len(retry))
        raise
    finally:
        progress.close()
        stop.set()
        for future in jobs:
            future.cancel()
        executor.shutdown(wait=True, cancel_futures=True)

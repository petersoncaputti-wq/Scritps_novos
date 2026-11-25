"""Workers HTTP; Playwright, progresso e SQLite ficam na thread principal."""
import logging
import queue
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import quote, urlsplit

import requests
from requests.adapters import HTTPAdapter

# Identificacao recomendada pela Microsoft para trafego automatizado no SharePoint.
USER_AGENT = 'NONISV|Ecorodovias|SharePointDownloader/1.1'
NETWORK_ERRORS = (requests.ConnectionError, requests.Timeout, requests.exceptions.ChunkedEncodingError)
COOKIE_REFRESH_SECONDS = 60
KEEPALIVE_SECONDS = 600  # leitura leve que mantem a sessao do navegador ativa
RENEW_MIN_INTERVAL = 60
RECOVERY_SUCCESSES = 20  # downloads sem limitacao antes de ampliar a concorrencia


class NetworkBreaker:
    """Pausa o lote quando a rede cai, sem consumir tentativas dos arquivos."""

    def __init__(self, threshold=5, first_wait=10, max_wait=120, clock=time.monotonic):
        self.threshold, self.first_wait, self.max_wait, self.clock = threshold, first_wait, max_wait, clock
        self.failures = 0
        self.open = False
        self.wait = first_wait
        self.next_probe = 0
        self.opened_at = 0

    def failure(self):
        self.failures += 1
        if not self.open and self.failures >= self.threshold:
            self.open = True
            self.wait = self.first_wait
            self.opened_at = self.clock()
            self.next_probe = self.opened_at + self.wait
            logging.warning('REDE_INDISPONIVEL falhas_consecutivas=%s pausa_segundos=%s', self.failures, self.wait)

    def success(self):
        if self.open:
            self.close('download_concluido')
        self.failures = 0

    def due(self):
        return self.open and self.clock() >= self.next_probe

    def probe_result(self, ok):
        if ok:
            self.close('teste_conexao')
        else:
            self.wait = min(self.max_wait, self.wait * 2)
            self.next_probe = self.clock() + self.wait
            logging.warning('REDE_AINDA_INDISPONIVEL proxima_verificacao_segundos=%s', self.wait)

    def close(self, reason):
        logging.info('REDE_RESTABELECIDA motivo=%s indisponivel_segundos=%.1f', reason, self.clock() - self.opened_at)
        self.open = False
        self.failures = 0


def network_probe(web):
    """Qualquer resposta HTTP (inclusive 401/403) prova que a rede voltou."""
    try:
        with requests.get('https://' + urlsplit(web).hostname + '/', timeout=(5, 10),
                          allow_redirects=False, headers={'User-Agent': USER_AGENT}, stream=True):
            return True
    except requests.RequestException:
        return False


def new_session(pool_size=2):
    session = requests.Session()
    session.mount('https://', HTTPAdapter(pool_connections=2, pool_maxsize=max(2, pool_size)))
    session.headers['User-Agent'] = USER_AGENT
    return session


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
            actual = transfer(session, endpoint, target, size, progress, resume_key=path + '|' + version, stop=stop)
            logging.info('DOWNLOAD_ROTA_SUCESSO remoto=%s local=%s rota=%s', path, target, route)
            return actual if actual is not None else size
        except DownloadError as error:
            if error.status not in (400, 404) or index == len(endpoints) - 1:
                raise
            # Only switch route on a rejected request, never hide auth, network,
            # integrity, or disk failures. Full size validation stays in transfer.
            logging.warning('DOWNLOAD_ROTA_ALTERNATIVA remoto=%s rota_recusada=%s http=%s proxima_rota=%s',
                            path, route, error.status, endpoints[index + 1][0])


def ensure_files_schema(db):
    """Colunas para pular a releitura: tamanho gravado e mtime apos verificacao."""
    columns = {row[1] for row in db.execute('PRAGMA table_info(files)')}
    for name in ('local_size', 'mtime_ns'):
        if name not in columns:
            db.execute(f'ALTER TABLE files ADD COLUMN {name} INTEGER')
    db.commit()


def save_file(db, path, text, size, version, status, target=None, local_size=None):
    mtime = None
    if target is not None:
        stat = Path(target).stat()
        local_size, mtime = stat.st_size, stat.st_mtime_ns
    db.execute('INSERT OR REPLACE INTO files (remote,local,size,version,status,local_size,mtime_ns) VALUES (?,?,?,?,?,?,?)',
               (path, text, size, version, status, local_size, mtime))
    db.commit()


def run_parallel(db, page, web, root, progress, renew, workers, transfer, verify_existing=False,
                 probe=None, breaker=None):
    from browser_bulk import DownloadError, safe_error
    probe = probe or (lambda: network_probe(web))
    breaker = breaker or NetworkBreaker()
    local = threading.local()
    sessions = []
    sessions_lock = threading.Lock()
    cookie_state = {'version': 0, 'cookies': [], 'key': None, 'read_at': None}
    successes_since_limit = 0
    keepalive_at = time.monotonic() + KEEPALIVE_SECONDS

    renewed_at = [float('-inf')]
    free_auth_retries = {}

    def renew_session(reason):
        """Uma renovacao para todo o lote; retorna True se os cookies mudaram."""
        nonlocal page
        if time.monotonic() - renewed_at[0] < RENEW_MIN_INTERVAL:
            return False  # sessao acabou de ser renovada: 403 repetido e falta de permissao
        renewed_at[0] = time.monotonic()
        before = cookie_state['version']
        progress.close()
        logging.warning('SESSAO_RENOVACAO motivo=%s', reason)
        page = renew()
        refresh_cookies(force=True)
        return cookie_state['version'] != before

    stop = threading.Event()
    updates = queue.Queue()
    active = {}
    jobs = {}
    retry = []
    downloaded = skipped = failed = 0
    limit = workers
    cooldown = 0
    ensure_files_schema(db)
    rows = iter(db.execute('SELECT remote,local,size,version FROM inventory ORDER BY remote'))
    exhausted = False
    last_log = time.monotonic()

    def refresh_cookies(force=False):
        """Le cookies do navegador na thread principal; workers recebem um snapshot."""
        read_at = cookie_state['read_at']
        if not force and read_at is not None and time.monotonic() - read_at < COOKIE_REFRESH_SECONDS:
            return
        cookies = page.context.cookies([web])
        key = sorted((c['name'], c['value'], c['domain'], c['path']) for c in cookies)
        cookie_state['read_at'] = time.monotonic()
        if key != cookie_state['key']:
            cookie_state.update(version=cookie_state['version'] + 1, cookies=cookies, key=key)

    def thread_session(cookies, version):
        # Uma sessao por thread reaproveita conexoes TCP/TLS entre arquivos.
        session = getattr(local, 'session', None)
        if session is None:
            session = local.session = new_session()
            local.version = None
            with sessions_lock:
                sessions.append(session)
        if local.version != version:
            session.cookies.clear()
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'],
                                    path=cookie['path'], secure=cookie.get('secure', True))
            local.version = version
        return session

    def worker(row, cookies, cookie_version):
        path, text, size, version = row
        session = thread_session(cookies, cookie_version)
        return transfer_with_fallback(session, web, path, Path(text), size,
                                      lambda n: updates.put((path, n)), version, stop, transfer)

    def prepare(row):
        nonlocal skipped, failed
        path, text, size, version = row
        target = Path(text)
        if not target.resolve().is_relative_to(root.resolve()):
            raise ValueError('Arquivo local fora do destino.')
        target.parent.mkdir(parents=True, exist_ok=True)
        record = db.execute('SELECT local,size,version,status,local_size,mtime_ns FROM files WHERE remote=?', (path,)).fetchone()
        if target.exists():
            same_version = not record or record[2] == version
            stat = target.stat() if target.is_file() else None
            known_size = record[4] if record and record[4] is not None else size
            # Office: o tamanho aceito no download (com propriedades) fica no registro.
            size_ok = stat is not None and stat.st_size in {size, known_size}
            if size_ok and same_version:
                # Ja verificado antes e intocado no disco: evita reler o arquivo inteiro.
                confirmed = (not verify_existing and record and record[3] in ('done', 'existing')
                             and stat.st_size == known_size
                             and (record[5] is None or record[5] == stat.st_mtime_ns))
                if confirmed:
                    how = 'registro_sqlite'
                else:
                    from bulk_transfer import verify_local_file
                    try:
                        verify_local_file(target, stat.st_size)
                    except DownloadError:
                        failed += 1
                        progress.failed = failed
                        save_file(db, path, text, size, version, 'failed')
                        progress.message(f'Falha na verificacao: {target.name} (detalhes no log)')
                        return False
                    how = 'leitura_completa'
                skipped += 1
                status = record[3] if confirmed else 'existing'
                if not confirmed or record[5] is None:
                    save_file(db, path, text, size, version, status, target)
                progress.message(f'Existente: {target.name}')
                logging.info('ARQUIVO_EXISTENTE remoto=%s local=%s bytes=%s verificacao=%s', path, text, stat.st_size, how)
                progress.complete(size)
            else:
                failed += 1
                progress.failed = failed
                progress.message(f'Conflito: {target.name} (detalhes no log)')
                logging.warning('ARQUIVO_CONFLITO remoto=%s local=%s esperado_bytes=%s local_bytes=%s motivo=tamanho_ou_versao', path, text, size, target.stat().st_size if target.is_file() else 'nao_e_arquivo')
            return False
        save_file(db, path, text, size, version, 'pending')
        return True

    executor = ThreadPoolExecutor(max_workers=workers)
    try:
        while not exhausted or jobs or retry:
            if breaker.due():
                progress.message('Testando a conexao com o SharePoint...')
                progress.show(sum(active.values()), force=True)
                breaker.probe_result(probe())
            if breaker.open:
                progress.message(f'Sem conexao; nova verificacao em {max(0, breaker.next_probe - time.monotonic()):.0f}s')
            else:
                if time.monotonic() >= keepalive_at:
                    keepalive_at = time.monotonic() + KEEPALIVE_SECONDS
                    from browser_test import is_auth_failure, read_json
                    try:
                        read_json(page, web + '/_api/web?$select=Url')
                        logging.info('SESSAO_KEEPALIVE ok=true')
                        refresh_cookies(force=True)
                    except RuntimeError as error:
                        if not is_auth_failure(error):
                            raise
                        renew_session('keepalive')
                refresh_cookies()
            while len(jobs) < limit and time.monotonic() >= cooldown and not breaker.open:
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
                progress.message(f'Baixando: {Path(row[1]).name}')
                logging.info('DOWNLOAD_INICIO remoto=%s local=%s bytes=%s tentativa=%s', row[0], row[1], row[2], attempt)
                jobs[executor.submit(worker, row, cookie_state['cookies'], cookie_state['version'])] = (
                    row, attempt, cookie_state['version'])
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
                row, attempt, job_cookies = jobs.pop(future)
                path, text, size, version = row
                active.pop(path, None)
                try:
                    actual = future.result()
                    save_file(db, path, text, size, version, 'done', target=text if Path(text).is_file() else None,
                              local_size=actual)
                    downloaded += 1
                    breaker.success()
                    if limit < workers:
                        successes_since_limit += 1
                        if successes_since_limit >= RECOVERY_SUCCESSES and time.monotonic() >= cooldown:
                            limit += 1
                            successes_since_limit = 0
                            logging.info('LOTE_CONCORRENCIA_AMPLIADA concorrencia=%s maximo=%s', limit, workers)
                    progress.complete(size)
                    progress.message(f'Concluido: {Path(text).name}')
                    logging.info('DOWNLOAD_CONCLUIDO remoto=%s local=%s bytes=%s tentativa=%s', path, text, size, attempt)
                except (requests.RequestException, RuntimeError, OSError) as error:
                    reason = safe_error(error)
                    progress.message(f'Nova tentativa: {Path(text).name} (detalhes no log)')
                    logging.warning('DOWNLOAD_ERRO remoto=%s local=%s tentativa=%s motivo=%s', path, text, attempt, reason)
                    if isinstance(error, NETWORK_ERRORS):
                        breaker.failure()
                        if breaker.open:
                            # Queda de rede: o arquivo volta para a fila sem gastar tentativa.
                            retry.append((0, row, attempt))
                            logging.info('DOWNLOAD_ADIADO_SEM_REDE remoto=%s tentativa=%s', path, attempt)
                            continue
                    if isinstance(error, DownloadError) and error.auth:
                        # Sessao expirada nao e culpa do arquivo: renova uma vez para o lote
                        # e repete sem gastar tentativa. Se os cookies nao mudarem, e permissao.
                        free = free_auth_retries.get(path, 0)
                        if free < 3 and (job_cookies != cookie_state['version']
                                         or renew_session(f'http_{error.status or "login"}')):
                            free_auth_retries[path] = free + 1
                            retry.append((0, row, attempt))
                            logging.info('DOWNLOAD_REPETIDO_APOS_RENOVACAO remoto=%s tentativa=%s', path, attempt)
                            continue
                    if attempt >= 3:
                        db.execute("UPDATE files SET status='failed' WHERE remote=?", (path,))
                        db.commit()
                        failed += 1
                        progress.failed = failed
                        progress.message(f'Falha: {Path(text).name} (detalhes no log)')
                        logging.error('DOWNLOAD_FALHA_FINAL remoto=%s local=%s tentativas=%s motivo=%s', path, text, attempt, reason)
                        continue
                    delay = getattr(error, 'retry_after', None)
                    delay = delay if delay is not None else attempt * 3
                    if isinstance(error, DownloadError) and error.status in (429, 503):
                        limit = 1
                        successes_since_limit = 0
                        cooldown = max(cooldown, time.monotonic() + delay)
                        progress.message(f'Limitado pelo SharePoint; pausa de {delay:.0f}s')
                        logging.warning('LOTE_LIMITADO http=%s concorrencia=1 espera_segundos=%.2f', error.status, delay)
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
        for session in sessions:
            session.close()

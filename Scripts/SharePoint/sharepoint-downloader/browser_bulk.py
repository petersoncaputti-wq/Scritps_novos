"""Download recursivo com streaming e registro persistente dos concluidos."""
import hashlib
import logging
import re
import sqlite3
import shutil
import time
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import quote, urljoin, urlsplit

import requests
from browser_test import folder_endpoint, read_json, wait_for_sharepoint


class DownloadError(RuntimeError):
    def __init__(self, message, status=None, retry_after=None):
        super().__init__(message)
        self.status = status
        self.retry_after = retry_after


def retry_delay(value):
    if not value:
        return None
    try:
        return max(0, float(value))
    except ValueError:
        try:
            return max(0, (parsedate_to_datetime(value) - datetime.now(timezone.utc)).total_seconds())
        except (ValueError, TypeError, OverflowError):
            return None


def safe_error(error):
    if isinstance(error, DownloadError):
        return str(error)
    if isinstance(error, requests.Timeout):
        return 'Tempo limite da conexao ou transferencia excedido.'
    if isinstance(error, requests.RequestException):
        return f'Falha de rede ({type(error).__name__}).'
    if isinstance(error, OSError):
        return f'Falha ao gravar o arquivo (errno={error.errno}, winerror={getattr(error, "winerror", None)}).'
    return f'Falha interna ({type(error).__name__}).'


def local_name(name: str) -> str:
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).rstrip(' .')
    if not clean or clean in ('.', '..'):
        clean = '_'
    if re.match(r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)', clean, re.I):
        clean = '_' + clean
    if clean != name or len(clean) > 120:
        suffix = Path(clean).suffix[:20]
        clean = clean[:80] + '_' + hashlib.sha256(name.encode()).hexdigest()[:12] + suffix
    return clean


def pages(page, endpoint: str):
    host = urlsplit(endpoint).hostname
    seen = set()
    while endpoint:
        if endpoint in seen or urlsplit(endpoint).hostname != host or urlsplit(endpoint).scheme != 'https':
            raise RuntimeError('Link de paginacao invalido.')
        seen.add(endpoint)
        data = read_json(page, endpoint)
        yield from data.get('value', [])
        next_url = data.get('@odata.nextLink') or data.get('odata.nextLink')
        endpoint = urljoin(endpoint, next_url) if next_url else None


def copy_cookies(page, session):
    session.cookies.clear()
    for cookie in page.context.cookies([page.url]):
        session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'],
                            path=cookie['path'], secure=cookie.get('secure', True))


def format_bytes(size):
    value = float(size)
    for unit in ('B', 'KiB', 'MiB', 'GiB', 'TiB'):
        if value < 1024 or unit == 'TiB':
            return f'{value:.1f} {unit}'
        value /= 1024


class Progress:
    def __init__(self, total_files, total_bytes):
        self.total_files = total_files
        self.total_bytes = total_bytes
        self.completed_files = self.completed_bytes = self.active_bytes = 0
        self.last_update = 0
        self.status = 'Preparando downloads'
        self.failed = 0
        self.visible = False

    def message(self, text):
        self.status = text.replace('\n', ' ').replace('\r', ' ')

    def close(self):
        if self.visible:
            print()
            self.visible = False

    def show(self, active_bytes=0, force=False):
        self.active_bytes = active_bytes
        now = time.monotonic()
        if not force and now - self.last_update < 0.25:
            return
        self.last_update = now
        current = self.completed_bytes + active_bytes
        ratio = (current / self.total_bytes if self.total_bytes else
                 self.completed_files / self.total_files if self.total_files else 1)
        ratio = min(1, max(0, ratio))
        width = max(10, shutil.get_terminal_size(fallback=(100, 24)).columns - 1)
        length = 20 if width >= 95 else 10
        filled = int(length * ratio)
        bar = '#' * filled + '-' * (length - filled)
        line = (f'[{bar}] {ratio:5.1%} | {self.completed_files}/{self.total_files} '
                f'| {format_bytes(current)}/{format_bytes(self.total_bytes)} | erros:{self.failed}')
        if len(line) + 5 < width:
            line += ' | ' + self.status
        print('\r' + line[:width].ljust(width), end='', flush=True)
        self.visible = True

    def complete(self, size):
        self.completed_files += 1
        self.completed_bytes += size
        self.show()


def transfer(session, url, target, size, progress=None, resume_key=None, stop=None):
    from bulk_transfer import transfer_file
    return transfer_file(session, url, target, size, progress, resume_key, stop)


def download_tree(page, web, remote_root, destination, login_url, workers=3, refresh_inventory=False):
    import msvcrt
    destination.mkdir(parents=True, exist_ok=True)
    with (destination / '.sharepoint_download.lock').open('a+b') as lock:
        if lock.tell() == 0:
            lock.write(b'0')
            lock.flush()
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            raise RuntimeError('Ja existe um lote rodando neste destino. Encerre-o antes de iniciar outro.') from None
        try:
            return _download_tree(page, web, remote_root, destination, login_url, workers, refresh_inventory)
        finally:
            lock.seek(0)
            msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)


def _download_tree(page, web, remote_root, destination, login_url, workers=3, refresh_inventory=False):
    destination.mkdir(parents=True, exist_ok=True)
    root_info = read_json(page, folder_endpoint(web, remote_root) + '?$select=Name,ServerRelativeUrl')
    remote_root = root_info['ServerRelativeUrl'].rstrip('/')
    root = destination if destination.name.casefold() == local_name(root_info['Name']).casefold() else destination / local_name(root_info['Name'])
    root.mkdir(exist_ok=True)
    key = hashlib.sha256((web + remote_root).encode()).hexdigest()[:16]
    db = sqlite3.connect(destination / ('.sharepoint_' + key + '.sqlite'))
    db.execute('CREATE TABLE IF NOT EXISTS files (remote TEXT PRIMARY KEY, local TEXT, size INTEGER, version TEXT, status TEXT)')
    db.execute('CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)')
    session = requests.Session()
    copy_cookies(page, session)
    downloaded = skipped = failed = 0
    stack = [(remote_root, root)]
    visited = set()
    hostname = urlsplit(web).hostname

    def renew():
        nonlocal page
        print('Refaca o login na janela Edge para continuar o lote.')
        logging.warning('LOGIN_RENOVACAO_INICIO hostname=%s', hostname)
        page.goto(login_url, wait_until='domcontentloaded', timeout=120000)
        page = wait_for_sharepoint(page.context, hostname)
        copy_cookies(page, session)
        logging.info('LOGIN_RENOVACAO_CONCLUIDA hostname=%s', hostname)
        return page

    def listing(endpoint):
        # Retry auth once; collecting one folder avoids duplicating a partial page.
        try:
            return list(pages(page, endpoint))
        except RuntimeError as error:
            if any(code in str(error) for code in ('HTTP 401', 'HTTP 403', 'HTTP 200')):
                renew()
                return list(pages(page, endpoint))
            raise

    print(f'Download completo iniciado. Destino: {root}')
    started = time.monotonic()
    logging.info('LOTE_INICIO origem=%s destino=%s workers=%s atualizar_inventario=%s', remote_root, root, workers, refresh_inventory)
    try:
        cached = db.execute("SELECT value FROM settings WHERE key='inventory_complete'").fetchone()
        if not refresh_inventory and cached and cached[0] == 'yes':
            print('Retomando o inventario salvo. Use --refresh-inventory para procurar alteracoes no SharePoint.')
            total_files, total_bytes = db.execute('SELECT COUNT(*), COALESCE(SUM(size),0) FROM inventory').fetchone()
            logging.info('INVENTARIO_REUTILIZADO arquivos=%s bytes=%s', total_files, total_bytes)
        else:
            db.execute("INSERT OR REPLACE INTO settings VALUES ('inventory_complete','no')")
            db.commit()
            db.execute('DROP TABLE IF EXISTS inventory')
            db.execute('CREATE TABLE inventory (remote TEXT PRIMARY KEY, local TEXT, size INTEGER, version TEXT)')
            total_files = total_bytes = 0
            print('Leitura inicial: contando todos os arquivos e subpastas...')
            logging.info('INVENTARIO_INICIO')
            while stack:
                remote, local = stack.pop()
                if remote in visited:
                    continue
                visited.add(remote)
                local.mkdir(parents=True, exist_ok=True)
                # Resolve rejects existing junctions/symlinks escaping the chosen root.
                if not local.resolve().is_relative_to(root.resolve()):
                    raise ValueError('Diretorio local fora do destino.')
                print(f'\nPasta: {remote}')
                logging.info('PASTA_LIDA remoto=%s local=%s', remote, local)
                endpoint = folder_endpoint(web, remote)
                children = listing(endpoint + '/Folders?$select=Name,ServerRelativeUrl&$top=200')
                for child in children:
                    path = child['ServerRelativeUrl']
                    if not path.startswith(remote + '/'):
                        raise RuntimeError('Subpasta fora da pasta solicitada.')
                    stack.append((path, local / local_name(child['Name'])))
                files = listing(endpoint + '/Files?$select=Name,Length,ServerRelativeUrl,TimeLastModified,ETag&$top=200')
                for item in files:
                    path = item['ServerRelativeUrl']
                    if not path.startswith(remote + '/'):
                        raise RuntimeError('Arquivo fora da pasta solicitada.')
                    target = local / local_name(item['Name'])
                    size = int(item['Length'])
                    version = item.get('ETag') or item.get('TimeLastModified', '')
                    db.execute('INSERT INTO inventory VALUES (?,?,?,?)', (path, str(target), size, version))
                    logging.info('ARQUIVO_INVENTARIADO remoto=%s local=%s bytes=%s versao=%s', path, target, size, version)
                    total_files += 1
                    total_bytes += size
                db.commit()
                print(f'Inventario: {len(visited)} pastas; {total_files} arquivos; {format_bytes(total_bytes)}')
            db.execute("INSERT OR REPLACE INTO settings VALUES ('inventory_complete','yes')")
            db.commit()
        print(f'Inventario pronto: {total_files} arquivos; {format_bytes(total_bytes)}.')
        logging.info('INVENTARIO_PRONTO arquivos=%s bytes=%s pastas_lidas=%s', total_files, total_bytes, len(visited))
        progress = Progress(total_files, total_bytes)
        from bulk_parallel import run_parallel
        downloaded, skipped, failed = run_parallel(
            db, page, web, root, progress, renew, workers, transfer)
        print(f'\nLote finalizado: {downloaded} baixados; {skipped} ja concluidos; {failed} falhas/conflitos.\nDestino: {root}')
        logging.info('LOTE_RESUMO estado=%s total_arquivos=%s total_bytes=%s baixados=%s existentes=%s falhas_conflitos=%s duracao_segundos=%.2f destino=%s', 'concluido' if not failed else 'concluido_com_falhas', total_files, total_bytes, downloaded, skipped, failed, time.monotonic()-started, root)
        return 1 if failed else 0
    except BaseException as error:
        logging.error('LOTE_INTERROMPIDO tipo=%s duracao_segundos=%.2f destino=%s', type(error).__name__, time.monotonic()-started, root)
        raise
    finally:
        db.close()
        session.close()

"""Downloader WebDAV Nextcloud. Python 3.10+, somente biblioteca padrao.

Adapta o fluxo streaming/.part/Range/If-Range do downloader SharePoint.
"""
import argparse
import base64
import getpass
import hashlib
import json
import logging
import os
import re
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import quote, unquote, urljoin, urlsplit

DEFAULT_URL = 'https://cloud.engemap.com.br/s/SmfDfn34BFy25Mo'
DAV = '{DAV:}'


class DownloadError(Exception):
    pass


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise DownloadError('Redirecionamento recusado. Verifique o link e a autenticacao.')


def local_name(name):
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).rstrip(' .')
    if not clean or clean in ('.', '..'):
        clean = '_'
    if re.match(r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)', clean, re.I):
        clean = '_' + clean
    # Preservar nomes validos mantem os vinculos TIFF/TFW e SHP/DBF/SHX.
    if clean == name and len(name) <= 150:
        return name
    suffix = hashlib.sha256(name.encode()).hexdigest()[:12]
    stem, ext = os.path.splitext(clean)
    return stem[:90] + '_' + suffix + ext[:20]


def validate_names(files, folders):
    """Recusar colisoes no Windows antes de criar arquivos."""
    seen = {}
    for relative in folders + [item['relative'] for item in files]:
        parts = relative.split('/')
        for index in range(1, len(parts) + 1):
            original = '/'.join(parts[:index])
            local = '/'.join(local_name(part) for part in parts[:index]).casefold()
            if local in seen and seen[local] != original:
                raise DownloadError(f'Nomes remotos colidem no Windows: {seen[local]} e {original}.')
            seen[local] = original


def save_json(path, data):
    temp = path.with_name(path.name + '.tmp')
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    temp.replace(path)


def safe_path(root, relative):
    path = root
    for part in relative.split('/'):
        if part in ('', '.', '..'):
            raise DownloadError('Caminho remoto invalido.')
        path = path / local_name(part)
        if path.is_symlink() or (os.name == 'nt' and path.exists()
                                and path.stat().st_file_attributes & 1024):
            raise DownloadError('Destino contem link/junction.')
    if not path.resolve().is_relative_to(root.resolve()):
        raise DownloadError('Caminho fora do destino.')
    return path


class Client:
    def __init__(self, share_url, password):
        parsed = urlsplit(share_url.strip())
        match = re.fullmatch(r'/s/([A-Za-z0-9]+)/?', parsed.path)
        if (parsed.scheme != 'https' or not parsed.hostname or parsed.username
                or parsed.password or parsed.query or parsed.fragment or not match):
            raise DownloadError('Informe um link HTTPS no formato https://servidor/s/token.')
        self.origin = f'https://{parsed.netloc}'
        token = match[1]
        self.roots = [f'/public.php/dav/files/{token}/', '/public.php/webdav/']
        self.root = self.roots[0]
        self.auth = 'Basic ' + base64.b64encode(f'{token}:{password}'.encode()).decode()
        self.opener = urllib.request.build_opener(NoRedirect())

    def request(self, path, method='GET', headers=None, data=None):
        headers = dict(headers or {})
        headers.update({'Authorization': self.auth, 'Accept-Encoding': 'identity'})
        request = urllib.request.Request(self.origin + path, data=data, headers=headers, method=method)
        return self.opener.open(request, timeout=90)

    def listing(self, path):
        body = (b'<?xml version="1.0"?><d:propfind xmlns:d="DAV:"><d:prop>'
                b'<d:resourcetype/><d:getcontentlength/><d:getetag/>'
                b'</d:prop></d:propfind>')
        with self.request(path, 'PROPFIND', {'Depth': '1', 'Content-Type': 'application/xml'}, body) as response:
            if response.status != 207:
                raise DownloadError('Servidor nao retornou listagem WebDAV (207).')
            xml = ET.fromstring(response.read())
        items = []
        for entry in xml.findall(DAV + 'response'):
            href = urlsplit(urljoin(self.origin + path, entry.findtext(DAV + 'href', '')))
            if href.netloc != urlsplit(self.origin).netloc:
                raise DownloadError('Listagem aponta para outro servidor.')
            remote = unquote(href.path).rstrip('/')
            current = unquote(path).rstrip('/')
            if remote == current:
                continue
            if not remote.startswith(current + '/'):
                raise DownloadError('Listagem fora da pasta solicitada.')
            name = remote[len(current) + 1:]
            if '/' in name or name in ('.', '..') or '\\' in name:
                raise DownloadError('Nome remoto invalido.')
            prop = None
            for group in entry.findall(DAV + 'propstat'):
                if ' 200 ' in group.findtext(DAV + 'status', ''):
                    prop = group.find(DAV + 'prop')
                    break
            if prop is None:
                raise DownloadError('Sem permissao para ler metadados de um item.')
            folder = prop.find(DAV + 'resourcetype/' + DAV + 'collection') is not None
            size = None if folder else int(prop.findtext(DAV + 'getcontentlength', '-1'))
            if not folder and size < 0:
                raise DownloadError('Arquivo sem tamanho no inventario.')
            items.append({'name': name, 'path': quote(remote, safe='/'), 'folder': folder,
                          'size': size, 'etag': prop.findtext(DAV + 'getetag')})
        return items

    def connect(self):
        for index, root in enumerate(self.roots):
            self.root = root
            try:
                return self.listing(root)
            except urllib.error.HTTPError as error:
                if index == 0 and error.code in (401, 403, 404, 405):
                    continue
                if error.code in (401, 403):
                    raise DownloadError('Acesso recusado. Confira a senha e a permissao de download/WebDAV.') from None
                raise

    def inventory(self):
        pending = [('', self.connect())]
        files, folders = [], []
        while pending:
            parent, entries = pending.pop()
            for item in entries:
                relative = parent + item['name']
                if item['folder']:
                    folders.append(relative)
                    pending.append((relative + '/', self.listing(item['path'] + '/')))
                else:
                    files.append(dict(item, relative=relative))
        return files, folders


def transfer(client, item, target):
    partial = target.with_name(target.name + '.engemap.part')
    meta = partial.with_name(partial.name + '.json')
    for path in (target, partial, meta):
        if path.is_symlink():
            raise DownloadError('Arquivo local e um link simbolico.')
    identity = {'source': client.origin + item['path'], 'size': item['size'], 'etag': item['etag']}
    saved = json.loads(meta.read_text(encoding='utf-8')) if meta.exists() else {}
    if target.exists():
        raise DownloadError('Arquivo existente sem confirmacao no estado; destino preservado.')
    offset = partial.stat().st_size if partial.exists() else 0
    if offset and not saved:
        raise DownloadError('Parcial sem identificacao; mova-o antes de repetir.')
    etag = item['etag']
    if saved != identity or not etag or etag.startswith('W/') or offset >= item['size']:
        offset = 0
    headers = {'If-Match': etag} if etag else {}
    if offset:
        headers.update({'Range': f'bytes={offset}-', 'If-Range': etag})
    with client.request(item['path'], headers=headers) as response:
        if response.status not in (200, 206):
            raise DownloadError(f'Download retornou HTTP {response.status}.')
        if etag and response.headers.get('ETag') != etag:
            raise DownloadError('Versao remota mudou ou servidor omitiu ETag. Refaca o inventario.')
        if response.status == 206:
            expected = f"bytes {offset}-{item['size'] - 1}/{item['size']}"
            if not offset or response.headers.get('Content-Range') != expected:
                raise DownloadError('Resposta de retomada invalida; parcial preservado.')
        else:
            offset = 0
        save_json(meta, identity)
        written = offset
        with partial.open('ab' if offset else 'wb') as output:
            while chunk := response.read(1024 * 1024):
                output.write(chunk)
                written += len(chunk)
                if written > item['size']:
                    raise DownloadError('Download excedeu tamanho esperado.')
            output.flush()
            os.fsync(output.fileno())
    if written != item['size'] or partial.stat().st_size != item['size']:
        raise DownloadError('Tamanho recebido diverge do inventario; parcial preservado.')
    # Digest local permite detectar alteracao posterior antes de pular um arquivo.
    digest = file_hash(partial)
    if target.exists():
        raise DownloadError('Destino criado durante a transferencia.')
    partial.rename(target)
    meta.unlink(missing_ok=True)
    return dict(identity, sha256=digest)


def file_hash(path):
    digest = hashlib.sha256()
    with path.open('rb') as source:
        while chunk := source.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def main():
    parser = argparse.ArgumentParser(description='Download recursivo de compartilhamento Engemap/Nextcloud.')
    parser.add_argument('--url', default=DEFAULT_URL)
    parser.add_argument('--destino', type=Path)
    parser.add_argument('--listar', action='store_true', help='Apenas inventariar; nao baixar arquivos.')
    parser.add_argument('--sem-senha', action='store_true')
    args = parser.parse_args()
    password = '' if args.sem_senha else getpass.getpass('Senha do compartilhamento (nao sera salva): ')
    client = Client(args.url, password)
    print('Consultando pastas e arquivos...')
    files, folders = client.inventory()
    validate_names(files, folders)
    print(f'{len(files)} arquivos; {len(folders)} pastas; {sum(i["size"] for i in files):,} bytes.')
    if args.listar:
        for item in files:
            print(f'{item["size"]:>14,}  {item["relative"]}')
        return 0
    if not args.destino:
        destination = input('Pasta de destino: ').strip().strip('"')
        if not destination:
            raise DownloadError('Informe uma pasta de destino.')
        args.destino = Path(destination)
    root = args.destino.resolve()
    root.mkdir(parents=True, exist_ok=True)
    # Uma execucao por destino; trava removida mesmo em Ctrl+C.
    lock = root / '.engemap.lock'
    with lock.open('x'):
        pass
    try:
        logging.basicConfig(filename=root / 'engemap.log', level=logging.INFO,
                            format='%(asctime)s %(levelname)s %(message)s', encoding='utf-8')
        state_path = root / 'engemap-estado.json'
        state = json.loads(state_path.read_text(encoding='utf-8')) if state_path.exists() else {}
        save_json(root / 'engemap-inventario.json', {'files': files, 'folders': folders})
        for folder in folders:
            safe_path(root, folder).mkdir(parents=True, exist_ok=True)
        failures = 0
        for number, item in enumerate(files, 1):
            print(f'[{number}/{len(files)}] {item["relative"]}', flush=True)
            target = safe_path(root, item['relative'])
            target.parent.mkdir(parents=True, exist_ok=True)
            identity = {'source': client.origin + item['path'], 'size': item['size'], 'etag': item['etag']}
            previous = state.get(item['relative'], {})
            if (item['etag'] and all(previous.get(k) == v for k, v in identity.items())
                    and target.is_file() and target.stat().st_size == item['size']
                    and previous.get('sha256') == file_hash(target)):
                print('  Ja baixado e conferido.')
                continue
            for attempt in range(3):
                try:
                    state[item['relative']] = transfer(client, item, target)
                    save_json(state_path, state)
                    logging.info('Concluido: %s', item['relative'])
                    break
                except (OSError, DownloadError, ValueError) as error:
                    # Nunca registrar URL autenticada ou senha.
                    logging.warning('Falha: %s; tipo=%s', item['relative'], type(error).__name__)
                    if attempt == 2 or isinstance(error, (DownloadError, ValueError)):
                        print(f'  FALHA: {error}')
                        failures += 1
                        break
                    time.sleep(2 ** attempt)
        print(f'Concluido: {len(files) - failures} arquivos conferidos; {failures} falhas. Destino: {root}')
        return 1 if failures else 0
    finally:
        lock.unlink(missing_ok=True)


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print('\nInterrompido. Parciais preservados para a proxima execucao.')
        raise SystemExit(130)
    except (OSError, DownloadError, ValueError, ET.ParseError) as error:
        print(f'ERRO: {error}')
        raise SystemExit(1)

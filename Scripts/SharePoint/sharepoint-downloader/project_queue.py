"""Fila de projetos XLSX; leitura sem executar formulas ou alterar a planilha."""
import hashlib
import json
import logging
import posixpath
import zipfile
import uuid
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit

from sharepoint import parse_sharepoint_url

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


@dataclass
class Project:
    row: int
    priority: int
    year: str
    lot: str
    sh: str
    url: str

    @property
    def folder(self):
        from browser_bulk import local_name
        identity = hashlib.sha256(self.url.encode()).hexdigest()[:12]
        return local_name(f'Ano_{self.year}_Lote_{self.lot}_SH_{self.sh}_{identity}')


def normalize_url(url):
    parts = urlsplit(url.strip())
    if parts.path.startswith('/:f:/r/'):
        # These links contain the actual folder path, unlike opaque share links.
        url = urlunsplit((parts.scheme, parts.netloc, parts.path[len('/:f:/r'):], '', ''))
    parsed = parse_sharepoint_url(url)
    from urllib.parse import quote
    return 'https://' + parsed.hostname + quote(parsed.path, safe='/')


def load_projects(path, sheet=None, priorities=None, rows=None):
    """Read E priority and F hyperlink; reject invalid project rows visibly."""
    with zipfile.ZipFile(path) as book:
        strings = []
        if 'xl/sharedStrings.xml' in book.namelist():
            strings = [''.join(t.text or '' for t in si.findall('.//s:t', NS))
                       for si in ET.fromstring(book.read('xl/sharedStrings.xml')).findall('s:si', NS)]
        sheets = ET.fromstring(book.read('xl/workbook.xml')).find('s:sheets', NS)
        selected = next((s for s in sheets if sheet is None or s.attrib['name'] == sheet), None)
        if selected is None:
            raise ValueError('Aba nao encontrada na planilha.')
        rels = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(book.read('xl/_rels/workbook.xml.rels'))}
        target = rels[selected.attrib[RID]]
        name = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        xml = ET.fromstring(book.read(name))
        relpath = posixpath.dirname(name) + '/_rels/' + posixpath.basename(name) + '.rels'
        links = {}
        if relpath in book.namelist():
            links = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(book.read(relpath))
                     if r.attrib.get('TargetMode') == 'External'}
        hyperlinks = {h.attrib['ref']: links.get(h.attrib.get(RID), '')
                      for h in xml.findall('s:hyperlinks/s:hyperlink', NS)}
        projects, issues, seen = [], [], set()
        for row in xml.findall('s:sheetData/s:row', NS):
            number = int(row.attrib['r'])
            cells = {}
            for cell in row:
                col = ''.join(c for c in cell.attrib['r'] if c.isalpha())
                value = cell.findtext('s:v', default='', namespaces=NS)
                if cell.attrib.get('t') == 's':
                    value = strings[int(value)] if value else ''
                elif cell.attrib.get('t') == 'inlineStr':
                    value = ''.join(t.text or '' for t in cell.findall('.//s:t', NS))
                cells[col] = value.strip()
            if cells.get('E', '').casefold() == 'prioridade':
                continue
            # Summary/notes are not projects. Missing priority on a project is an issue.
            if not (cells.get('C') or cells.get('E') or cells.get('F')):
                continue
            if rows and number not in rows:
                continue
            try:
                priority = int(cells.get('E', ''))
                if priority < 1:
                    raise ValueError
            except ValueError:
                issues.append({'linha': number, 'motivo': 'Prioridade ausente ou invalida.'})
                continue
            if priorities and priority not in priorities:
                continue
            try:
                url = normalize_url(hyperlinks.get('F' + str(number)) or cells.get('F', ''))
            except ValueError:
                issues.append({'linha': number, 'motivo': 'Link ausente ou nao suportado na coluna F.'})
                continue
            if url in seen:
                issues.append({'linha': number, 'motivo': 'Pasta duplicada na fila.'})
                continue
            seen.add(url)
            projects.append(Project(number, priority, cells.get('A', ''), cells.get('B', ''), cells.get('C', ''), url))
        projects.sort(key=lambda p: (p.priority, p.row))
        return projects, issues


def resolve_web(page, parsed):
    from browser_test import read_json
    from urllib.parse import quote
    segments = parsed.path.strip('/').split('/')
    for count in range(len(segments), -1, -1):
        candidate = 'https://' + parsed.hostname + ('/' + '/'.join(quote(s, safe='') for s in segments[:count]) if count else '')
        try:
            info = read_json(page, candidate + '/_api/web?$select=Url,Title')
            web = info['Url'].rstrip('/')
            if urlsplit(web).hostname != parsed.hostname:
                raise RuntimeError('Site retornado fora do dominio esperado.')
            return web
        except RuntimeError as error:
            if 'HTTP 404' not in str(error):
                raise
    raise RuntimeError('Nao foi possivel localizar o site SharePoint.')


def execute_queue(projects, destination, workers=3, refresh_inventory=False, issues=None):
    from playwright.sync_api import sync_playwright
    from browser_test import wait_for_sharepoint
    from browser_bulk import download_tree, safe_error
    destination = Path(destination).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    report = {'pendencias_planilha': issues or [], 'projetos': []}
    report_path = destination / ('fila_projetos_resumo_' + uuid.uuid4().hex[:12] + '.json')

    def save():
        temporary = report_path.with_suffix('.json.tmp')
        temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        temporary.replace(report_path)

    save()
    with sync_playwright() as pw:
        browser = pw.chromium.launch(channel='msedge', headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        authenticated_host = None
        try:
            for index, project in enumerate(projects, 1):
                target = destination / project.folder
                record = {'linha': project.row, 'prioridade': project.priority,
                          'destino': str(target), 'estado': 'em_andamento'}
                report['projetos'].append(record)
                save()
                print(f'\nProjeto {index}/{len(projects)}: linha {project.row}, prioridade {project.priority}')
                logging.info('FILA_PROJETO_INICIO linha=%s prioridade=%s', project.row, project.priority)
                try:
                    parsed = parse_sharepoint_url(project.url)
                    page.goto(project.url, wait_until='domcontentloaded', timeout=120000)
                    if authenticated_host != parsed.hostname or urlsplit(page.url).hostname != parsed.hostname:
                        print('Conclua o login/MFA no Edge para iniciar ou renovar a sessao.')
                        page = wait_for_sharepoint(context, parsed.hostname)
                        authenticated_host = parsed.hostname
                    try:
                        web = resolve_web(page, parsed)
                    except RuntimeError as error:
                        if not any(code in str(error) for code in ('HTTP 401', 'HTTP 403', 'HTTP 200')):
                            raise
                        print('Confira a sessao e o acesso ao projeto no Edge.')
                        page = wait_for_sharepoint(context, parsed.hostname)
                        web = resolve_web(page, parsed)
                    result = download_tree(page, web, parsed.path, target, project.url,
                                           workers=workers, refresh_inventory=refresh_inventory)
                    record['estado'] = 'concluido' if result == 0 else 'com_pendencias'
                except (KeyboardInterrupt, EOFError):
                    record['estado'] = 'interrompido'
                    save()
                    raise
                except Exception as error:
                    record['estado'] = 'falha'
                    record['erro'] = safe_error(error)
                    print(f"Projeto com falha: {record['erro']}. Seguindo para o proximo.")
                save()
                logging.info('FILA_PROJETO_FIM linha=%s estado=%s', project.row, record['estado'])
        finally:
            context.close()
            browser.close()
    failed = sum(r['estado'] != 'concluido' for r in report['projetos'])
    print(f'\nFila finalizada: {len(projects)-failed} projetos concluidos; {failed} com pendencias. Resumo: {report_path}')
    return 1 if failed or issues else 0


def run_project_queue(path, destination='', sheet=None, priorities=None, rows=None,
                      plan_only=False, workers=3, refresh_inventory=False):
    projects, issues = load_projects(path, sheet, priorities, rows)
    for issue in issues:
        print(f"Pendencia na linha {issue['linha']}: {issue['motivo']}")
    print(f'Fila: {len(projects)} projetos validos; {len(issues)} pendencias na planilha.')
    for project in projects:
        print(f'Prioridade {project.priority}: linha {project.row} -> {project.folder}')
    if plan_only:
        return 1 if issues else 0
    if not projects:
        raise ValueError('Nenhum projeto valido selecionado para download.')
    if not destination:
        from browser_test import select_destination
        destination = select_destination()
    return execute_queue(projects, destination, workers, refresh_inventory, issues)

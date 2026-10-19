import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

from project_queue import Project, execute_queue, load_projects, normalize_url


class QueueTests(unittest.TestCase):
    def test_link_formats(self):
        self.assertEqual(normalize_url('https://a.sharepoint.com/:f:/r/sites/P/Docs/100%25?e=secret'),
                         'https://a.sharepoint.com/sites/P/Docs/100%25')
        self.assertEqual(normalize_url('https://a.sharepoint.com/sites/P/Docs/Forms/AllItems.aspx?id=%2Fsites%2FP%2FDocs%2FA'),
                         'https://a.sharepoint.com/sites/P/Docs/A')
        for url in ['L17', 'https://a.sharepoint.com/:f:/s/P/opaque', 'https://evil.test/:f:/r/a']:
            with self.assertRaises(ValueError):
                normalize_url(url)

    def test_reader_order_hyperlink_filters_and_issues(self):
        def cell(ref, value):
            return f'<c r="{ref}" t="inlineStr"><is><t>{value}</t></is></c>'
        rows = []
        for row, priority, label in [(1, 'Prioridade', 'Link'), (2, '2', 'display'),
                                     (3, '1', 'https://a.sharepoint.com/sites/P/B'),
                                     (4, '1', 'https://a.sharepoint.com/sites/P/C'), (5, '3', 'L17')]:
            rows.append(f'<row r="{row}">' + cell('E'+str(row), priority) + cell('F'+str(row), label) + '</row>')
        rows.append('<row r="6">' + cell('A6', 'Total') + cell('D6', '1 TB') + '</row>')
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'input.xlsx'
            with zipfile.ZipFile(path, 'w') as z:
                z.writestr('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Lista" r:id="r1"/></sheets></workbook>')
                z.writestr('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>')
                z.writestr('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>'+''.join(rows)+'</sheetData><hyperlinks><hyperlink ref="F2" r:id="h1"/></hyperlinks></worksheet>')
                z.writestr('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="h1" TargetMode="External" Target="https://a.sharepoint.com/:f:/r/sites/P/A?e=secret"/></Relationships>')
            projects, issues = load_projects(path)
            self.assertEqual([p.row for p in projects], [3, 4, 2])
            self.assertEqual([i['linha'] for i in issues], [5])
            self.assertEqual(projects[-1].url, 'https://a.sharepoint.com/sites/P/A')
            self.assertEqual([p.row for p in load_projects(path, priorities=[1], rows=[4])[0]], [4])
            with self.assertRaises(ValueError):
                load_projects(path, sheet='Missing')

    def test_single_session_failure_continuation_and_summary(self):
        import json
        projects = [Project(i, i, '4', '1', str(i), f'https://a.sharepoint.com/sites/P/{i}') for i in range(1, 4)]
        manager = MagicMock()
        pw = manager.return_value.__enter__.return_value
        browser = pw.chromium.launch.return_value
        context = browser.new_context.return_value
        page = context.new_page.return_value
        page.url = 'https://a.sharepoint.com/sites/P'
        with tempfile.TemporaryDirectory() as tmp, \
             patch('playwright.sync_api.sync_playwright', manager), \
             patch('browser_test.wait_for_sharepoint', return_value=page) as login, \
             patch('project_queue.resolve_web', return_value='https://a.sharepoint.com/sites/P'), \
             patch('browser_bulk.download_tree', side_effect=[RuntimeError('failure'), 1, 0]) as download:
            self.assertEqual(execute_queue(projects, tmp), 1)
            self.assertEqual(login.call_count, 1)
            self.assertEqual(download.call_count, 3)
            self.assertEqual(browser.new_context.call_count, 1)
            report = json.loads(next(Path(tmp).glob('fila_projetos_resumo_*.json')).read_text(encoding='utf-8'))
            self.assertEqual([r['estado'] for r in report['projetos']], ['falha', 'com_pendencias', 'concluido'])
            context.close.assert_called_once()


if __name__ == '__main__':
    unittest.main()

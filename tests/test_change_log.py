"""Workbook-level change history without replacing the existing native objects."""
from datetime import datetime, timedelta, timezone
import io
from pathlib import Path
import sys
import unittest
import xml.etree.ElementTree as ET
import zipfile

CLOUD_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(CLOUD_ROOT))
from change_log import append_change_log
from test_cloud import minimal_xlsx, zip_bytes

M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
COMPLETED = datetime(2026, 11, 2, 10, 30, 0, tzinfo=timezone(timedelta(hours=8)))
CHANGE = {'month': 9, 'unit': '测试场站', 'cell': 'BH3', 'field': '总结算电费',
          'old': None, 'new': 0, 'source': '9月原始结算单.xlsx'}


def parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {item.filename: archive.read(item) for item in archive.infolist()}


def sheets(data):
    import posixpath
    files = parts(data)
    rels = {item.attrib['Id']: item.attrib['Target']
            for item in ET.fromstring(files['xl/_rels/workbook.xml.rels'])}
    result = {}
    for item in ET.fromstring(files['xl/workbook.xml']).find(M + 'sheets'):
        target = rels[item.attrib[R + 'id']]
        path = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        result[item.attrib['name']] = (path, ET.fromstring(files[path]))
    return result


def row_values(row):
    values = []
    for cell in row.findall(M + 'c'):
        if cell.attrib.get('t') == 'inlineStr':
            values.append(''.join(cell.find(M + 'is').itertext()))
        else:
            value = cell.find(M + 'v')
            values.append(None if value is None else value.text)
    return values


def log_data_rows(tree):
    rows = tree.find(M + 'sheetData').findall(M + 'row')
    header = next(i for i, row in enumerate(rows)
                  if '更新前' in row_values(row) and '更新后' in row_values(row))
    return rows[header + 1:]


class EmbeddedChangeLogTests(unittest.TestCase):
    def test_first_export_preserves_original_parts_and_typed_values(self):
        original = minimal_xlsx()
        result = append_change_log(original, [CHANGE], COMPLETED)
        before, after = parts(original), parts(result)
        mutable = {'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml', 'xl/styles.xml'}
        for name, data in before.items():
            if name not in mutable:
                self.assertEqual(after[name], data, name)
        tree = sheets(result)['变更明细'][1]
        rows = log_data_rows(tree)
        self.assertEqual(len(rows), 1)
        cells = {cell.attrib['r'].rstrip('0123456789'): cell for cell in rows[0].findall(M + 'c')}
        self.assertNotEqual(cells['A'].attrib.get('t'), 'inlineStr')
        self.assertIsNotNone(cells['A'].find(M + 'v'))
        if 'F' in cells:
            self.assertIsNone(cells['F'].find(M + 'v'))
        self.assertEqual(float(cells['G'].find(M + 'v').text), 0)
        self.assertIsNotNone(tree.find(M + 'autoFilter'))
        self.assertIsNotNone(tree.find('.//' + M + 'pane'))

    def test_later_update_appends_and_unchanged_run_keeps_history(self):
        first = append_change_log(minimal_xlsx(), [CHANGE], COMPLETED)
        second_change = dict(CHANGE, month=10, old=0, new=-123.456789)
        second = append_change_log(first, [second_change], COMPLETED + timedelta(days=30))
        self.assertEqual(list(sheets(second)), ['数据表', '变更明细'])
        first_rows = log_data_rows(sheets(first)['变更明细'][1])
        second_rows = log_data_rows(sheets(second)['变更明细'][1])
        self.assertEqual(len(second_rows), 2)
        self.assertEqual(row_values(first_rows[0]), row_values(second_rows[0]))
        self.assertEqual(parts(first)['xl/styles.xml'], parts(second)['xl/styles.xml'])
        self.assertEqual(append_change_log(second, [], COMPLETED), second)

    def test_existing_user_sheet_name_is_preserved(self):
        original_parts = parts(minimal_xlsx())
        original_parts['xl/workbook.xml'] = original_parts['xl/workbook.xml'].replace('数据表'.encode(), '变更明细'.encode())
        original = zip_bytes(list(original_parts.items()))
        result = append_change_log(original, [CHANGE], COMPLETED)
        result_sheets = sheets(result)
        self.assertEqual(len(result_sheets), 2)
        self.assertEqual(parts(result)[result_sheets['变更明细'][0]], original_parts['xl/worksheets/sheet1.xml'])
        again = append_change_log(result, [CHANGE], COMPLETED)
        self.assertEqual(len(sheets(again)), 2)

    def test_formula_like_text_is_stored_as_text(self):
        change = dict(CHANGE, old='=SUM(A1:A2)', new='@text', source='原单<&>.xlsx')
        result = append_change_log(minimal_xlsx(), [change], COMPLETED)
        row = log_data_rows(sheets(result)['变更明细'][1])[0]
        self.assertFalse(row.findall('.//' + M + 'f'))
        self.assertIn('=SUM(A1:A2)', row_values(row))
        self.assertIn('原单<&>.xlsx', row_values(row))

    def test_date_1904_uses_matching_excel_serial(self):
        files = parts(minimal_xlsx())
        files['xl/workbook.xml'] = files['xl/workbook.xml'].replace(b'<sheets>', b'<workbookPr date1904="1"/><sheets>')
        result = append_change_log(zip_bytes(list(files.items())), [CHANGE], COMPLETED)
        row = log_data_rows(sheets(result)['变更明细'][1])[0]
        value = float(row.find(M + 'c/' + M + 'v').text)
        expected = (COMPLETED.replace(tzinfo=None) - datetime(1904, 1, 1)).total_seconds() / 86400
        self.assertAlmostEqual(value, expected, places=8)


if __name__ == '__main__':
    unittest.main()

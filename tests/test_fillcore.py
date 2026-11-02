import copy
import json
from pathlib import Path
import tempfile
import unittest

import openpyxl
import fillcore

ROOT = Path(__file__).resolve().parents[1]


class FillCoreTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.profile = json.loads((ROOT / 'examples/demo_rule.json').read_text(encoding='utf-8'))
        self.profile['output_dir'] = self.temp.name
        self.profile['output_name'] = 'result.xlsx'

    def run_profile(self):
        return Path(fillcore.run_profile(self.profile, str(ROOT)))

    def test_example_preserves_footer_and_numeric_values(self):
        path = self.run_profile()
        wb = openpyxl.load_workbook(path)
        self.addCleanup(wb.close)
        ws = wb['示例表']
        self.assertEqual(ws['B20'].value, 'DEMO0018')
        self.assertEqual(ws['E3'].value, '盒')
        self.assertIsInstance(ws['G3'].value, (int, float))
        self.assertIn('备注', ws['A22'].value)
        self.assertEqual(ws['A23'].value, '制表人：')
        self.assertIn('A22:I22', {str(m) for m in ws.merged_cells.ranges})
        self.assertEqual(ws.auto_filter.ref, 'A2:I20')

    def test_literal_filename_uses_source_directory(self):
        self.profile['tasks'][0]['source']['pattern'] = 'source.xlsx'
        self.assertTrue(self.run_profile().is_file())

    def test_failure_leaves_existing_output_untouched(self):
        path = Path(self.temp.name) / 'result.xlsx'
        path.write_bytes(b'previous result')
        self.profile['tasks'][0]['source']['pattern'] = 'missing.xlsx'
        with self.assertRaises(FileNotFoundError):
            self.run_profile()
        self.assertEqual(path.read_bytes(), b'previous result')
        self.assertEqual([p.name for p in Path(self.temp.name).iterdir()], ['result.xlsx'])

    def test_failure_does_not_create_output(self):
        self.profile['tasks'][0]['columns'][0]['source'] = 'missing column'
        with self.assertRaises(ValueError):
            self.run_profile()
        self.assertFalse((Path(self.temp.name) / 'result.xlsx').exists())

    def test_output_cannot_replace_template_or_source(self):
        for path in ('examples/template.xlsx', 'examples/source.xlsx'):
            with self.subTest(path=path):
                self.profile['output_dir'] = str(ROOT / 'examples')
                self.profile['output_name'] = Path(path).name
                before = (ROOT / path).read_bytes()
                with self.assertRaises(ValueError):
                    self.run_profile()
                self.assertEqual((ROOT / path).read_bytes(), before)

    def test_old_output_is_backed_up_after_success(self):
        path = self.run_profile()
        before = path.read_bytes()
        self.run_profile()
        backups = list(path.parent.glob('result.旧文件-*.xlsx'))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), before)

    def test_shorter_data_clears_old_mapped_cells_and_keeps_text_literal(self):
        wb = openpyxl.Workbook()
        self.addCleanup(wb.close)
        ws = wb.active
        ws.append(['名称', '数量'])
        ws.append(['old', 10])
        ws.append(['stale', 20])
        task = {'target': {'header_row': 1, 'data_start': 2}, 'columns': [
            {'target': '名称', 'source': 'name'}, {'target': '数量', 'source': 'qty'}]}
        fillcore.fill_sheet(ws, task, [{'name': '=1+1', 'qty': 2}], {})
        self.assertEqual(ws['A2'].data_type, 's')
        self.assertEqual(ws['B2'].value, 2)
        self.assertIsNone(ws['A3'].value)
        self.assertIsNone(ws['B3'].value)

    def test_invalid_data_start_rejected(self):
        wb = openpyxl.Workbook()
        self.addCleanup(wb.close)
        with self.assertRaises(ValueError):
            fillcore.fill_sheet(wb.active, {'target': {'header_row': 2, 'data_start': 1}}, [], {})


if __name__ == '__main__':
    unittest.main()

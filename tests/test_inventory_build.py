"""驗證資料管理到獨立 HTML 的邊界與失敗保護，不依賴正式資料。"""
import hashlib
import json
from pathlib import Path
import sqlite3
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook
import manage_summary

ROOT = Path(__file__).resolve().parent.parent


class InventoryBuildTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.db = self.folder / 'summary.db'
        self.target = self.folder / 'inventory.html'
        self.base = self.folder / 'base.xlsx'
        self.monthly = self.folder / 'summay.xlsx'
        self.write_book(self.base, manage_summary.BASE_HEADERS, {
            'base': [('ABC01O', '舊藥名', 'O'), ('ABC02I', '另一品項', 'I')],
        })
        self.write_book(self.monthly, manage_summary.MONTH_HEADERS, {
            '11501': [('ABC01O', '舊藥名', 5, -2, -3, '115', '01', 'O'),
                      ('ABC02I', '另一品項', 0, 7, 9, '115', '01', 'I')],
            '11502': [('ABC01O', '新藥名', 7, 4, 0, '115', '02', 'I')],
        })
        manage_summary.initialize(self.db, self.base, self.monthly)

    def write_book(self, path, headers, sheets):
        book = Workbook()
        book.remove(book.active)
        for name, rows in sheets.items():
            sheet = book.create_sheet(name)
            sheet.append(headers)
            for row in rows:
                sheet.append(row)
        book.save(path)
        book.close()

    def build(self):
        return subprocess.run([
            sys.executable, str(ROOT / 'scripts/build_inventory.py'),
            '--db', str(self.db), '--output', str(self.target),
        ], cwd=self.folder, capture_output=True, text=True)

    def test_independent_build_retains_historical_rows_after_master_removal(self):
        self.write_book(self.base, manage_summary.BASE_HEADERS, {
            'base': [('ABC02I', '更新主檔', 'I')],
        })
        manage_summary.replace_base(self.db, self.base)
        before = hashlib.sha256(self.db.read_bytes()).digest()
        result = self.build()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(hashlib.sha256(self.db.read_bytes()).digest(), before)
        html = self.target.read_text()
        data = json.loads(html.split('<script id="inventoryData" type="application/json">')[1].split('</script>')[0])
        self.assertEqual(len(data['rows']), 2)
        self.assertEqual([r['藥品名稱'] for r in data['rows']], ['舊藥名', '新藥名'])
        self.assertEqual([r['劑型'] for r in data['rows']], ['O', 'I'])
        self.assertEqual(data['rows'][0]['住院耗用'], -2)
        self.assertEqual(data['rows'][0]['庫存量'], -3)
        self.assertNotIn('__INVENTORY_DATA__', html)
        self.assertNotIn('/DB_Viz/', html)

    def test_invalid_database_never_replaces_existing_html(self):
        for sql in (
            'PRAGMA ignore_check_constraints=ON; UPDATE monthly_summary SET issued_quantity=0;',
            'PRAGMA ignore_check_constraints=ON; UPDATE monthly_summary SET roc_month=13 WHERE rowid=1;',
            'DELETE FROM drug_codes;',
            'ALTER TABLE monthly_summary DROP COLUMN stock_quantity;',
            'ALTER TABLE monthly_summary RENAME TO old_summary; CREATE TABLE monthly_summary AS SELECT * FROM old_summary; INSERT INTO monthly_summary SELECT * FROM old_summary;',
            'DELETE FROM monthly_summary;',
        ):
            with self.subTest(sql=sql):
                original = self.db.read_bytes()
                try:
                    with sqlite3.connect(self.db) as db:
                        db.executescript(sql)
                    self.target.write_text('既有成品')
                    self.assertNotEqual(self.build().returncode, 0)
                    self.assertEqual(self.target.read_text(), '既有成品')
                finally:
                    self.db.write_bytes(original)

    def test_output_cannot_overwrite_source_database(self):
        before = self.db.read_bytes()
        self.target = self.db
        self.assertNotEqual(self.build().returncode, 0)
        self.assertEqual(self.db.read_bytes(), before)

    def test_header_only_month_clears_old_rows_and_preserves_other_months(self):
        self.write_book(self.monthly, manage_summary.MONTH_HEADERS, {'11502': []})
        report = manage_summary.update_monthly(self.db, self.monthly)
        self.assertEqual(report['各月替換前筆數'], {'115/02': 1})
        with sqlite3.connect(self.db) as db:
            self.assertEqual(db.execute('SELECT roc_month FROM monthly_summary').fetchall(), [(1,)])
        self.assertEqual(self.build().returncode, 0)

    def test_ambiguous_empty_month_does_not_change_database(self):
        self.write_book(self.monthly, manage_summary.MONTH_HEADERS, {'unknown': []})
        before = self.db.read_bytes()
        with self.assertRaises(manage_summary.InputError):
            manage_summary.update_monthly(self.db, self.monthly)
        self.assertEqual(self.db.read_bytes(), before)


if __name__ == '__main__':
    unittest.main()

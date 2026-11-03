"""驗證三份原始檔至資料庫與部署檔的整體流程。"""
import csv
from pathlib import Path
import sqlite3
import tempfile
import unittest

from openpyxl import Workbook
import monthly_pipeline as pipeline
import update_summary


def workbook(path, title, headers, rows):
    book = Workbook()
    sheet = book.active
    sheet.title = title
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    book.save(path)
    book.close()


def csv_file(path, headers, rows):
    with path.open('w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


class MonthlyPipelineTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.p = Path(self.temp.name)
        self.base = self.p / 'base.xlsx'
        self.db = self.p / 'summary.db'
        self.monthly = self.p / '2026年2月報表.xlsx'
        self.out = self.p / '11502_out.csv'
        self.ud = self.p / '11502_ud.csv'
        self.targets = [self.p / f'{x}.html' for x in ('site', 'local', 'dist')]
        workbook(self.base, 'base', ['drug_id', 'drug_name', 'drug_type'], [['ABC01O', 'Alpha', 'O']])
        m = self.p / 'initial.xlsx'
        t = self.p / 'total.xlsx'
        workbook(m, '11501', ['drug_id', '實發量', '庫存量'], [['ABC01O', 3, 4]])
        workbook(t, 'total', ['drug_id', 'total_qty', 'inv_year', 'inv_month'], [['ABC01O', 1, '115', '01']])
        update_summary.run(self.db, self.base, m, t, initialize=True)
        workbook(self.monthly, '2026年2月', ['藥品代碼', '實發量', '庫存量'], [['ABC01O', 12, 5]])
        csv_file(self.out, ['drug_id', 'drug_name', 'total_qty', 'phoutid'], [['ABC01O', 'Alpha', 7, 5630]])
        csv_file(self.ud, ['drug_id', 'drug_name', 'total_qty', 'phtxid', 'phoutid'], [['ABC01O', 'Alpha', 2, 2, 5630]])
        for target in self.targets:
            target.write_text('old')

    def run_pipeline(self, **kw):
        return pipeline.run(self.monthly, self.out, self.ud, self.db, self.base, self.targets, **kw)

    def rows(self):
        with sqlite3.connect(self.db) as db:
            return db.execute('select roc_month,issued_quantity,inpatient_usage,stock_quantity from monthly_summary order by roc_month').fetchall()

    def test_dry_run_then_commit_and_repeat(self):
        before = self.db.read_bytes()
        self.run_pipeline(dry_run=True)
        self.assertEqual(self.db.read_bytes(), before)
        self.assertTrue(all(t.read_text() == 'old' for t in self.targets))
        result = self.run_pipeline()
        self.assertEqual(self.rows(), [(1, 3, 1, 4), (2, 12, 5, 5)])
        self.assertTrue(Path(result['匯入結果']['更新前備份']).is_file())
        self.assertEqual(len({t.read_bytes() for t in self.targets}), 1)
        self.run_pipeline()
        self.assertEqual(self.rows(), [(1, 3, 1, 4), (2, 12, 5, 5)])

    def test_mismatched_period_and_formula_preserve_files(self):
        before = self.db.read_bytes()
        self.ud = self.p / '11503_ud.csv'
        self.ud.write_text('')
        with self.assertRaisesRegex(ValueError, '年月不一致'):
            self.run_pipeline()
        self.ud = self.p / '11502_ud.csv'
        workbook(self.monthly, '2026年2月', ['藥品代碼', '實發量', '庫存量'], [['ABC01O', '=1+1', 5]])
        with self.assertRaisesRegex(ValueError, '公式'):
            self.run_pipeline()
        self.assertEqual(before, self.db.read_bytes())
        self.assertTrue(all(t.read_text() == 'old' for t in self.targets))

    def test_unknown_drug_and_bad_csv_do_not_change_outputs(self):
        before = self.db.read_bytes()
        workbook(self.monthly, '2026年2月', ['藥品代碼', '實發量', '庫存量'], [['BAD01O', 4, 5]])
        with self.assertRaises(ValueError):
            self.run_pipeline()
        workbook(self.monthly, '2026年2月', ['藥品代碼', '實發量', '庫存量'], [['ABC01O', 4, 5]])
        self.ud.write_text('not,the,required,columns\n1,2,3,4\n')
        with self.assertRaises(ValueError):
            self.run_pipeline()
        self.assertEqual(before, self.db.read_bytes())
        self.assertTrue(all(t.read_text() == 'old' for t in self.targets))


if __name__ == '__main__':
    unittest.main()

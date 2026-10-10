"""以三份來源 Excel 驗證單一入口的交易、備份與失敗保護。"""

import contextlib
import io
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook
from sqlalchemy.exc import IntegrityError

import manage_summary as manager
import summay
import update_summary as pipeline


def excel(path, columns, sheets):
    book = Workbook()
    book.remove(book.active)
    for title, rows in sheets.items():
        sheet = book.create_sheet(title)
        sheet.append(columns)
        for row in rows:
            sheet.append(row)
    book.save(path)
    book.close()


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.addCleanup(self.folder.cleanup)
        self.root = Path(self.folder.name)
        self.db = self.root / "summary.db"
        self.base = self.root / "base.xlsx"
        self.monthly = self.root / "月報表.xlsx"
        self.total = self.root / "total_screen.xlsx"
        self.bases = [["ABC01O", "Alpha", "O"], ["ABC02I", "Beta", "I"]]
        excel(self.base, summay.BASE_COLUMNS, {"base": self.bases})
        self.inputs({"11501": [["ABC01O", 10, -3], ["ABC02I", 20, 5]],
                     "11502": [["ABC01O", 30, 7]]},
                    {"耗量": [["ABC01O", -2, "115", "01"], ["ABC01O", 8, "115", "02"]]})
        self.run_pipeline(initialize=True)

    def inputs(self, months, totals=None):
        excel(self.monthly, summay.MONTHLY_COLUMNS, months)
        excel(self.total, summay.TOTAL_COLUMNS, totals or {"耗量": []})

    def run_pipeline(self, **kwargs):
        return pipeline.run(self.db, self.base, self.monthly, self.total, **kwargs)

    def rows(self, table="monthly_summary", db=None):
        with sqlite3.connect(db or self.db) as connection:
            return connection.execute(f"SELECT * FROM {table} ORDER BY 1,2,3").fetchall()

    def test_single_month_replaces_complete_month_retains_history_and_repeats(self):
        self.inputs({"11502": [["ABC02I", 99, -4]]}, {"耗量": [["ABC02I", -7, "115", "2"]]})
        excel(self.base, summay.BASE_COLUMNS, {"base": [["ABC01O", "新名稱", "E"], self.bases[1]]})
        before = self.rows()
        report = self.run_pipeline()
        self.assertEqual(self.rows(), [before[0], ("ABC02I", 115, 1, "Beta", "I", 20, 0, 5),
                                      ("ABC02I", 115, 2, "Beta", "I", 99, -7, -4)])
        self.assertEqual(self.rows(db=Path(report["更新前備份"])), before)
        self.assertEqual(self.rows("drug_master")[0], ("ABC01O", "Alpha", "O"))
        expected = self.rows()
        self.run_pipeline()
        self.assertEqual(self.rows(), expected)

    def test_multiple_sheets_match_by_period_and_code(self):
        self.inputs({"11503": [["ABC01O", 4, 8]], "11504": [["ABC02I", 6, 9]]},
                    {"不同排序四月": [["ABC02I", 77, "115", "04"]],
                     "不同排序三月": [["ABC01O", -5, "115", "3"]]})
        report = self.run_pipeline()
        self.assertEqual(report["各月匯入筆數"], {"115/03": 1, "115/04": 1})
        self.assertIn(("ABC01O", 115, 3, "Alpha", "O", 4, -5, 8), self.rows())
        self.assertIn(("ABC02I", 115, 4, "Beta", "I", 6, 77, 9), self.rows())

    def test_zero_and_header_only_months_clear_old_rows(self):
        self.inputs({"11501": [["ABC01O", 0, 8], ["ABC02I", -1, 0]], "11502": []})
        report = self.run_pipeline()
        self.assertEqual(self.rows(), [])
        self.assertEqual(report["來源筆數"], 2)
        self.assertEqual(report["排除筆數"], 2)
        self.assertEqual(report["各月匯入筆數"], {"115/01": 0, "115/02": 0})

    def test_missing_db_is_not_created(self):
        self.db.unlink()
        with self.assertRaisesRegex(manager.InputError, "不存在"):
            self.run_pipeline()
        self.assertFalse(self.db.exists())

    def test_dry_run_preserves_all_files_and_creates_no_backups(self):
        before = {p.name: p.read_bytes() for p in self.root.iterdir()}
        report = self.run_pipeline(dry_run=True)
        self.assertTrue(report["預演"])
        self.assertEqual({p.name: p.read_bytes() for p in self.root.iterdir()}, before)

    def test_replace_base_only_retains_history_without_monthly_inputs(self):
        before = self.rows()
        excel(self.base, summay.BASE_COLUMNS, {"base": [self.bases[1], ["ABC03E", "Gamma", "E"]]})
        self.monthly.unlink()
        self.total.unlink()
        self.run_pipeline(replace_base=True)
        self.assertEqual(self.rows(), before)
        self.assertEqual(self.rows("drug_master"), [("ABC02I", "Beta", "I"), ("ABC03E", "Gamma", "E")])
        with sqlite3.connect(self.db) as connection:
            self.assertEqual(connection.execute("SELECT count(*) FROM drug_codes").fetchone()[0], 3)

    def test_excluded_unknown_db_code_is_rejected_before_backup(self):
        excel(self.base, summay.BASE_COLUMNS, {"base": self.bases + [["ABC03E", "Gamma", "E"]]})
        self.inputs({"11503": [["ABC01O", 5, 9]], "11504": [["ABC03E", 0, 9]]})
        before = self.db.read_bytes()
        with self.assertRaisesRegex(manager.InputError, "ABC03E.*月報表.xlsx.*11504.*第 2 列"):
            self.run_pipeline()
        self.assertEqual(self.db.read_bytes(), before)
        self.assertFalse((self.root / "backups").exists())

    def test_invalid_sources_including_excluded_rows_never_write(self):
        before = self.db.read_bytes()
        cases = [("total", 1, 1.5), ("total", 1, str(2**63)),
                 ("total", 1, "=1+2"), ("total", 1, "#DIV/0!"),
                 ("monthly", 2, 1.5), ("monthly", 2, "=1+2"),
                 ("monthly", 2, str(2**63)), ("base", 1, "=1+2")]
        for source, column, value in cases:
            with self.subTest(source=source, value=value):
                excel(self.base, summay.BASE_COLUMNS, {"base": self.bases})
                self.inputs({"11503": [["ABC01O", 0, 2]]}, {"耗量": [["ABC01O", 5, "115", "03"]]})
                if source == "total":
                    row = ["ABC01O", 5, "115", "03"]
                    row[column] = value
                    excel(self.total, summay.TOTAL_COLUMNS, {"耗量": [row]})
                elif source == "monthly":
                    row = ["ABC01O", 0, 2]
                    row[column] = value
                    excel(self.monthly, summay.MONTHLY_COLUMNS, {"11503": [row]})
                else:
                    row = self.bases[0][:]
                    row[column] = value
                    excel(self.base, summay.BASE_COLUMNS, {"base": [row, self.bases[1]]})
                with self.assertRaises(ValueError):
                    self.run_pipeline()
                self.assertEqual(self.db.read_bytes(), before)
                self.assertFalse((self.root / "backups").exists())

    def test_duplicate_month_across_sheets_rejects_entire_input(self):
        before = self.db.read_bytes()
        self.inputs({"1153": [["ABC01O", 1, 0]], "11503": [["ABC01O", 2, 0]]})
        with self.assertRaisesRegex(ValueError, "重複鍵"):
            self.run_pipeline()
        self.assertEqual(self.db.read_bytes(), before)

    def test_write_failure_rolls_back_and_backup_is_before_update(self):
        with sqlite3.connect(self.db) as connection:
            connection.execute("CREATE TRIGGER reject_beta BEFORE INSERT ON monthly_summary "
                               "WHEN NEW.drug_code='ABC02I' BEGIN SELECT RAISE(ABORT,'測試失敗'); END")
        before = self.rows()
        self.inputs({"11501": [["ABC01O", 99, 1], ["ABC02I", 88, 2]]})
        with self.assertRaises(IntegrityError):
            self.run_pipeline()
        self.assertEqual(self.rows(), before)
        backups = list((self.root / "backups").glob("*.db"))
        self.assertEqual(len(backups), 1)
        self.assertEqual(self.rows(db=backups[0]), before)

    def test_default_cli_uses_script_directory_and_reports_errors(self):
        with patch.object(pipeline, "ROOT", self.root), contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(pipeline.main(["--dry-run"]), 0)
        with contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(pipeline.main(["--db", str(self.root / "missing.db")]), 2)

    def test_init_and_alias_paths_cannot_overwrite_files(self):
        before = self.db.read_bytes()
        with self.assertRaisesRegex(ValueError, "已存在"):
            self.run_pipeline(initialize=True)
        with self.assertRaisesRegex(ValueError, "輸入檔案相同"):
            pipeline.run(self.base, self.base, self.monthly, self.total)
        self.assertEqual(self.db.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()

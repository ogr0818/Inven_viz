"""以暫存 Excel / SQLite 驗證匯入、資料替換與失敗回復。"""

from pathlib import Path
import sqlite3
import tempfile
import unittest

from openpyxl import Workbook
from sqlalchemy import delete
from sqlalchemy.exc import IntegrityError

import manage_summary as manager


def excel(path, sheets, headers=manager.MONTH_HEADERS):
    book = Workbook()
    book.remove(book.active)
    for name, rows in sheets.items():
        sheet = book.create_sheet(name)
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
    book.save(path)
    book.close()
    return path


def month_row(code="A", month=1, issued=10, used=8, stock=2, name=None, kind="I", year="115"):
    return [code, name or f"{code} 當月名稱", issued, used, stock, year, f"{month:02d}", kind]


class DatabaseTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.root = Path(self.folder.name)
        self.db = self.root / "summary.db"
        self.base = excel(self.root / "base.xlsx", {"base": [
            ["A", "A 目前名稱", "I"], ["B", "B 目前名稱", 0],
        ]}, manager.BASE_HEADERS)
        self.summary = excel(self.root / "summary.xlsx", {"一月": [
            month_row("A", 1, 10, -2, -5), month_row("B", 1, 0),
        ], "二月": [month_row("A", 2, 20), month_row("B", 2, 5)]})
        manager.initialize(self.db, self.base, self.summary)

    def tearDown(self):
        self.folder.cleanup()

    def query(self, sql, args=(), path=None):
        connection = sqlite3.connect(path or self.db)
        try:
            return connection.execute(sql, args).fetchall()
        finally:
            connection.close()

    def snapshot(self, path=None):
        return {table: self.query(f"SELECT * FROM {table} ORDER BY 1, 2", path=path)
                if table != "drug_codes" else self.query(f"SELECT * FROM {table} ORDER BY 1", path=path)
                for table in ("drug_codes", "drug_master", "monthly_summary")}

    def test_initial_filter_signed_numbers_and_month_snapshots(self):
        self.assertEqual(self.query("SELECT count(*) FROM monthly_summary"), [(3,)])
        self.assertEqual(self.query("SELECT inpatient_usage,stock_quantity,drug_name FROM monthly_summary WHERE drug_code='A' AND roc_month=1"), [(-2, -5, "A 當月名稱")])
        self.assertEqual(self.query("SELECT drug_type FROM drug_master WHERE drug_code='B'"), [("0",)])
        info = manager.information(self.db)
        self.assertEqual(info["各月筆數"], {"115/01": 1, "115/02": 2})

    def test_corrected_complete_month_removes_omitted_drugs_only_in_that_month(self):
        new = excel(self.root / "correction.xlsx", {"二月更正": [month_row("B", 2, 30, -7, -11)]})
        report = manager.update_monthly(self.db, new)
        self.assertEqual(self.query("SELECT drug_code,roc_month,issued_quantity FROM monthly_summary ORDER BY roc_month,drug_code"), [("A", 1, 10), ("B", 2, 30)])
        self.assertEqual(report["各月替換前筆數"], {"115/02": 2})
        manager.update_monthly(self.db, new)
        self.assertEqual(self.query("SELECT count(*) FROM monthly_summary"), [(2,)])

    def test_all_zero_month_clears_old_month_and_preserves_other_month(self):
        new = excel(self.root / "zeros.xlsx", {"更正": [month_row("A", 2, 0), month_row("B", 2, -3)]})
        report = manager.update_monthly(self.db, new)
        self.assertEqual(report["各月匯入筆數"], {"115/02": 0})
        self.assertEqual(self.query("SELECT drug_code,roc_month FROM monthly_summary"), [("A", 1)])

    def test_multiple_sheets_combine_then_replace_complete_periods(self):
        new = excel(self.root / "multiple.xlsx", {"三月A": [month_row("A", 3)],
                    "三月B": [month_row("B", 3, 22)], "四月": [month_row("B", 4)]})
        manager.update_monthly(self.db, new)
        self.assertEqual(self.query("SELECT roc_month,count(*) FROM monthly_summary GROUP BY roc_month ORDER BY roc_month"), [(1, 1), (2, 2), (3, 2), (4, 1)])

    def test_replace_master_retains_registry_history_and_month_names(self):
        old_months = self.query("SELECT * FROM monthly_summary ORDER BY 1,2,3")
        new = excel(self.root / "new_base.xlsx", {"base": [["B", "B 新名稱", "O"], ["C", "C 新藥品", "E"]]}, manager.BASE_HEADERS)
        report = manager.replace_base(self.db, new)
        self.assertEqual(self.query("SELECT drug_code,drug_name FROM drug_master ORDER BY drug_code"), [("B", "B 新名稱"), ("C", "C 新藥品")])
        self.assertEqual(self.query("SELECT drug_code FROM drug_codes ORDER BY drug_code"), [("A",), ("B",), ("C",)])
        self.assertEqual(self.query("SELECT * FROM monthly_summary ORDER BY 1,2,3"), old_months)
        self.assertEqual(report["目前主檔未收錄的歷史月紀錄"], 2)
        retired = excel(self.root / "retired.xlsx", {"三月": [month_row("A", 3)]})
        with self.assertRaises(manager.InputError):
            manager.update_monthly(self.db, retired)

    def test_unknown_drug_in_later_sheet_rejects_entire_file(self):
        before = self.snapshot()
        new = excel(self.root / "unknown.xlsx", {"可用月份": [month_row("A", 3)], "錯誤月份": [month_row("UNKNOWN", 4)]})
        with self.assertRaisesRegex(manager.InputError, "UNKNOWN"):
            manager.update_monthly(self.db, new)
        self.assertEqual(self.snapshot(), before)
        self.assertFalse((self.root / "backups").exists())

    def test_duplicate_key_across_sheets_rejects_entire_file(self):
        before = self.snapshot()
        new = excel(self.root / "duplicate.xlsx", {"一": [month_row("A", 3)], "二": [month_row("A", 3)]})
        with self.assertRaisesRegex(manager.InputError, "重複藥碼年月"):
            manager.update_monthly(self.db, new)
        self.assertEqual(self.snapshot(), before)

    def test_fractional_invalid_month_and_formula_reject_without_changes(self):
        before = self.snapshot()
        cases = [month_row(issued=1.2), month_row(month=13), month_row(used="=1+2")]
        for n, row in enumerate(cases):
            with self.subTest(n=n):
                new = excel(self.root / f"invalid-{n}.xlsx", {"月資料": [row]})
                with self.assertRaises(manager.InputError):
                    manager.update_monthly(self.db, new)
                self.assertEqual(self.snapshot(), before)

    def test_rollback_after_delete_and_partial_insert_failure(self):
        connection = sqlite3.connect(self.db)
        connection.execute("CREATE TRIGGER reject_b BEFORE INSERT ON monthly_summary WHEN NEW.drug_code='B' BEGIN SELECT RAISE(ABORT,'測試資料庫寫入失敗'); END")
        connection.commit()
        connection.close()
        before = self.snapshot()
        new = excel(self.root / "failure.xlsx", {"一月": [month_row("A", 1, 99), month_row("B", 1, 55)]})
        with self.assertRaises(IntegrityError):
            manager.update_monthly(self.db, new)
        self.assertEqual(self.snapshot(), before)
        self.assertEqual(len(list((self.root / "backups").glob("*.db"))), 1)

    def test_dry_run_never_changes_data_or_creates_backup(self):
        before = self.snapshot()
        new = excel(self.root / "dry.xlsx", {"一月": [month_row("A", 1, 999)]})
        self.assertTrue(manager.update_monthly(self.db, new, dry_run=True)["預演"])
        manager.replace_base(self.db, self.base, dry_run=True)
        self.assertEqual(self.snapshot(), before)
        self.assertFalse((self.root / "backups").exists())

    def test_backup_contains_exact_preupdate_data(self):
        before = self.snapshot()
        new = excel(self.root / "update.xlsx", {"二月": [month_row("B", 2, 999)]})
        report = manager.update_monthly(self.db, new)
        self.assertEqual(self.snapshot(Path(report["更新前備份"])), before)
        self.assertNotEqual(self.snapshot(), before)
        self.assertEqual(manager.information(Path(report["更新前備份"]))["完整性檢查"], "通過")

    def test_init_refuses_overwrite_and_empty_master(self):
        before = self.snapshot()
        with self.assertRaisesRegex(manager.InputError, "資料庫已存在"):
            manager.initialize(self.db, self.base, self.summary)
        empty = excel(self.root / "empty.xlsx", {"base": []}, manager.BASE_HEADERS)
        with self.assertRaisesRegex(manager.InputError, "空主檔"):
            manager.replace_base(self.db, empty)
        self.assertEqual(self.snapshot(), before)

    def test_schema_version_and_foreign_key_guard(self):
        engine = manager.database_engine(self.db)
        try:
            with engine.connect().execution_options(write_transaction=True) as connection:
                with self.assertRaises(IntegrityError):
                    with connection.begin():
                        connection.execute(delete(manager.DrugCode).where(manager.DrugCode.drug_code == "A"))
        finally:
            engine.dispose()
        connection = sqlite3.connect(self.db)
        connection.execute("PRAGMA user_version = 999")
        connection.close()
        with self.assertRaisesRegex(manager.InputError, "版本"):
            manager.information(self.db)

    def test_exported_schema_recreates_matching_database(self):
        db = self.root / "empty-schema.db"
        connection = sqlite3.connect(db)
        connection.executescript(manager.schema_sql())
        connection.close()
        self.assertEqual(manager.information(db)["資料表筆數"], {
            "drug_codes": 0, "drug_master": 0, "monthly_summary": 0,
        })


if __name__ == "__main__":
    unittest.main()

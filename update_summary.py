#!/usr/bin/env python3
"""單一入口：三份 Excel 更新既有 summary.db；主檔更新需明確指定。"""

import argparse
import json
from pathlib import Path
import sqlite3
import sys
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException
from sqlalchemy.exc import SQLAlchemyError

import manage_summary as manager
import summay


ROOT = Path(__file__).resolve().parent


def monthly_data(merged: summay.MonthlyMerge) -> manager.MonthlyImport:
    """保留篩選前藥碼、月份與筆數，供 DB 交易完整驗證。"""
    data = manager.MonthlyImport(
        [], set(), set(merged.code_locations), dict(merged.code_locations),
        sum(c["input"] for c in merged.counts.values()), list(merged.outputs),
    )
    for sheet, frame in merged.outputs.items():
        year, month = map(int, summay.inventory_period(sheet))
        if year <= 0:
            raise manager.InputError(f"月報工作表 {sheet}：民國年份須大於零。")
        data.periods.add((year, month))
        for code, name, issued, used, stock, _, _, kind in frame.itertuples(index=False, name=None):
            location = data.locations[code]
            data.rows.append({
                "drug_code": code, "roc_year": year, "roc_month": month,
                "drug_name": manager.text_value(name, "藥品名稱", location),
                "drug_type": manager.text_value(kind, "劑型", location),
                "issued_quantity": manager.integer_value(issued, "實發量", location),
                "inpatient_usage": manager.integer_value(used, "住院耗用", location),
                "stock_quantity": manager.integer_value(stock, "庫存量", location),
            })
    return data


def run(db: Path, base: Path, monthly: Path, total: Path, *,
        replace_base: bool = False, initialize: bool = False,
        dry_run: bool = False) -> dict:
    db, base, monthly, total = (p.expanduser().resolve() for p in (db, base, monthly, total))
    if replace_base and initialize:
        raise manager.InputError("首次建立與更新主檔不可同時執行。")
    if db in {base, monthly, total}:
        raise manager.InputError("資料庫路徑不可與任何輸入檔案相同。")
    if initialize:
        if db.exists():
            raise manager.InputError(f"資料庫已存在：{db}；首次建立不會覆寫。")
    else:
        manager.require_database(db)
    if replace_base:
        return manager.replace_base(db, base, dry_run)
    merged = summay.prepare_monthly(base, monthly, total)
    data = monthly_data(merged)
    if initialize:
        report = manager.initialize_data(db, manager.read_master(base), data, dry_run)
    else:
        report = manager.update_monthly_data(db, data, dry_run)
    report["來源檔案"] = {"主檔": str(base), "月報": str(monthly), "耗量": str(total)}
    report["各表彙整"] = merged.counts
    report["缺耗量補零筆數"] = sum(c["filled_zero"] for c in merged.counts.values())
    return report


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    action = parser.add_mutually_exclusive_group()
    action.add_argument("--replace-base", action="store_true", help="只更新目前主檔，保留歷史月明細")
    action.add_argument("--init", action="store_true", help="明確首次建立；拒絕覆寫既有 DB")
    parser.add_argument("--db", type=Path, default=ROOT / "summary.db")
    parser.add_argument("--base", type=Path, default=ROOT / "base.xlsx")
    parser.add_argument("--monthly", type=Path, default=ROOT / "月報表.xlsx")
    parser.add_argument("--total", type=Path, default=ROOT / "total_screen.xlsx")
    parser.add_argument("--dry-run", action="store_true", help="只驗證及預演，不修改檔案或建立備份")
    args = parser.parse_args(argv)
    try:
        report = run(args.db, args.base, args.monthly, args.total,
                     replace_base=args.replace_base, initialize=args.init, dry_run=args.dry_run)
    except (ValueError, OSError, sqlite3.Error, SQLAlchemyError, BadZipFile, InvalidFileException) as exc:
        print(f"作業失敗：{exc}\n未提交資料庫更新；請修正原因後重新執行。", file=sys.stderr)
        return 2
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
'''
uv run update_summary.py
基本檔更新使用 --replace-base
'''
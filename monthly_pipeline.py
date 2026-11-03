#!/usr/bin/env python3
"""由月報、OUT 及 UD 原始檔更新資料庫並準備 Sites 網頁。"""
import argparse
from contextlib import closing
import hashlib
import json
from pathlib import Path
import re
import shutil
import sqlite3
import sys
import tempfile

from openpyxl import load_workbook

import manage_summary as manager
import out_screen
import rename_monthly_xlsx
import total_screen
import ud_screen
import update_summary
from scripts.build_inventory import build, read_data

ROOT = Path(__file__).resolve().parent


def check_inputs(monthly, out, ud):
    for path in (monthly, out, ud):
        if not path.is_file():
            raise ValueError(f"找不到來源檔案：{path}")
    if len({p.resolve() for p in (monthly, out, ud)}) != 3:
        raise ValueError("三份來源必須是不同檔案。")
    if monthly.suffix.lower() != '.xlsx':
        raise ValueError('月報必須是 .xlsx 檔案。')
    periods = []
    for path, kind in ((out, 'out'), (ud, 'ud')):
        match = re.fullmatch(r'(\d{3})(\d{2})_' + kind + r'\.csv', path.name, re.I)
        if not match or int(match[1]) == 0 or not 1 <= int(match[2]) <= 12:
            raise ValueError(f'{path.name}：檔名須為民國年月_{kind}.csv，例如 11510_{kind}.csv。')
        periods.append(match[1] + match[2])
    if periods[0] != periods[1]:
        raise ValueError('OUT 與 UD 檔名的年月不一致。')
    book = load_workbook(monthly, read_only=True, data_only=False)
    try:
        names = [rename_monthly_xlsx.month_sheet_name(s.title) for s in book]
        if names != [periods[0]]:
            raise ValueError(f'月報必須只有一張 {periods[0]} 的工作表；實際為 {names}。')
        # 原始公式不得被快取值靜默取代，避免缺少或過期計算結果。
        sheet = book.worksheets[0]
        headers = [c.value for c in next(sheet.iter_rows(max_row=1))]
        required = ('藥品代碼', 'drug_id', '實發量', '庫存量')
        cols = [i for i, name in enumerate(headers) if name in required]
        for row in sheet.iter_rows(min_row=2):
            for i in cols:
                if row[i].data_type in ('f', 'e'):
                    raise ValueError(f'月報 {sheet.title}！{row[i].coordinate} 含公式或錯誤值，請先提供確定數值。')
    finally:
        book.close()
    return periods[0]


def snapshot(source, target):
    manager.require_database(source)
    with closing(sqlite3.connect(source.resolve().as_uri() + '?mode=ro', uri=True)) as src:
        with closing(sqlite3.connect(target)) as dst:
            src.backup(dst)


def run(monthly, out, ud, db, base, targets, dry_run=False):
    monthly, out, ud, db, base = [Path(p).expanduser().absolute() for p in (monthly, out, ud, db, base)]
    targets = [Path(p).expanduser().resolve() for p in targets]
    protected = {p.resolve() for p in (monthly, out, ud, db, base, ROOT / 'scripts/inventory_template.html')}
    if len(set(targets)) != 3 or set(targets) & protected:
        raise ValueError('三個 HTML 輸出路徑必須互不相同，且不可覆寫來源、資料庫或範本。')
    period = check_inputs(monthly, out, ud)
    if not base.is_file():
        raise ValueError(f'找不到本機藥品主檔：{base}')
    with tempfile.TemporaryDirectory(prefix='inventory-monthly-') as tmp:
        temp = Path(tmp)
        snapshot(db, temp / 'summary.db')
        rename_monthly_xlsx.convert_xlsx(monthly, output_path=temp / '月報表.xlsx')
        for source, name, module in ((out, 'out', out_screen), (ud, 'ud', ud_screen)):
            folder = temp / name
            folder.mkdir()
            shutil.copyfile(source, folder / f'{period}.csv')
            module.screen_to_excel(folder, temp / f'{name}_screen.xlsx')
        total_screen.total_to_excel(temp)
        update_args = (base, temp / '月報表.xlsx', temp / 'total_screen.xlsx')
        trial = update_summary.run(temp / 'summary.db', *update_args)
        build(temp / 'summary.db', temp / 'index.html')
        staged_data = read_data(temp / 'summary.db')
        for key in ('資料庫', '預演', '更新前備份', '來源檔案'):
            trial.pop(key, None)
        print('暫存資料庫與網頁驗證完成。', flush=True)
        if dry_run:
            content = (temp / 'index.html').read_bytes()
            return {'預演': True, '月份': period, '匯入驗證': trial,
                    '月明細總筆數': len(staged_data['rows']),
                    '與現有HTML相同': {str(p): p.is_file() and p.read_bytes() == content for p in targets}}
        for target in targets:
            if not target.parent.is_dir():
                raise ValueError(f'輸出資料夾不存在：{target.parent}')
        report = update_summary.run(db, *update_args)
        report['來源檔案'] = {'月報': str(monthly), 'OUT': str(out), 'UD': str(ud), '本機主檔': str(base)}
        try:
            if read_data(db) != staged_data:
                raise ValueError('正式資料庫與預演結果不同，未更新 HTML；請檢查是否有其他程序同時修改資料。')
            content = (temp / 'index.html').read_bytes()
            for target in targets:
                with tempfile.NamedTemporaryFile(dir=target.parent, prefix='.inventory-', delete=False) as f:
                    staged = Path(f.name)
                    try:
                        f.write(content)
                    except Exception:
                        staged.unlink(missing_ok=True)
                        raise
                try:
                    staged.replace(target)
                finally:
                    staged.unlink(missing_ok=True)
            if any(t.read_bytes() != content for t in targets):
                raise ValueError('HTML 同步檢查失敗。')
        except Exception as exc:
            raise RuntimeError(f'資料庫已更新，但網頁準備未完成：{exc}。備份：{report.get("更新前備份")}；修正後重新執行。') from exc
        return {'預演': False, '月份': period, '匯入結果': report,
                '月明細總筆數': len(staged_data['rows']), 'HTML': [str(p) for p in targets],
                'HTML_SHA256': hashlib.sha256(content).hexdigest(), 'Sites': '僅準備檔案，尚未部署'}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('monthly', type=Path, help='原始月份報表 .xlsx')
    parser.add_argument('out', type=Path, help='民國年月_out.csv')
    parser.add_argument('ud', type=Path, help='民國年月_ud.csv')
    parser.add_argument('--dry-run', action='store_true', help='完整預演，不更新正式資料庫或 HTML')
    parser.add_argument('--db', type=Path, default=ROOT / 'summary.db')
    parser.add_argument('--base', type=Path, default=ROOT / 'base.xlsx')
    args = parser.parse_args(argv)
    try:
        report = run(args.monthly, args.out, args.ud, args.db, args.base,
                     [ROOT / 'sites/inven-viz/dist/index.html', ROOT / 'inventory.html', ROOT / 'dist/index.html'], args.dry_run)
    except Exception as exc:
        print(f'作業失敗：{exc}', file=sys.stderr)
        return 1
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

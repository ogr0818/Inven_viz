"""驗證唯讀 SQLite 來源，產生可獨立開啟的進耗存 HTML。"""
import argparse
import json
import re
import sqlite3
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
COLUMNS = ['藥品代碼', '藥品名稱', '實發量', '住院耗用', '庫存量', '撥補年份', '撥補月份', '劑型']


def read_data(source):
    with sqlite3.connect(source.resolve().as_uri() + '?mode=ro', uri=True) as db:
        if db.execute('PRAGMA integrity_check').fetchall() != [('ok',)]:
            raise ValueError('資料庫完整性檢查失敗')
        if db.execute('PRAGMA foreign_key_check').fetchone():
            raise ValueError('資料庫存在無法配對的藥品代碼')
        records = db.execute('''SELECT drug_code, drug_name, issued_quantity,
            inpatient_usage, stock_quantity, roc_year, roc_month, drug_type
            FROM monthly_summary ORDER BY roc_year, roc_month, drug_code''').fetchall()
    if not records:
        raise ValueError('月資料為空')
    seen, rows, periods = set(), [], set()
    for index, values in enumerate(records, 1):
        code, name, issued, usage, stock, year, month, kind = values
        key = (code, year, month)
        context = f'第 {index} 筆 {key}'
        if not isinstance(code, str) or not re.fullmatch('[A-Z0-9]{6}', code):
            raise ValueError(context + '：藥品代碼須為六碼大寫英數')
        if not isinstance(name, str) or not name.strip() or not isinstance(kind, str) or not kind.strip():
            raise ValueError(context + '：名稱或劑型缺漏')
        if any(type(v) is not int for v in (issued, usage, stock, year, month)):
            raise ValueError(context + '：數量與年月須為整數')
        if issued <= 0 or year <= 0 or not 1 <= month <= 12:
            raise ValueError(context + '：實發量或年月不符本次資料契約')
        if key in seen:
            raise ValueError(context + '：重複品項月份')
        seen.add(key)
        periods.add(year * 100 + month)
        rows.append(dict(zip(COLUMNS, [code, name, issued, usage, stock, str(year), f'{month:02}', kind])))
    return {'columns': COLUMNS, 'sourcePeriods': sorted(periods), 'rows': rows}


def build(source, target):
    if source.resolve() == target.resolve():
        raise ValueError('輸出不可覆寫來源資料庫')
    if target.resolve() == (ROOT / 'scripts/inventory_template.html').resolve():
        raise ValueError('輸出不可覆寫 HTML 範本')
    data = read_data(source)
    template = (ROOT / 'scripts/inventory_template.html').read_text()
    if template.count('__INVENTORY_DATA__') != 1:
        raise ValueError('HTML 範本資料標記不唯一')
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('<', '\\u003c').replace('&', '\\u0026')
    html = template.replace('__INVENTORY_DATA__', payload)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix('.html.tmp')
    temporary.write_text(html)
    temporary.replace(target)
    print(json.dumps({'output': str(target), 'records': len(data['rows']),
                      'items': len({r['藥品代碼'] for r in data['rows']}),
                      'periods': data['sourcePeriods']}, ensure_ascii=False))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--db', type=Path, default=ROOT / 'summary.db')
    parser.add_argument('--output', type=Path, default=ROOT / 'inventory.html')
    args = parser.parse_args()
    build(args.db, args.output)

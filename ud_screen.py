"""依序篩選 UD CSV，將每個來源的結果寫入 Excel 工作表。"""

import argparse
import math
from pathlib import Path
import re
import sys
import tempfile

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from quantity_rules import filter_quantity_range


COLUMNS = ["drug_id", "drug_name", "total_qty", "phtxid", "phoutid"]
OUTPUT_COLUMNS = [*COLUMNS, "inv_year", "inv_month"]
MAX_DATA_ROWS = 1_048_575  # Excel 列數上限扣除標題列。
BASE_DIR = Path(__file__).resolve().parent


def read_screened_csv(path: Path) -> pd.DataFrame:
    """讀取必要欄位，篩選院區並依交易代碼統一數量符號。"""
    with path.open("rb") as source:
        prefix = source.read(4)
    encoding = "utf-16" if prefix.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig"
    try:
        frame = pd.read_csv(
            path,
            encoding=encoding,
            # 保留 NUL 等控制字元，避免 C 解析器靜默截斷文字。
            engine="python",
            usecols=lambda column: column in COLUMNS,
            dtype="string",
            keep_default_na=False,
        )
    except (ValueError, UnicodeError) as exc:
        raise ValueError(f"{path.name}：無法讀取 CSV：{exc}") from exc
    missing = [column for column in COLUMNS if column not in frame.columns]
    if missing:
        raise ValueError(f"{path.name}：缺少必要欄位：{', '.join(missing)}")

    outlet = pd.to_numeric(frame["phoutid"], errors="coerce")
    result = frame.loc[outlet.eq(5630).fillna(False), COLUMNS].copy()
    result = filter_quantity_range(result, path.name)
    for column in ("drug_id", "drug_name"):
        invalid = result[column].str.contains(ILLEGAL_CHARACTERS_RE, na=False)
        if invalid.any():
            indices = result.index[invalid]
            rows = ", ".join(str(int(index) + 2) for index in indices[:5])
            suffix = " 等" if len(indices) > 5 else ""
            raise ValueError(f"{path.name}：第 {rows}{suffix} 筆資料列（含標題列）{column} 含 Excel 不允許的控制字元")
    blank = result["drug_id"].isna() | result["drug_id"].str.strip().eq("")
    if blank.any():
        indices = result.index[blank]
        rows = ", ".join(str(int(index) + 2) for index in indices[:5])
        suffix = " 等" if len(indices) > 5 else ""
        raise ValueError(f"{path.name}：第 {rows}{suffix} 筆資料列（含標題列）drug_id 空白")
    transactions = pd.to_numeric(result["phtxid"], errors="coerce")
    valid = transactions.notna() & transactions.map(
        lambda value: math.isfinite(value) if pd.notna(value) else False
    )
    valid &= transactions.mod(1).eq(0).fillna(False)
    if not valid.all():
        raise ValueError(f"{path.name}：符合篩選條件的 phtxid 含無效數值")
    result["phtxid"] = transactions

    result["phoutid"] = 5630
    result["total_qty"] = result["total_qty"].abs()
    even = result["phtxid"].mod(2).eq(0)
    result.loc[even, "total_qty"] = -result.loc[even, "total_qty"]
    result.loc[result["total_qty"].eq(0), "total_qty"] = 0
    if len(result) > MAX_DATA_ROWS:
        raise ValueError(f"{path.name}：結果超過 Excel 每張工作表 {MAX_DATA_ROWS:,} 筆資料的上限")
    return result


def sheet_name(stem: str, used: set[str]) -> str:
    """處理 Excel 工作表名稱限制及不分大小寫的名稱衝突。"""
    base = re.sub(r"[\\/*?:\[\]]", "_", ILLEGAL_CHARACTERS_RE.sub("_", stem)).strip("'") or "Sheet"
    name = base[:31]
    suffix = 2
    while name.casefold() in used:
        tail = f"_{suffix}"
        name = base[: 31 - len(tail)] + tail
        suffix += 1
    used.add(name.casefold())
    return name


def inventory_period(name: str) -> tuple[str, str]:
    """從工作表名稱取得三位年份及兩位月份文字。"""
    if not re.fullmatch(r"[0-9]{3}[0-9]{1,2}", name):
        raise ValueError(f"工作表 {name!r}：名稱須為 3 位數字年份及 1–12 月")
    year, month = name[:3], name[3:]
    if not 1 <= int(month) <= 12:
        raise ValueError(f"工作表 {name!r}：月份須介於 1–12")
    return year, month.zfill(2)


def screen_to_excel(input_dir: Path, output: Path) -> dict[str, int]:
    """依檔名排序處理 CSV；成功完成後才取代輸出檔案。"""
    if not input_dir.is_dir():
        raise ValueError(f"輸入資料夾不存在：{input_dir}")
    sources = sorted(
        (path for path in input_dir.iterdir() if path.is_file() and path.suffix.lower() == ".csv"),
        key=lambda path: path.name,
    )
    if not sources:
        raise ValueError(f"找不到 CSV 檔案：{input_dir}")
    if output.suffix.lower() != ".xlsx":
        raise ValueError("輸出檔案副檔名必須是 .xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)
    counts = {}
    used = set()
    # 逐列寫入，避免九個月份的資料同時佔用記憶體。
    # 每個來源先完成驗證，再開啟工作表。
    workbook = Workbook(write_only=True)
    temporary = None
    try:
        for path in sources:
            name = sheet_name(path.stem, used)
            year, month = inventory_period(name)
            frame = read_screened_csv(path)
            frame["inv_year"] = pd.Series(year, index=frame.index, dtype="string")
            frame["inv_month"] = pd.Series(month, index=frame.index, dtype="string")
            worksheet = workbook.create_sheet(name)
            worksheet.append(OUTPUT_COLUMNS)
            for row in frame.itertuples(index=False, name=None):
                # 明確寫入文字，避免以 = 開頭的藥名被 Excel 當成公式。
                cells = []
                for value in row:
                    cell = WriteOnlyCell(worksheet, value=value)
                    if isinstance(value, str):
                        cell.data_type = "s"
                    cells.append(cell)
                worksheet.append(cells)
            counts[path.name] = len(frame)
            ignored = frame.attrs["ignored_quantity_rows"]
            print(f"{path.name}：保留 {len(frame):,} 筆；數量超出範圍忽略 {ignored:,} 筆", flush=True)
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as handle:
            temporary = Path(handle.name)
        workbook.save(temporary)
        temporary.replace(output)
    finally:
        for worksheet in workbook.worksheets:
            if not worksheet.closed:
                worksheet.close()
            if worksheet._writer is not None and Path(worksheet._writer.out).exists():
                worksheet._writer.cleanup()
        workbook.close()
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return counts


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", type=Path, default=BASE_DIR / "Data" / "UD", help="CSV 資料夾")
    parser.add_argument("--output", type=Path, default=BASE_DIR / "ud_screen.xlsx", help="輸出 Excel 路徑")
    args = parser.parse_args()
    try:
        counts = screen_to_excel(args.input_dir, args.output)
    except (OSError, ValueError) as exc:
        print(f"錯誤：{exc}", file=sys.stderr)
        return 1
    print(f"完成：{args.output.resolve()}（共 {sum(counts.values()):,} 筆）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

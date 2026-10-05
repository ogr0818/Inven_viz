"""驗證單筆數量，並排除超出允許範圍的資料。"""

from decimal import Decimal, InvalidOperation

import pandas as pd


MIN_QUANTITY = Decimal("-1000000")
MAX_QUANTITY = Decimal("1000000")


def filter_quantity_range(frame: pd.DataFrame, context: str) -> pd.DataFrame:
    """保留含邊界的數量範圍；以精確十進位比較，避免先轉型時溢位。"""
    def classify(value: object) -> int:
        try:
            text = str(value).strip()
            # Decimal 接受底線，但來源數量不接受這種寫法。
            if "_" in text:
                return 0
            quantity = Decimal(text)
        except (InvalidOperation, ValueError):
            return 0
        if not quantity.is_finite():
            return 0
        return 1 if MIN_QUANTITY <= quantity <= MAX_QUANTITY else 2

    status = frame["total_qty"].map(classify)
    invalid = status.eq(0)
    if invalid.any():
        indices = frame.index[invalid]
        rows = ", ".join(str(int(index) + 2) for index in indices[:5])
        suffix = " 等" if len(indices) > 5 else ""
        raise ValueError(f"{context}：第 {rows}{suffix} 筆資料列（含標題列）total_qty 不是有效數字")

    result = frame.loc[status.eq(1)].copy()
    # 排除極大值後重新轉換，確保負號運算不會使用 uint64。
    quantities = pd.to_numeric(result["total_qty"], errors="coerce")
    if quantities.isna().any():
        raise ValueError(f"{context}：total_qty 不是有效數字")
    result["total_qty"] = quantities
    result.attrs["ignored_quantity_rows"] = int(status.eq(2).sum())
    return result

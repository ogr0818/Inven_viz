# 進銷存視覺化圖表

## 使用 Streamlit 呈現

## 外部資料清理

執行 `uv run ud_screen.py`，依檔名順序讀取 `Data/UD` 的 UTF-16 或 UTF-8 CSV，
篩選 `phoutid == 5630`，並輸出至 `ud_screen.xlsx`，每個來源檔案各佔一個工作表。
`phtxid` 為偶數時，`total_qty` 使用負的絕對值；奇數使用正的絕對值。
UD 輸出欄位依序為 `drug_id`、`drug_name`、`total_qty`、`phtxid`、`phoutid`、`inv_year`、`inv_month`。
`inv_year`、`inv_month` 從工作表名稱取得，分別以 3 位及 2 位文字儲存，保留前導零；
例如工作表 `11501` 的所有資料列均填入 `"115"`、`"01"`。
工作表名稱須為 3 位數字年份接續 1–12 月（月份可為 1 或 2 位數字），無法解析時停止並保留既有輸出檔。

自訂路徑：

```sh
uv run ud_screen.py --input-dir /Data/UD --output ud_screen.xlsx
```

預設路徑以腳本所在目錄為準；命令列指定的相對路徑以執行時的目錄為準。
成功執行後會覆寫同名輸出檔；如果資料驗證失敗，既有輸出檔會保留。
OUT、UD 僅處理院區篩選後 `total_qty` 介於 -1,000,000 至 1,000,000（含邊界）的資料；
超出範圍會忽略整筆資料，並逐檔回報保留及忽略筆數。空白、非數字、NaN 與無限大會使驗證失敗。
數量範圍檢查後的保留資料若有空白 `drug_id`，或藥品代碼、藥名含 Excel 不允許的控制字元，
會停止處理並回報來源檔名、資料列及欄位；不會自動刪除字元或略過該筆資料。

執行 `uv run out_screen.py` 處理 `Data/OUT`，保留來源數量的正負號並輸出 `out_screen.xlsx`。
OUT 輸出依序包含 `drug_id`、`drug_name`、`total_qty`、`phoutid`、`inv_year`、`inv_month`。
年月從工作表名稱擷取，使用三位年份及兩位月份文字，例如 `11501` 對應 `"115"`、`"01"`；
月份可為一或兩位數字，須介於 1 至 12，名稱格式不符時停止處理並保留既有輸出。
執行 `uv run total_screen.py` 讀取 OUT、UD 輸出的所有工作表，按 `drug_id` 加總至 `total_screen.xlsx`。
彙總讀取時同樣檢查每筆數量範圍，逐表回報保留及忽略筆數；最終加總允許超過單筆數量範圍。

## 部署在 OpenAI Sites

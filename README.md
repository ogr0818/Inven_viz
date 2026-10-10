# 住院藥局藥品進耗存

正式網站：[住院藥局藥品進耗存](https://inven-viz.chi-nan.chatgpt.site)（公開免登入）。

本機入口為 `inventory.html`：以本機 `summary.db` 產生的獨立網頁，直接開啟即可查詢與下載，不需要 DB_Viz 資料夾或聊天。

## 每月資料處理：單一 Python 入口

首次建立已完成。將 `base.xlsx`、`月報表.xlsx`、`total_screen.xlsx` 放在專案根目錄，執行：

```sh
uv sync
uv run update_summary.py --dry-run
uv run update_summary.py
```

也可在已使用專案 Python 環境的編輯器直接執行 `update_summary.py`。不帶參數即更新既有 `summary.db`；預設檔案路徑以腳本所在資料夾為準，不受執行時工作目錄影響。找不到 DB 時停止，不自動建立新 DB。

### 輸入與更新規則

- `base.xlsx`：使用 `base` 工作表，必要欄位為 `drug_id`、`drug_name`、`drug_type`。
- `月報表.xlsx`：每張工作表名稱為三位民國年加一或兩位月份，例如 `11509`；必要欄位為 `drug_id`、`實發量`、`庫存量`。它決定本次更新的品項與月份。
- `total_screen.xlsx`：必要欄位為 `drug_id`、`total_qty`、`inv_year`、`inv_month`；依藥品代碼及年月合併，缺對應紀錄的住院耗用補 0。
- 月報與耗量檔常規各有一張工作表，亦支援多張；讀取全部工作表，不按工作表順序配對。耗量可在多張表分列同一月份，合計不得重複藥品代碼與年月。
- 月資料必須涵蓋欲更新月份的完整資料，不能只提供新增或更正的幾筆。完整替換月報涵蓋的月份，保留其他月份；重跑不追加重複紀錄。
- 實發量必須大於 0；非正實發量的整筆排除。住院耗用與庫存保留正負號。整月皆被排除，或有效年月表僅有標題時，仍清除 DB 該月份舊紀錄。
- 數量只接受 SQLite 範圍內的整數，不自動四捨五入。必要欄位的公式、Excel 錯誤、缺值、無效藥碼、無效年月及重複鍵均停止處理，回報來源位置。被排除的月報列仍須匹配來源基本資料與 DB 目前主檔。
- 所有來源先完成驗證，再以單次交易更新 DB。正式更新前自動備份至 DB 旁的 `backups/`；寫入失敗回復整次交易。`--dry-run` 不修改檔案、不建立備份。

此入口在記憶體完成彙整，不讀取或改寫既有 `summay.xlsx`，也不改寫三份來源 Excel。若仍需要獨立彙整 Excel，可另外執行 `uv run summay.py`。

### 主檔更新

平常執行只更新月資料，不因讀取 `base.xlsx` 而自動替換 DB 目前主檔。有新藥品時，先提供完整新版 `base.xlsx` 並明確執行：

```sh
uv run update_summary.py --replace-base --dry-run
uv run update_summary.py --replace-base
uv run update_summary.py
```

`--replace-base` 只替換目前主檔，不讀取月報及耗量檔；穩定藥碼與歷史月明細的名稱、劑型、數量保留。它不會連帶執行月資料更新。

### 備份還原

需要還原時，先停止其他資料庫讀寫程序並另存目前 DB 副本，再將選定的 `backups/*.db` 複製為 `summary.db`，執行 `uv run manage_summary.py info` 檢查。還原不會自動更新既有 HTML 或已發布網站，須重新產製及發布。

### 自訂路徑與移機重建

所有路徑可明確指定；自訂相對路徑以執行時目錄為準：

```sh
uv run update_summary.py --db summary.db --base base.xlsx --monthly 月報表.xlsx --total total_screen.xlsx
```

移機時攜帶現有 DB，或備妥完整歷史來源後，明確使用 `--init --db another_summary.db` 建立另一份 DB。首次建立拒絕覆寫既有路徑，不是日常更新指令；重建結果只涵蓋所提供的來源月份。

資料結構與歷史查詢見 [SCHEMA.md](SCHEMA.md)，來源整併及獨立運作驗證見 [Inven_DB 整併驗證](docs/Inven_DB整併驗證報告.md)。
Inven_DB 原始材料已完整保存於本機 `backups/inven-db-source-20261010.zip`，檔案清單與 SHA-256 位於同目錄的 `.manifest.json`；原始文件只供追溯。DB、來源 Excel、備份包不納入 Git，移機須另攜帶。原資料夾及對應專案由使用者手動封存。

## 新版網站與資料更新

在專案根目錄執行：

```sh
uv sync
uv run manage_summary.py info
uv run scripts/build_inventory.py
```

產製成功後開啟 `inventory.html`。後續資料與網站更新依序為：

```sh
# 三份 Excel 完整替換輸入月份，保留其他月份
uv run update_summary.py
uv run scripts/build_inventory.py
```

如需完整替換目前主檔，使用上述 `uv run update_summary.py --replace-base`；歷史月紀錄保留原名稱與劑型。舊 `manage_summary.py` 的維護介面仍保留，可用於直接匯入已彙整 Excel。
支援先加 `--dry-run` 檢查匯入或主檔更新；正式更新前自動備份至 `backups/`。網站更新仍須重新產生 HTML，資料庫變動不會自動反映於已產生或已發布的網頁。

可指定其他輸入與輸出：

```sh
uv run scripts/build_inventory.py --db summary.db --output inventory.html
```

實發量必須大於 0，整筆零值或負值月份排除；住院耗用與庫存保留原始正負號。空年月工作表可清除該月舊資料，表名須為三位民國年加一或兩位月份。產製檢查不通過時保留既有 HTML，且不修改來源 DB。
期間統計的名稱、劑型與庫存採區間內最近可納入月份，下載月明細則保留各月原值。決策見 [正實發量邊界](docs/adr/0001-positive-issued-months.md) 與 [歷史主檔](docs/adr/0002-historical-monthly-master.md)。

新版提供期間實發量上下限、E／I／O／S 劑型複選、共同「套用條件」與待套用提示。圖表、統計表與下載明細共用已套用篩選；圖表每張最多 100 品項，多張以 ZIP 下載。空結果點下載只顯示原因，不產生檔案。

## DB_Viz 整併與保存

程式與範本已納入 `scripts/`，正式入口為 `inventory.html`；舊 `inventory_trial_prototype.html`、`inventory_release_preview.html` 僅為歷史預覽。整併未復原工作區原本刪除的 SPEC.md，議題與規格繼續使用 GitHub Issues。

本機 `summary.db` 為來源資料庫的獨立副本；完整來源材料另存於 `backups/db-viz-source-20261010.zip`，含原規格、討論結論、程式、資料庫與舊驗證證據。此備份與 DB 為本機檔案，未納入 Git；若要移至其他電腦，須另行攜帶 DB 或從來源 Excel 重建。既有彙整 Excel 不在本次整併中重寫。

整併驗證見 [驗證報告](docs/整併驗證報告.md)。DB_Viz 資料夾刪除與聊天封存由使用者手動處理。

## 驗證

```sh
uv run python -m unittest discover -s tests -p 'test_*.py'
npm install
npx playwright install chromium
node tests/verify_inventory.cjs
node tests/verify_filters.cjs
uv run --with pillow python tests/verify_outputs.py
```

Python 測試亦涵蓋三份 Excel 到 DB 的單一入口，以及自 Inven_DB 納入的資料庫管理測試。

瀏覽器驗證使用目前隨整併帶入的 11501–11509、579 品項、3,400 筆資料作為全量驗收基準；未來改用其他資料時須同步更新驗收基準。`test_inventory_build.py` 使用獨立暫存資料驗證歷史主檔、零實發量與失敗保護，不依賴正式 DB。瀏覽器腳本可用 `CHROME_BIN` 指定已安裝 Chrome，用 `PYTHON` 指定 Python，用 `INVENTORY_URL` 指定要驗證的本機 HTTP 網址（預設驗證本機檔案）；下載與截圖證據產生於 `verification/`。Excel／PNG 回讀另需 Pillow，可用 `uv run --with pillow python tests/verify_outputs.py`。

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

## 部署

2026-10-10 公開部署前複查通過，詳見 [Sites 部署前檢查](docs/Sites部署前檢查.md)。以以下指令產生 Sites 靜態入口：

```sh
uv run scripts/build_inventory.py --output dist/index.html
```

公開內容僅為 `dist/index.html`（內嵌網站月明細），部署目錄指定 `dist`；不將專案根目錄作為公開目錄。DB、來源 Excel、備份與驗證輸出留在本機。這份網站無伺服器執行依賴，不需要線上 SQLite、Python 或 npm 套件。

正式 Sites 發布仍需建立 Site、保存回傳的 Site ID、同步專用來源並打包、將訪客權限設為公開，以及部署已保存版本。2026-10-10 已完成正式公開發布：[正式網址](https://inven-viz.chi-nan.chatgpt.site)。發布紀錄與後續更新方式見 [Sites 正式發布](docs/Sites正式發布.md)。

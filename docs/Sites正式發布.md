# Sites 正式發布

2026-10-10 已正式發布第 1 版：[住院藥局藥品進耗存](https://inven-viz.chi-nan.chatgpt.site)。Sites 回報部署終態 `succeeded`，訪客權限為 `public`。

## 發布識別

- Site ID：`appgprj_6aca1f709c2c8191a02933ab4954b626`
- 版本 ID：`appgprj_6aca1f709c2c8191a02933ab4954b626~appgver_ea48cb8bcb688191bf0bfa60a18d1a34`
- 部署 ID：`appgdep_6aca1fc13dd081919038e3d7844b309e`
- 已同步來源 commit：`dcc908b1439041aab965b2e431acf42d9d9ecfb9`
- 部署 HTML SHA-256：`b742c6d5120f509adddc94e41c2db016ef086042101351be133229e9b5bcedd6`

## 來源與發布內容

專用來源儲存庫位於 `sites/inven-viz/`，Site 身分與靜態目錄設定保存在該目錄的 `.openai/hosting.json`；它與本資料整理專案的 Git 儲存庫分開管理，來源已由 Sites 官方 helper 同步。根目錄忽略此獨立儲存庫，避免誤加為嵌入式儲存庫。

封裝由已同步 commit 產生，內含 `.openai/hosting.json` 與 `dist/index.html`，無 DB、Excel、驗證輸出或來源備份。網站內嵌已驗收的 3,400 筆月明細、579 品項、11501–11509，支援篩選與三種下載。

## 後續更新

1. 先依本專案流程更新本機 `summary.db`。
2. 重新產生網頁並完成適用驗收；網站來源入口可用 `uv run scripts/build_inventory.py --output sites/inven-viz/dist/index.html` 產生。
3. 使用 Sites 技能，從專用來源目錄沿用上述 Site ID 與 `.openai/hosting.json`，取得新的短期憑證、同步來源、封裝、保存版本並發布。
4. 確認部署成功與正式網址。維持公開權限；本機 DB 變更不會自動更新已發布網站。

不要再次建立新 Site，也不要從資料整理專案根目錄執行網站來源同步。短期憑證不保存於檔案或文件。


## 正式環境驗證

未登入的 HTTP 請求取得 200 與 HTML，無登入跳轉；正式網頁內嵌資料與驗收版本一致。另以全新、未登入的 Chrome 工作階段確認預設最新月份 367 品項、套用全期間 579 品項及單品項查詢；實際下載統計表、月明細與 PNG。正式下載的 579 筆期間統計與 ACT03I 月明細已回讀，逐欄對照本機 SQLite 一致，PNG 格式有效。

正式瀏覽器證據保存於本機 `verification/sites-live-results.json`、`verification/sites-live.png` 與三份 `sites-live-*` 下載檔。平台回應可能加入託管程式，故正式 HTML 不以整檔位元組雜湊宣稱完全相同；內嵌業務資料已逐值核對。

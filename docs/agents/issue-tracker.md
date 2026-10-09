# Issue tracker: GitHub

本專案的議題與規格存放於 GitHub Issues，使用 gh CLI 操作。
執行時由 Git remote 確認目標儲存庫。

## 操作慣例

- 建立議題：gh issue create --title "..." --body-file <檔案>
- 讀取議題：gh issue view <編號> --comments
- 列出議題：gh issue list，依需求篩選狀態與標籤
- 新增留言：gh issue comment <編號> --body-file <檔案>
- 加入標籤：gh issue edit <編號> --add-label "..."
- 移除標籤：gh issue edit <編號> --remove-label "..."
- 關閉議題：gh issue close <編號>

多行內容先寫入暫存文字檔，再透過 --body-file 提交。

## Pull requests as a triage surface

**PRs as a request surface: no.**

## 技能指令對應

- 「發布至議題追蹤器」：建立 GitHub issue。
- 「取得相關 ticket」：讀取指定 issue 及留言。

## Wayfinding

- Map 使用單一 issue，標籤為 wayfinder:map。
- Child ticket 優先使用 GitHub sub-issue；無法使用時，
  在 map 的 task list 連結子議題，並於子議題標示 Part of #<map>。
- Ticket 類型使用 wayfinder:research、wayfinder:prototype、
  wayfinder:grilling 或 wayfinder:task。
- 阻擋關係優先使用原生 issue dependencies；無法使用時，
  在子議題標示 Blocked by: #<編號>。
- Frontier 為 map 中尚未關閉、無未完成阻擋且無負責人的子議題，
  依 map 順序選取。
- Claim 時指派負責人；完成時留下結果、關閉子議題，
  並在 map 的 Decisions-so-far 加入摘要與連結。

# AGENTS.md

本檔提供給自動化代理（與協作者）在此 repo 內進行修改時的共同約定與操作指引。

## 專案概覽

- **語言/平台**：Excel VBA（`*.bas` 模組）
- **目的（依現有註解推定）**：用巨集產生「新年度薪資明細」相關活頁簿/工作表資料。

## 主要檔案

- `Module5.bas`
  - VBA 模組（含中文巨集名稱），目前可見：
    - `Sub newsalarydetail()`: 目前是空的殼（僅註解）。
    - `Sub 巨集3()`: 以互動輸入/確認後，嘗試複製舊年度檔案並清理特定工作表/列資料。

## 開發/執行方式（Excel）

- **匯入模組**
  - Excel → `ALT+F11` 開啟 VBA 編輯器
  - `File > Import File...` 匯入 `Module5.bas`
- **執行巨集**
  - 回到 Excel → `開發人員` → `巨集` → 選擇 `巨集3`（或在 VBA 編輯器中直接執行）
- **安全建議**
  - 先用「複本」活頁簿測試；此類巨集可能會刪除工作表/列。
  - 若程式碼有 `Sheets.Delete` / `Rows(...).Delete`，務必在測試時開啟備份與版本控管。

## 重要注意事項（代理與協作者必讀）

- **避免不可逆操作**
  - 任何會刪除資料的流程，建議加入：
    - `Application.DisplayAlerts = False/True`（控制刪除確認視窗）
    - 清楚的二次確認對話框與乾跑（dry-run）選項（若適用）
- **不要在 `For Each sh In Worksheets` 迴圈中直接刪除工作表**
  - 這在 VBA 常導致迭代錯亂；建議改為先收集待刪清單再逐一刪除。
- **路徑/檔名處理**
  - `filePath` 建議明確指定並確保結尾有路徑分隔符號（Windows 通常是 `\`）。
  - `FileCopy`、`Workbooks.Open` 一律使用完整路徑以避免「目前工作目錄」造成誤判。
- **外部依賴**
  - 程式碼若呼叫 `FileExists(...)` 之類函式，請確認專案內確實存在其實作（目前 repo 只看到 `Module5.bas`，可能尚缺）。

## 程式碼風格與品質要求

- **一律加上 `Option Explicit`**（建議）
  - 並補齊所有變數宣告（避免 `Userdata` 等未宣告變數默認為 Variant）。
- **避免魔法字串**
  - 工作表名稱（如「行政總表」、「總表」等）建議集中常數化，並在修改時同步更新說明。
- **錯誤處理**
  - 與檔案/工作簿/工作表操作相關的程式碼，建議加入可讀的錯誤訊息與復原步驟（例如恢復 `DisplayAlerts`、關閉檔案不儲存等）。

## 版本控管（Git）

- **提交粒度**
  - 以「單一意圖」為單位提交（例如：只做路徑處理修正、只補齊缺失函式、只重構刪表流程）。
- **提交訊息**
  - 使用明確描述，例如：`docs: add AGENTS guidance for VBA macros`、`fix: avoid deleting sheets during iteration`。

## 測試建議（最低要求）

本 repo 未提供自動化測試；修改 VBA 時請至少完成：

- 在 Excel 匯入並能編譯（`Debug > Compile VBAProject` 無錯）
- 用測試檔案跑過主要流程（至少能到對話框、開檔/複製邏輯、目標工作表存在性檢查）
- 確認任何全域狀態都在結束前恢復（例如 `DisplayAlerts`、`ScreenUpdating`）


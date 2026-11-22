# 2027 GAS 程式整理與保留清單（第一輪｜唯讀盤點）

> 2026-11-22。依 GitHub 既有 `apps-script/Code.gs`、`manager-gas/Code.gs` 與使用者從 Codespaces 貼出的最新 `clasp clone` 檔案清單、函式清單及五份欄位工具全文整理。**尚未直接讀取 Codespaces 中的完整最新 GAS 備份**，因此不代表線上 GAS 已通過驗證。

## 安全原則

- **全部暫時保留，不刪檔、不改名、不搬動函式、不執行寫入函式。**
- GitHub 的 `apps-script/`、`manager-gas/` 是不同版本／用途；不可直接覆蓋目前線上 GAS。
- Codespaces `gas-backup-2027/` 內的 `.clasp.json` 與登入憑證不得提交到公開 GitHub。
- 目前正式觸發器有四項：`sj27TgPoll2027`、`sj27TgDaily2027`、`paymentEditHandler2027`、`overdueHourlyHandler2027`。其所在程式與相依函式必須先確認，不能刪。
- `doGet`、`doPost` 是網頁入口，不能隨意重複定義或改名。

## 最新 GAS 備份：12 個檔案（使用者貼出的 clasp clone 清單）

| 檔案 | 已知功能 | 目前處置 |
|---|---|---|
| `程式碼.js` | 訂單建立、產能、付款同步、匯款回報、逾期保護、管理快照等；含 `doPost` | **核心，保留** |
| `CalendarAPI.js` | `doGet` 日曆／公開查詢入口 | **核心，保留** |
| `Telegram2027.js` | Telegram 新訂單輪詢、每日製作量通知、觸發器 | **核心，保留** |
| `TelegramSetup.js` | Telegram Chat ID 設定、測試通知 | 保留，確認設定依賴 |
| `Manager2027.html` | 舊版 GAS 管理 APP 畫面，呼叫 `managerSnapshot2027` | 保留，未確認可移除 |
| `ColumnAudit2027.js` | `sj27AuditColumnsReadOnly` 欄位唯讀稽核 | 保留 |
| `ColumnStage2027.js` | 新版預覽工作表的預檢、建立、驗證 | 保留；`sj27BuildStage` 會寫入預覽表 |
| `StageValidation2027.js` | 預覽表下拉選單修復與檢查 | 保留；修復函式會寫入驗證規則 |
| `未命名.js` | `sj27DiagnoseStageDifferences`：比對原表／預覽表差異，唯讀 | 保留，可待確認後重新命名 |
| `2 個未命名的副本.js` | `sj27FixStageRemittanceDateText`：修正預覽表匯款日期型態 | 保留；會寫入預覽表 |
| `3 個未命名的副本.js` | `sj27RestoreStageDropdowns`：修復預覽表付款／銀行下拉選單 | 保留；會寫入驗證規則 |
| `appsscript.json` | GAS 專案設定及權限範圍 | **核心，保留** |

## 欄位整理工具的相依關係

- `sj27DiagnoseStageDifferences` 依賴 `ColumnStage2027.js` 的 `SJ27_STAGE_ID`、`SJ27_STAGE_SPECS`。
- `sj27PreviewStage` 與 `sj27VerifyStage` 為唯讀；`sj27BuildStage` 會建立影子工作表，且遇到既有影子表會停止。
- `sj27RepairAllStageDropdowns` 與 `sj27RestoreStageDropdowns` 功能有部分重疊，但規則覆蓋範圍不同，**不可直接視為重複廢檔**。
- `sj27FixStageRemittanceDateText` 僅針對「_新版預覽_門市自取」匯款日期文字型態，屬一次性修復工具。
- `sj27DiagnoseStageDifferences` 會逐格讀取資料；雖然不記錄客戶欄位內容，仍需注意執行時間與授權。

## 待核對風險（尚未修正）

1. GitHub 舊 `apps-script/Code.gs` 的 `PICKUP_HEADERS` 為 18 欄、`DELIVERY_HEADERS` 為 25 欄；最新備份 `ColumnStage2027.js` 的舊版來源規格則是自取 22 欄、宅配 25 欄，新版預覽規格是自取 23 欄、宅配 26 欄。**先確認目前線上實際表頭與運作中的程式版本**。
2. GitHub `manager-gas/Code.gs` 的 `markPickupCollected2027` 以固定第 14 欄判斷付款、第 17 欄修改訂單狀態；若直接套到 23 欄新版表，會讀錯或寫錯。不得直接部署。
3. GitHub 的 `apps-script/Code.gs` 和 `manager-gas/Code.gs` 都有 `managerSnapshot2027`；這是兩套程式的同名函式，不可不經確認就合併成同一 GAS 專案。
4. GitHub `manager-gas/Code.gs` 有 `doGet`，最新 GAS 備份 `CalendarAPI.js` 也有 `doGet`。若合併同專案會造成入口衝突。
5. `README.md` 尚記載舊版「一張訂單總表」方案，與目前自取／宅配分表不一致；只需日後更新文件，**不是改資料表的依據**。
6. 目前已知四個觸發器均與付款或通知有關，重構時應先保留函式名稱與既有觸發器。

## 第二輪建議

1. **只讀取**最新 GAS 備份 `程式碼.js`、`CalendarAPI.js`、`Telegram2027.js`、`Manager2027.html`、`ColumnAudit2027.js` 的完整內容，製作跨檔呼叫表。
2. 核對實際試算表表頭及現有觸發器，不執行任何建立、修復、同步、取消、通知函式。
3. 將每個檔案標為「運作核心／維護工具／歷史工具／待確認」，確認後再考慮改名、歸檔。
4. 如需更動程式，先在 TEST 備份分支修改並做唯讀驗證，**不直接 `clasp push`**。

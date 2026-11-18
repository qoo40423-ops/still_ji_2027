# 紀家烘焙坊｜2027 TEST 手機管理 GAS

此資料夾是**獨立的管理 Apps Script 專案**，不要合併到現有公開匯款回報的 Apps Script，避免公開 doPost 部署也能呼叫管理資料函式。

1. 以管理用 Google 帳號在 script.google.com 新建 Apps Script 專案。
2. 將 Code.gs 內容完整貼入新專案；新增 HTML 檔案「Manager2027」，貼入 Manager2027.html。
3. 部署 → 新增部署 → 網頁應用程式。執行身分選「我」，存取權限選「只有自己」。如果選項不存在，停止部署，不要改成「任何人」。
4. 用部署網址在 iPhone Safari 開啟、完成 Google 登入，確認只看得到 TEST 訂單；再用分享 → 加入主畫面。

此版本只讀取 TEST 試算表，不能修改訂單；公開匯款回報安全開關維持原樣。Google 帳號登入可能需要在 Safari 內完成，主畫面啟動的登入體驗依 iOS 與 Google 驗證流程而異。

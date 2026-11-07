# 方案 A 部署說明：LINE 增刪股票（Google 試算表 + Apps Script）

這一步讓你可以**直接在 LINE 打字增刪股票**。屬於第三階段，建議在推播功能（第一、二階段）跑順之後再做。

## 一、建立 Google 試算表

1. 到 [Google 試算表](https://sheets.google.com) 新建一個空白試算表，命名例如「美股台股清單」。
2. 把左下角分頁改名為 `watchlist`。
3. 第一列填入標題（五欄）：`market`　`code`　`yf_symbol`　`name`　`group`
4. 把 `watchlist.csv` 的內容貼進去當初始清單（可用「檔案 → 匯入 → 上傳」選 `watchlist.csv`，匯入位置選「取代目前工作表」）。

## 二、貼上 Apps Script

1. 在試算表上方選單：**擴充功能 → Apps Script**。
2. 把 `Code.gs` 的全部內容貼進去，存檔。
3. 左側 **專案設定（齒輪）→ 指令碼屬性 → 新增屬性**：
   - `LINE_CHANNEL_ACCESS_TOKEN` = 你的 Channel access token（與推播用的同一個）
   - （可選）`SHEET_NAME` = `watchlist`

## 三、部署成網頁應用程式（webhook）

1. 右上 **部署 → 新增部署作業 → 類型選「網頁應用程式」**。
2. 設定：**執行身分 = 我**；**誰可以存取 = 任何人**。
3. 按「部署」，第一次會要求授權，依指示允許。
4. 複製產生的 **網頁應用程式網址**（`https://script.google.com/macros/s/..../exec`）。

## 四、把網址設為 LINE 的 webhook

1. 到 [LINE Developers](https://developers.line.biz/) → 你的 channel → **Messaging API** 分頁。
2. **Webhook URL** 貼上剛剛的網址，按 **Update** 與 **Verify**（應顯示成功）。
3. 打開 **Use webhook**（啟用）。
4. 建議把 **Auto-reply messages（自動回覆）** 與 **Greeting messages** 關掉，避免官方罐頭訊息干擾。

## 五、測試

在 LINE 對你的官方帳號傳：

- `清單` → 會回目前清單
- `加 台股 2317`、`加 AMZN` → 會驗證代碼並新增
- `刪 2317` → 會移除

## 六、讓推播程式讀這張表

1. 試算表 **檔案 → 共用 → 發布到網路 → 選 `watchlist` 分頁、格式 CSV → 發布**，複製 CSV 網址。
2. 把該網址設為環境變數／GitHub Secret 的 `SHEET_CSV_URL`。
3. 之後推播程式就會讀這張表的最新清單（發布版 CSV 可能有數分鐘快取，增刪後下一次推播才必定反映）。

> 未設定 `SHEET_CSV_URL` 時，推播程式會改用本機 `watchlist.csv`，功能照常。

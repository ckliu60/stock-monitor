# 美股／台股監控推播系統 — 交接文檔（Session Handoff）

> 給下一個對話的接手者：你對本專案沒有任何記憶，請把本文件從頭讀完再動作（尤其第 5 節「踩過的坑」）。使用者在對話中被稱為「老爺」，全程使用**繁體中文**，偏好簡潔、誠實、先問清楚再動手。本專案有一條**硬規則：兩份規格文件經老爺簽核前，不撰寫任何程式。**
> 更新日期：2026-11-07　｜　狀態：老爺已簽核（指示「開始寫程式」）。第一階段推播程式已完成並 dry-run 驗證通過；第二、三階段檔案（GitHub Actions、方案 A Apps Script、README）也已建立，等待老爺實際設定金鑰與部署。

---

## 1. 我們在做什麼（任務）

做一支程式：**每週固定時間（台灣時間 週二、四、六 早上 06:30）自動抓美股與台股的真實股價，計算跨期漲跌、標示異動警示，整理成一則精簡訊息推到老爺的手機 LINE。** 另外要能**在 LINE 直接打字增刪追蹤的股票**。

核心規格（已與老爺確認定案）：

- **數據源**：Yahoo Finance，透過 Python 套件 `yfinance`。只用真實收盤價，**拒絕任何模擬／推算值**。
- **四大指數（固定、不可增刪）**：道瓊 `^DJI` → 那斯達克 `^IXIC` → S&P500 `^GSPC` → 費半 `^SOX`（排序固定）。
- **美股核心 01–07**：`TSM`、`NVDA`、`INTC`、`MRVL`、`PLTR`、`TSLA`、`SPCX`。
- **美股觀察 08–13**：`AMD`、`MSFT`、`GOOG`、`META`、`NOK`、`BTC-USD`。
- **台股（v2 新增）**：`2330`台積電、`2454`聯發科、`0050`元大台灣50、`0052`富邦科技、`00919`群益台灣精選高息（給 yfinance 時一律加 `.TW`）。
- **跨期區間**：最新收盤 vs 1 週前（5 交易日）、1 月前（21 交易日）、1 年前（252 交易日），算報酬率 %。
- **階梯式警示（依報酬率「絕對幅度」，漲跌皆適用）**：`>10%` ⚠️、`>20%` 🚨、`>30%` ⚡、`>50%` 💥。**漲跌都要報，但下跌特別標示（▼ 與跌幅），目的是方便老爺低檔加碼（逢低買進）。** 漲用 ▲、跌用 ▼。
- **訊息排版**：純文字極簡三段式——①【大盤現況】四大指數；②【異動警示】達門檻清單（無則「今日無觸發警示門檻」）；③【個股數據】各檔週/月/年報酬（美股、台股分段）。不用 Flex 卡片。
- **推播管道**：LINE Messaging API 的 Push Message。
- **排程**：GitHub Actions cron `30 22 * * 1,3,5`（UTC）＝台灣週二四六 06:30。老爺明確選 GitHub Actions（電腦免開機、免費），**不走** Claude 雲端排程（會耗 Claude 用量且仍需 GitHub）。
- **LINE 增刪股票（v2 新增）**：採**方案 A** = Google 試算表存清單 + Google Apps Script 當 LINE webhook 接收端。指令如 `加 台股 2330`、`加 NVDA`、`刪 2330`、`清單`；新增前先驗證代碼有效。分階段交付（先推播、後互動）。

---

## 2. 目前進度

本專案到目前為止**只做了需求訪談與規格文件，尚未寫任何程式**。

**訪談過程（重要）**：老爺一開始貼了一份美股推播 PRD，但隨即要求「先不要寫程式、先把需求問清楚」。於是逐題確認，定案了上述規格。老爺**偏好直接打字回覆，不愛點選選項**（用 `AskUserQuestion` 工具出的選項題，老爺兩次都改用打字回，工具回報「user sent a message instead」）→ 之後一律用純文字問答。

**已驗證的事（實測，非推測）**：

- 用 yfinance 實際抓取，以下皆回傳真實資料（`period='1mo'`）：
  - `SPCX` 123.62（2026-11-05）、`TSM` 353.10、`^DJI` 51,285.73、`^SOX` 9,390.25、`BTC-USD` 75,773.23。
  - 台股：`2330.TW` 2,135.00、`2454.TW` 3,390.00、`0050.TW` 92.30、`0052.TW` 55.90、`00919.TW` 31.19（2026-11-05/06）。
- **SPCX 確認是 SpaceX**（詳見第 5 節坑）。

**文件版本演進**：

- **v1**（只有美股）：產出兩份 .docx——《程式需求規格書 v1》《程式開發設計書 v1》。檔名前綴 `美股監控推播系統_`。
- **v2**（現行，取代 v1）：老爺要求加台股 5 檔 + LINE 增刪功能。重寫為兩份——《程式需求規格書 v2》《程式開發設計書 v2》。檔名前綴 `美股台股監控推播系統_`。v2 變更：①加台股；②加 LINE 增刪（方案 A）；③清單由「寫死在程式」改為「存 Google 試算表」，推播流與互動流共用；④開發重排為三階段。

**程式開發（老爺簽核後完成）**：第一階段推播程式 `stock_monitor_line.py` 已寫好，以真實數據 `--dry-run` 驗證通過——四大指數、13 檔美股、5 檔台股的週/月/年報酬與警示符號都正確，並修掉了 0052 的分割假訊號（見坑 5.9）。第二、三階段的檔案（GitHub Actions workflow、`run_job.bat`、方案 A 的 `apps_script/`、`README.md`、`requirements.txt`、`.env.example`、`.gitignore`、`watchlist.csv`）也都建好了，`py_compile` 通過。尚未實際接 LINE／部署雲端／啟用增刪（都需老爺設定金鑰與帳號）。

**目前資料夾內的檔案**（`C:\Users\peril\OneDrive\圖片\桌面\Claude AI\美股自動通知\`）：

- 程式與設定：`stock_monitor_line.py`、`watchlist.csv`、`requirements.txt`、`.env.example`、`.gitignore`、`run_job.bat`、`README.md`
- 雲端排程：`.github/workflows/stock_monitor.yml`
- 方案 A：`apps_script/Code.gs`、`apps_script/README.md`
- 規格文件（現行）：`美股台股監控推播系統_程式需求規格書_v2.docx`、`美股台股監控推播系統_程式開發設計書_v2.docx`
- `HANDOFF.md`（本檔）
- 註：v1 舊 docx 已不在資料夾（應已被移除）。

---

## 3. 目前卡在哪

**沒有技術性卡關。** 第一階段推播程式已完成並以真實數據 dry-run 驗證通過（含分割修復，見坑 5.9）。第二、三階段的檔案也已建立。目前的「阻擋點」純屬待老爺動手設定的外部步驟：

- 取得並填入 LINE 金鑰（`LINE_CHANNEL_ACCESS_TOKEN`、`LINE_USER_ID`），做本機實推測試。
- 建 GitHub 私有 repo、設 Secrets、啟用 Actions（雲端排程）。
- （選做）建 Google 試算表 + 部署 Apps Script，啟用 LINE 增刪。

（v1 舊 docx 已不在資料夾，應已被移除。）

---

## 4. 下一步（依序）

程式都寫好了，剩下是陪老爺做設定。詳細照 `README.md` 的步驟一～五。

1. **本機實推**：協助老爺從 LINE Developers 後台取 token 與真正的 User ID（見坑 5.3），填入 `.env` 或環境變數，執行 `python stock_monitor_line.py` 確認手機收到。
2. **雲端排程**：建 GitHub 私有 repo，上傳全部檔案；在 repo Settings → Secrets 設 `LINE_CHANNEL_ACCESS_TOKEN`、`LINE_USER_ID`（選 `SHEET_CSV_URL`）；到 Actions 手動 Run workflow 驗證；之後 cron `30 22 * * 1,3,5`（UTC＝台灣二四六 06:30）自動跑。
3. **LINE 增刪（方案 A，選做）**：照 `apps_script/README.md`——建 Google 試算表（匯入 `watchlist.csv`）、貼 `apps_script/Code.gs`、部署成網頁應用程式、把網址設為 LINE webhook、把試算表發布為 CSV 設成 `SHEET_CSV_URL`。
4. 接手第一步：先確認老爺進行到哪一步、卡在哪，再往下帶，不要重做已完成的程式。

---

## 5. 踩過的坑 — 不要重蹈

**5.1 SPCX 是真股票，別再說它不存在。**
我（上一個對話）一開始憑記憶斷言「美股沒有 SPCX」，被老爺糾正。查證結果：**SPCX = SpaceX（Space Exploration Technologies），約 2026 年 6 月於 Nasdaq 上市**。模型知識截止在 2026-01，早於它 IPO，故記憶中沒有。**yfinance 抓得到真實資料**（123.62）。→ 規則：涉及 2026 年之後的股票／事件，**先用 WebSearch 或 yfinance 實查，不要憑記憶否定使用者**。SPCX 保留在清單。

**5.2 yfinance 的 Close 欄位要壓成 Series，否則 `float()` 會炸。**
單一代碼、`auto_adjust=True` 時，`df['Close']` 可能回傳 **DataFrame**（MultiIndex 欄位），直接 `float(series.iloc[-1])` 會報 `TypeError: float() argument must be a string or a real number, not 'Series'`。正確取法：
```python
c = df['Close']
if isinstance(c, pd.DataFrame):
    c = c.iloc[:, 0]
c = c.dropna()
```
這段驗證過可行，正式程式務必沿用。

**5.3 LINE 的「User ID」不是老爺給的 `ckliu60`。**
老爺說「有 ID 了：ckliu60」——那是 **LINE 個人／加好友 ID**，**不能**用於 Messaging API 推播。Push 的 `to` 需要 **API User ID**：`U` 開頭、約 33 碼字串（如 `U4af4980629...`），在 LINE Developers 後台「Basic settings」的 **Your user ID** 取得；另需「Messaging API」頁的 **Channel access token**。老爺的既有 channel（「之前程式設定過」）已存在，到設定階段用瀏覽器開 `developers.line.biz`（老爺已登入的話）協助抓這兩樣，**不要把金鑰打在對話裡**。另外：手機 LINE **必須先加該官方帳號為好友**，否則 push 會失敗。

**5.4 台股代碼要加 `.TW`。** `2330` → `2330.TW`。已實測 5 檔皆可抓。

**5.9 有些台股 ETF 的分割 Yahoo 沒收錄，`auto_adjust` 修不了，會產生假訊號。**
`0052.TW` 在 2025-11-17 有一次分割（單日比例 ≈ 0.143，即 1/7），但 `yf.Ticker('0052.TW').splits` 是 **0 筆**（Yahoo 沒這筆分割），所以 `auto_adjust=True` 無法還原，年報酬被算成假的「▼77%」（真實其實是 ▲58%，正負號還反了，會誤導「逢低加碼」判斷）。`repair=True` 需要 `sklearn` 且會報錯，不可靠。解法：程式內的 `repair_splits()` 會自掃單日跳動 ≤0.4 或 ≥2.5（一般股票/ETF 不可能的幅度）判為漏記分割並重新對齊歷史，閾值保守故不會誤改真實的 -40% 這類大跌。注意 2454 的「年 +161%」查無斷點、是**真實上漲**，別當成 bug 去「修」。

**5.5 老爺偏好打字、不愛點選項。** 別用 `AskUserQuestion` 出選項題（老爺會改用打字，工具收不到內容）。用純文字一題一題問。

**5.6 「LINE 增刪」需要 webhook 接收端，GitHub Actions 做不到。**
推播是單向、GitHub Actions 足夠；但「在 LINE 打字就改清單」需要**永遠在線的接收端**。GitHub Actions 只能定時跑、不能聽訊息。故採方案 A（Apps Script Web App 當 webhook）。別試圖用 GitHub Actions 接收 LINE 訊息。

**5.7 產 Word 檔用技能內建的 `build_docx.js`，別自己重寫排版。**
技能 `程式開發規格書` 位於 `C:\Users\peril\.claude\skills\程式開發規格書\`。做法：在暫存資料夾 `npm install docx`，寫內容腳本 `require("C:/Users/peril/.claude/skills/程式開發規格書/scripts/build_docx.js")`，用它匯出的 `titleBlock/H1/H2/P/Runs/table/RULE/makeDoc/ACCENT` 組章節；**執行時務必設** `NODE_PATH="<暫存>/node_modules"` 否則 build_docx.js 內部 `require('docx')` 找不到。表格欄寬用 DXA、各欄加總約 9000。**不要例行性轉 PDF**（本機預設印表機離線，LibreOffice 轉檔會跳印表機等待視窗；只有老爺要 PDF 預覽時才轉）。

**5.8 排程不要走 Claude 雲端排程。** 老爺比較過後選 GitHub Actions（免費、免開機、不耗 Claude 用量）。Claude 雲端排程會耗用量且通常仍需綁 GitHub repo。

---

## 6. 環境與限制

- **作業系統**：Windows 11 Enterprise。Shell 以 PowerShell 為主，另有 Bash（POSIX）。
- **Python**：3.12.10，路徑 `C:\Users\peril\AppData\Local\Programs\Python\Python312\python.exe`。已安裝 `yfinance`、`pandas`、`requests`。
- **Node**：v24.19.0、npm 11.17.0（產 Word 檔用）。
- **專案資料夾**：`C:\Users\peril\OneDrive\圖片\桌面\Claude AI\美股自動通知`（交付檔都放這）。
- **使用者**：稱「老爺」，繁體中文，偏好簡潔、誠實、先確認再做。email `ckliu60@gmail.com`。
- **交付格式**：規格文件一律 Word `.docx`（用上述技能產生）。
- **成本**：本案用量下，yfinance、LINE Messaging API、GitHub Actions、Google 試算表／Apps Script 皆免費。
- **硬規則（最重要）**：兩份規格文件（需求規格書、開發設計書）**經老爺簽核前，不寫任何程式**。每份文件結尾都載明此點。

---

*結尾提醒：此 HANDOFF.md **不會自動被下一個新對話看到**——請老爺在開新對話時，**手動把這個檔案（連同 v2 兩份 .docx 若需要）上傳進去**，並說「先讀這份交接文檔」。若 v2 文件未附，內容重點已濃縮在本檔第 1 節，可據此接續。*

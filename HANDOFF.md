# 美股／台股監控系統 — 交接文檔（Session Handoff）

> 給下一個對話的接手者：你對本專案沒有記憶，請先讀完本檔（尤其第 5 節「踩過的坑」與第 3 節「正在討論中的設計」）。使用者稱「老爺」，全程**繁體中文**，偏好簡潔、誠實、**先確認再動手、打字回覆（不要用 AskUserQuestion 選項卡）**。數據**真實性是最高原則**（絕不模擬，抓不到就說沒抓到）。
> 更新日期：2026-11-09　｜　狀態：核心系統（資料檔／Excel／Google Sheet 自動同步／歷史庫／LINE 文字推播）皆已完成。**正在討論：手機 LINE 通知的改版設計（塊1 大盤狀況、塊2 位階警示），尚未動手。**

---

## 1. 這個系統在做什麼

一支 Python 程式 `stock_monitor_line.py`，從 **Yahoo Finance（yfinance）抓真實股價**（**原始成交價、未還原配息，auto_adjust=False**，貼近看盤軟體），算出美股與台股大盤、個股的漲跌／均線／KD／位階，輸出多種檔案到 **Google 雲端硬碟資料夾**，並可推 LINE。

**監控標的**：指數 `^DJI`/`^IXIC`/`^GSPC`/`^SOX` + 台股大盤 `^TWII`；美股 `TSM NVDA INTC MRVL PLTR TSLA SPCX AMD MSFT GOOG META NOK BTC-USD`；台股 `2330 2454 0050 0052 00919`（台股加 `.TW`）。

**執行方式**：
- `python stock_monitor_line.py --html`：產報表＋所有資料檔（不推 LINE）
- `python stock_monitor_line.py --dry-run`：印 LINE 訊息不推
- `python stock_monitor_line.py`：推 LINE

---

## 2. 目前已完成

**產出的檔案（都在雲端資料夾 `G:\我的雲端硬碟\工作\AI 程式相關\美股投資`）：**
- `report.html`：電腦視覺報表（大盤狀況+燈號、個股表、超跌警示、中文新聞）。**注意：HTML 報表還是較舊版，沒有股價位階/20周位階/KD 欄，也沒依位階排序**（這些目前只在資料檔/Sheet）。
- `股價數據抓取_Yahoo.csv`：主資料表。欄位順序：**序號**(首欄,全列連號)、類別、市場、代碼、名稱、**股價位階**、**20周均線位階**、最新收盤、最新日期、當日%、週前收盤/日期/週%、月前.../月%、年前.../年%、MA10(雙週)/MA20(月)/MA60(季)/MA240(年)/MA20週、來源網址。
- `股價KD_Yahoo.csv`：KD(9,3,3)，**只有 0050 與 ^TWII** 的日/週/月 K、D。
- `股價數據抓取_Yahoo.xlsx`：**格式化 Excel**（位階上色、標題篩選、數值化、凍結、自動欄寬）。
- `歷史_日線.csv`/`歷史_週線.csv`/`歷史_月線.csv`：全標的歷史收盤寬表（資料庫，免重抓）。

**衍生指標**：當日/週(5)/月(21)/年(252) 報酬；MA10/20/60/240（個股＋指數都算）；20週均（**只 2330/2454/0050/^TWII 這 4 檔**，老爺明示維持 4 檔）；KD(9,3,3) 日/週/月（**只 0050/^TWII**）；股價位階（最新價相對 雙週MA10/月MA20/季MA60/年MA240：雙週線上→雙週~月→月~季→季~年→年線下）；20周均線位階（**線上/破線**）；分割自我修復（修 0052）。

**Google Sheet 自動同步（方案甲：Apps Script）**：
- Sheet「美股台股_股價數據抓取(Yahoo)」id `1vVKuff17l-xLoVLWmpEqZI-fu3Mpj9F5QJnA1lBEODo`。
- 綁定的 Apps Script（原始碼 `apps_script/sheet_sync.gs`）會讀雲端 CSV → 寫「主表」「KD」分頁，套格式（置中、表頭藍底粗體、凍結、**位階顏色填滿**、**標題列篩選**、最適欄寬），並刪多餘分頁。有每日 07:10 觸發器。老爺用 Sheet 上「📈 股價同步」選單手動同步。**改 Apps Script 需老爺重貼一次**（已重貼多次，盡量別再改）。

**LINE 文字推播**：`build_message()` 產三段式純文字（大盤/異動警示/個股 週月年），已實測成功。Channel = KENNY。

**關鍵識別碼：**
- 程式碼夾：`C:\Users\peril\OneDrive\圖片\桌面\Claude AI\美股自動通知`
- 雲端輸出夾：`G:\我的雲端硬碟\工作\AI 程式相關\美股投資`（Drive 資料夾 id `1RT3Kt8vma1aVWdbw3QnW6F02D8yv8zd8`，owner ckliu60@gmail.com）
- Google Sheet id：`1vVKuff17l-xLoVLWmpEqZI-fu3Mpj9F5QJnA1lBEODo`（存在 .env `STATS_SHEET_URL`）
- LINE：KENNY，bot `@887vlsxy`，推播 User ID `U661bab7536c87e60be76f4dd342df91e`，token 在 `.env`（勿外洩）。
- `.env` 變數：`LINE_CHANNEL_ACCESS_TOKEN`、`LINE_USER_ID`、`REPORT_OUTPUT_DIR`、`STATS_SHEET_URL`。

---

## 3. 正在討論中的設計（手機 LINE 通知改版）— 尚未動手

老爺要把手機通知做成「**簡短文字 + 連結/圖片**」，分兩塊。**已確認：用 20 周（非 24）、台積=台灣 2330、手機可用截圖(圖片)**。

**塊1：大盤狀況**（老爺給了 Excel 版型）— 兩段、各 4 列，欄位：
`NO｜指數/標的｜收盤｜當日｜燈號｜股價位階｜周K值｜周D值｜20周均位階`
- 〔美股〕道瓊、那斯達克、S&P500、費半
- 〔台股〕台灣加權大盤(^TWII)、台積電(2330)、0050、聯發科(2454)
- **需擴充**：目前 KD 只算 0050/^TWII；塊1 需要「週線 K/D」給全部 8 個標的（含美股四大指數、2330、2454）。

**塊2：位階警示**（含股價位階/KD/20周）：破季線、破年線、破 20 周、🟢K<20（低接/加碼）、🔴K>80（過熱/減碼）。

**手機送達方式的現實**：LINE 傳**圖片或連結**都需要「**公開網址**」（建議 GitHub Pages，順便做雲端排程→報表每天自動更新、網址固定）；**純文字訊息則免架設、現在就能傳**。→ 建議節奏：先做「文字版通知 + 升級 HTML（加位階/20周/KD/依位階排序）」，之後再做雲端(GitHub Pages)→連結/圖片自動接上。老爺尚未拍板此節奏。

**等老爺回答的 3 個塊1 問題**（我剛問、他還沒答）：
1. 周K值/周D值 = **週線** KD 對嗎？
2. **美股四大指數**要不要也算 20 周均位階（它們沒設過，但可算），還是留空？
3. 塊1 要同時用於「HTML 報表的大盤狀況」+「手機通知圖片」兩邊嗎？

---

## 4. 下一步（依序）

1. **接著談手機通知**：先收老爺對塊1 那 3 個問題的答案 → 定塊1；再談塊2（位階警示的門檻與格式）。**先不要寫程式，老爺要逐塊談清楚。**
2. 塊1/塊2 定案後：先做 **LINE 文字版**（免架設）＋**升級 report.html**（加股價位階、20周位階、KD 區塊，台/美股各自依位階由強到弱排序；footer 的 report_data.csv 連結要改成新檔名/Sheet）。
3. **雲端（GitHub Pages + GitHub Actions）**：報表固定網址＋每天自動更新 → LINE 才能傳連結/圖片。
4. **買賣提醒 + 資金配置**：KD 的 K≤20 買進/≥80 賣出提醒、連續天數；**資金配置比例由老爺給規則**，開發者只做機械式訊號（不提供個人投資建議，見坑 5.11）。
5. 接手第一步：先問老爺進行到哪、別重做已完成的。

---

## 5. 踩過的坑 — 不要重蹈

**5.1 SPCX 是真股票**（SpaceX，2026/6 Nasdaq 上市）。涉 2026 後標的先用 WebSearch/yfinance 實查，別憑記憶否定。

**5.2 yfinance Close 要壓成 Series**：`c=df['Close']; if isinstance(c,pd.DataFrame): c=c.iloc[:,0]`，否則 `float()` 報錯。

**5.3 LINE User ID ≠ 加好友 ID**：推播用 U 開頭（KENNY 的是 `U661...f91e`），不是 `ckliu60`。手機須先加 `@887vlsxy` 好友否則 push 失敗。

**5.4 台股加 `.TW`**；台股加權指數 = `^TWII`。

**5.5 老爺偏好打字、不用選項卡**：別用 `AskUserQuestion`，純文字問。

**5.6 批次下載偶爾漏抓**（曾漏 INTC/00919）。`fetch_data` 已加「漏抓個別重試」。

**5.7 0052 等台股分割 Yahoo 沒收錄**，`repair_splits()` 會自掃單日跳動 ≤0.4 或 ≥2.5 修正。2454「年 +155~161%」是真實上漲、非 bug。

**5.8 全部用原始價（auto_adjust=False）**：均線、KD、週/月/年報酬一律原始成交價（未還原配息），符合元大/富邦看盤軟體行情頁慣例（老爺拍板）。只有庫存損益才含息，本專案未做。

**5.9 KD(9,3,3)**：種子 50、原始 OHLC 計算，各家網站慣例不同可能差 1~2 點，屬真實計算非模擬。

**5.10 Google Sheet 改用 Apps Script 從 CSV 同步**（方案甲），不要再用連接器 `create_file` 重建 Sheet（它無法原地更新、會一直產生重複檔）。歷史上曾建了多張同名 Sheet，已清到只剩現行 `1vVKuff…`。`.gsheet` 捷徑檔無法用 cat/Read 讀，要清舊 Sheet 用 Drive 連接器 `search_files`（parentId 查）＋`trash_file`。

**5.11 不提供個人化投資建議**：買賣提醒的機械式訊號（K≤20/≥80）可做；「加碼幾成」要老爺給規則。

**5.12 CSV 被 Excel 鎖住**：老爺若把 `股價數據抓取_Yahoo.csv` 開在 Excel，程式寫檔會 `Permission denied`。請他先關 Excel 再跑。

**5.13 LINE 圖片/連結需公開網址**：LINE push 的 image 或連結都要 HTTPS 公開 URL（要架 GitHub Pages 或圖床）；純文字則免。

**5.14 LINE 官方帳號名稱改不了**（LINE 限制，後台無編輯欄）；大頭貼要老爺自己在「商家檔案→基本檔案圖片」上傳（選檔視窗無法代操作；照片已放雲端夾 `KENNY頭貼.png`）。

**5.15 產 Word 檔**用技能 `程式開發規格書` 的 `scripts/build_docx.js`：暫存夾 `npm install docx`，執行設 `NODE_PATH=<暫存>/node_modules`。不要例行轉 PDF（印表機彈窗）。

---

## 6. 環境與限制

- Windows 11；Python 3.12（已裝 yfinance/pandas/requests/python-dotenv/openpyxl）。Node v24（產 docx 用）。
- 連接器：Google Drive 已連線（建/刪 Sheet、查檔），帳號 ckliu60@gmail.com。
- 成本：yfinance、LINE、Google 試算表/Apps Script、GitHub 皆免費。
- 固定規則（也存在記憶 `data-table-format-rules`）：**資料表第一欄一律序號**、**數值用原始價**、**歷史盡量存分頁/CSV 當資料庫**。
- 硬原則：數據不可模擬；個人投資建議不代為提供。

---

*提醒：此 HANDOFF.md 不會自動被新對話看到——開新對話時請老爺手動上傳本檔並說「先讀這份交接文檔」。相關檔案（程式、報表、Sheet、apps_script）如未附，重點已濃縮於本檔。*

/**
 * 美股／台股監控推播系統 — LINE 增刪股票接收端（方案 A）
 * ---------------------------------------------------------------------------
 * 這段 Apps Script 綁在一個 Google 試算表上，部署成「網頁應用程式」當作 LINE webhook。
 * 使用者在 LINE 打指令，它就更新試算表並回覆確認。
 *
 * 試算表需有一個分頁（預設名稱 watchlist），第一列為標題：
 *   market | code | yf_symbol | name | group
 *
 * 指令：
 *   清單                 → 列出目前追蹤清單
 *   加 台股 2330         → 新增（可省略市場，純數字自動判為台股並加 .TW）
 *   加 NVDA              → 新增美股
 *   刪 2330 / 刪 NVDA    → 移除
 *   其他                 → 回覆使用說明
 *
 * 部署前請到「專案設定 → 指令碼屬性」新增：
 *   LINE_CHANNEL_ACCESS_TOKEN = 你的 Channel access token
 *   （可選）SHEET_NAME = 分頁名稱，預設 watchlist
 */

var YF_CHART = "https://query1.finance.yahoo.com/v8/finance/chart/";

function sheetName_() {
  return PropertiesService.getScriptProperties().getProperty("SHEET_NAME") || "watchlist";
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName_());
  if (!sh) {
    sh = ss.insertSheet(sheetName_());
    sh.appendRow(["market", "code", "yf_symbol", "name", "group"]);
  }
  return sh;
}

/** LINE webhook 入口 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    (body.events || []).forEach(function (ev) {
      if (ev.type === "message" && ev.message && ev.message.type === "text") {
        var reply = handleCommand_(ev.message.text.trim());
        replyLine_(ev.replyToken, reply);
      }
    });
  } catch (err) {
    // 回 200 給 LINE，避免它一直重送
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 給瀏覽器開網址驗證用 */
function doGet() {
  return ContentService.createTextOutput("stock-monitor webhook OK");
}

function handleCommand_(text) {
  var t = text.replace(/\s+/g, " ").trim();
  if (t === "清單" || t.toLowerCase() === "list") return listText_();

  var mAdd = t.match(/^(加|新增|add)\s+(.+)$/i);
  if (mAdd) return doAdd_(mAdd[2]);

  var mDel = t.match(/^(刪|刪除|移除|del|delete|remove)\s+(.+)$/i);
  if (mDel) return doRemove_(mDel[2]);

  return helpText_();
}

function helpText_() {
  return [
    "可用指令：",
    "• 清單",
    "• 加 台股 2330（或 加 2330）",
    "• 加 美股 NVDA（或 加 NVDA）",
    "• 刪 2330 / 刪 NVDA",
  ].join("\n");
}

/** 解析市場與代號，推導 yf_symbol */
function parseTarget_(arg) {
  var parts = arg.trim().split(" ");
  var market = "";
  var code = "";
  if (parts.length >= 2 && (parts[0] === "台股" || parts[0] === "美股")) {
    market = parts[0];
    code = parts[1];
  } else {
    code = parts[0];
  }
  code = code.toUpperCase();
  var yf = code;
  var isTw = market === "台股" || /^\d{3,6}[A-Z]?$/.test(code);
  if (isTw) {
    market = "台股";
    yf = /\.TW$/.test(code) ? code : code + ".TW";
  } else {
    market = market || "美股";
    yf = code;
  }
  return { market: market, code: code, yf: yf };
}

function doAdd_(arg) {
  var tgt = parseTarget_(arg);
  if (!tgt.code) return "格式有誤。\n" + helpText_();

  var sh = getSheet_();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).toUpperCase() === tgt.yf.toUpperCase()) {
      return "⚠️ " + tgt.code + " 已在清單中。";
    }
  }
  var v = validateSymbol_(tgt.yf);
  if (!v.ok) return "❌ 查無此代碼：" + tgt.code + "（" + tgt.yf + "）";
  sh.appendRow([tgt.market, tgt.code, tgt.yf, v.name || tgt.code, "觀察"]);
  return "✅ 已新增 " + tgt.market + " " + tgt.code +
    (v.name ? "（" + v.name + "）" : "") + "，現價 " + v.price + "。";
}

function doRemove_(arg) {
  var code = arg.trim().split(" ").pop().toUpperCase();
  var sh = getSheet_();
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var rowCode = String(data[i][1]).toUpperCase();
    var rowYf = String(data[i][2]).toUpperCase();
    if (rowCode === code || rowYf === code || rowYf === code + ".TW") {
      sh.deleteRow(i + 1);
      return "✅ 已移除 " + data[i][1] + "。";
    }
  }
  return "⚠️ 清單中找不到 " + code + "。";
}

function listText_() {
  var sh = getSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return "目前清單是空的。";
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var g = data[i][0] || "其他";
    (groups[g] = groups[g] || []).push(data[i][1]);
  }
  var out = ["目前追蹤清單："];
  Object.keys(groups).forEach(function (g) {
    out.push("—" + g + "—");
    out.push(groups[g].join("、"));
  });
  return out.join("\n");
}

/** 以 Yahoo chart API 驗證代碼是否存在，順便取名稱與現價 */
function validateSymbol_(symbol) {
  try {
    var url = YF_CHART + encodeURIComponent(symbol) + "?range=5d&interval=1d";
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return { ok: false };
    var j = JSON.parse(resp.getContentText());
    var r = j && j.chart && j.chart.result && j.chart.result[0];
    if (!r || !r.meta) return { ok: false };
    var price = r.meta.regularMarketPrice;
    if (price === undefined || price === null) return { ok: false };
    var name = r.meta.shortName || r.meta.longName || "";
    return { ok: true, name: name, price: price };
  } catch (err) {
    return { ok: false };
  }
}

function replyLine_(replyToken, text) {
  var token = PropertiesService.getScriptProperties()
    .getProperty("LINE_CHANNEL_ACCESS_TOKEN");
  if (!token || !replyToken) return;
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: "text", text: text }],
    }),
  });
}

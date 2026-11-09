/**
 * 美股台股_股價數據抓取(Yahoo) — 自動同步 Apps Script（方案甲）v2
 * ---------------------------------------------------------------------------
 * 作用：把雲端「美股投資」資料夾的 CSV 自動讀進本試算表的分頁，並套用
 *       置中對齊 + 最適欄寬 + 表頭粗體凍結 + 股價位階/20周位階 顏色填滿。
 *       同步時也會刪掉「主表」「KD」以外的多餘分頁。
 * 安裝：試算表 → 擴充功能 → Apps Script → 貼上本檔 → 存檔 → 執行 syncNow 授權。
 *       （可選）執行 setupDailyTrigger 設定每天 07:10 自動同步。
 */

var FOLDER_ID = "1RT3Kt8vma1aVWdbw3QnW6F02D8yv8zd8";   // 美股投資 資料夾
var SOURCES = [
  { file: "股價數據抓取_Yahoo.csv", tab: "主表" },
  { file: "股價KD_Yahoo.csv", tab: "KD" },
];

// 股價位階顏色（最健康綠 → 最差紅，顏色清淡）
var PRICE_LEVEL_COLORS = {
  "雙週線上": "#C6EFCE",
  "雙週~月": "#E2F0D9",
  "月~季": "#FFF2CC",
  "季~年": "#FCE4D6",
  "年線下": "#FFC7CE",
};
// 20周均線位階顏色
var MA20W_COLORS = { "線上": "#C6EFCE", "破線": "#FFC7CE" };

function syncNow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var done = [];
  SOURCES.forEach(function (src) {
    var rows = readCsv_(folder, src.file);
    if (!rows || rows.length === 0) return;
    writeTab_(ss, src.tab, rows);
    done.push(src.tab + "（" + (rows.length - 1) + "列）");
  });
  cleanupTabs_(ss);  // 刪掉主表/KD 以外的多餘分頁
  try { SpreadsheetApp.getActive().toast("已同步：" + done.join("、"), "完成", 6); } catch (e) {}
}

function readCsv_(folder, name) {
  var it = folder.getFilesByName(name);
  if (!it.hasNext()) return null;
  return Utilities.parseCsv(it.next().getBlob().getDataAsString("UTF-8").replace(/^﻿/, ""));
}

function writeTab_(ss, tabName, rows) {
  var sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sheet.clear();
  var cols = 0;
  rows.forEach(function (r) { cols = Math.max(cols, r.length); });
  var data = rows.map(function (r) {
    var x = r.slice();
    while (x.length < cols) x.push("");
    return x;
  });
  sheet.getRange(1, 1, data.length, cols).setValues(data);
  // 基本格式
  sheet.getRange(1, 1, data.length, cols)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(1, 1, 1, cols).setFontWeight("bold").setBackground("#D9E7FD");
  sheet.setFrozenRows(1);
  // 位階顏色填滿（只有主表有這兩欄）
  applyLevelColors_(sheet, data, "股價位階", PRICE_LEVEL_COLORS);
  applyLevelColors_(sheet, data, "20周均線位階", MA20W_COLORS);
  // 最適欄寬
  sheet.autoResizeColumns(1, cols);
  // 標題列篩選
  var existing = sheet.getFilter();
  if (existing) existing.remove();
  sheet.getRange(1, 1, data.length, cols).createFilter();
}

function applyLevelColors_(sheet, data, header, colorMap) {
  var col = data[0].indexOf(header);
  if (col < 0 || data.length < 2) return;
  var bg = [];
  for (var i = 1; i < data.length; i++) {
    var v = String(data[i][col] || "").trim();
    bg.push([colorMap[v] || null]);
  }
  sheet.getRange(2, col + 1, data.length - 1, 1).setBackgrounds(bg);
}

function cleanupTabs_(ss) {
  var keep = SOURCES.map(function (s) { return s.tab; });
  var sheets = ss.getSheets();
  if (sheets.length <= keep.length) return;
  sheets.forEach(function (sh) {
    if (keep.indexOf(sh.getName()) === -1 && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  });
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "syncNow") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("syncNow").timeBased().everyDays(1).atHour(7).nearMinute(10).create();
  try { SpreadsheetApp.getActive().toast("已設定每天 07:10 自動同步", "完成", 5); } catch (e) {}
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📈 股價同步")
    .addItem("立即同步", "syncNow")
    .addItem("設定每日自動同步（07:10）", "setupDailyTrigger")
    .addToUi();
}

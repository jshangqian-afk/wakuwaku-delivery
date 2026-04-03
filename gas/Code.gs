// === スプレッドシートID（デプロイ前に設定） ===
var SPREADSHEET_ID = "";

// === 定数 ===
var HISTORY_SHEET_NAME = "出荷履歴";
var PRODUCT_ORDER = [
  "あいかのキムチ210g",
  "匠220g",
  "匠大辛220g",
  "匠一本漬け",
  "匠大辛一本漬け"
];

// === 共通処理：データ受信→スプレッドシート記録 ===
function processDeliveryData(data) {
  Logger.log("データ受信: storeName=" + data.storeName + ", date=" + data.date + ", items数=" + (data.items ? data.items.length : 0));

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
    sheet.appendRow(["日付", "店舗名", "住所", "電話番号", "商品名", "数量", "単価", "金額"]);
    Logger.log("出荷履歴シートを新規作成");
  }

  var items = data.items || [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    sheet.appendRow([
      data.date,
      data.storeName,
      data.address || "",
      data.phone || "",
      item.productName,
      item.qty,
      item.price,
      item.qty * item.price
    ]);
    Logger.log("行追記: " + item.productName + " x" + item.qty);
  }

  Logger.log("処理完了: " + items.length + "件記録");
  return { status: "ok", count: items.length };
}

// === Webアプリ：GETリクエスト受信（CORS回避用） ===
function doGet(e) {
  var result;
  try {
    var jsonStr = e.parameter.data || "{}";
    Logger.log("doGet受信: " + jsonStr.substring(0, 200));
    var data = JSON.parse(jsonStr);
    result = processDeliveryData(data);
  } catch (err) {
    Logger.log("doGetエラー: " + err.message);
    result = { status: "error", message: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// === Webアプリ：POSTリクエスト受信（後方互換） ===
function doPost(e) {
  var result;
  try {
    var data = JSON.parse(e.postData.contents);
    Logger.log("doPost受信");
    result = processDeliveryData(data);
  } catch (err) {
    Logger.log("doPostエラー: " + err.message);
    result = { status: "error", message: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// === メニュー追加 ===
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("出荷管理")
    .addItem("出荷表を生成", "showMonthDialog")
    .addToUi();
}

// === 年月入力ダイアログ ===
function showMonthDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "出荷表を生成",
    "対象年月を入力してください（例：2026-04）",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  var input = result.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(input)) {
    ui.alert("エラー", "「YYYY-MM」の形式で入力してください（例：2026-04）", ui.ButtonSet.OK);
    return;
  }

  generateShippingReport(input);
  ui.alert("完了", input.replace("-", "年") + "月の出荷表を生成しました。", ui.ButtonSet.OK);
}

// === 出荷表生成 ===
function generateShippingReport(yearMonth) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!historySheet) {
    SpreadsheetApp.getUi().alert("エラー", "「出荷履歴」シートが見つかりません。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 履歴データ取得（ヘッダー除く）
  var lastRow = historySheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("エラー", "出荷履歴データがありません。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var data = historySheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // 該当月のデータを抽出
  var filtered = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[0].toString();
    // Date型の場合はYYYY-MM-DD文字列に変換
    if (row[0] instanceof Date) {
      dateVal = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    if (dateVal.substring(0, 7) === yearMonth) {
      filtered.push({
        date: dateVal,
        storeName: row[1].toString(),
        address: row[2].toString(),
        phone: row[3].toString(),
        productName: row[4].toString(),
        qty: Number(row[5]) || 0,
        price: Number(row[6]) || 0,
        amount: Number(row[7]) || 0
      });
    }
  }

  if (filtered.length === 0) {
    SpreadsheetApp.getUi().alert("該当月のデータがありません。");
    return;
  }

  // 店舗×商品ごとに数量を集計
  // storeMap: { storeName: { productName: qty, ... } }
  var storeMap = {};
  var storeOrder = [];
  for (var i = 0; i < filtered.length; i++) {
    var r = filtered[i];
    if (!storeMap[r.storeName]) {
      storeMap[r.storeName] = {};
      storeOrder.push(r.storeName);
    }
    if (!storeMap[r.storeName][r.productName]) {
      storeMap[r.storeName][r.productName] = 0;
    }
    storeMap[r.storeName][r.productName] += r.qty;
  }

  // 出荷表シート作成（既存なら上書き）
  var parts = yearMonth.split("-");
  var sheetName = parts[0] + "年" + parts[1] + "月_出荷表";
  var reportSheet = ss.getSheetByName(sheetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(sheetName);
  }

  // ヘッダー行
  var headerRow = ["店舗名"];
  for (var p = 0; p < PRODUCT_ORDER.length; p++) {
    headerRow.push(PRODUCT_ORDER[p]);
  }
  headerRow.push("合計数量");
  reportSheet.appendRow(headerRow);

  // ヘッダー書式
  var headerRange = reportSheet.getRange(1, 1, 1, headerRow.length);
  headerRange.setBackground("#FF6B35");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");

  // 総合計用
  var grandTotal = {};
  for (var p = 0; p < PRODUCT_ORDER.length; p++) {
    grandTotal[PRODUCT_ORDER[p]] = 0;
  }
  var grandTotalQty = 0;

  // 店舗ごとにデータ行を出力
  for (var s = 0; s < storeOrder.length; s++) {
    var storeName = storeOrder[s];
    var products = storeMap[storeName];
    var row = [storeName];
    var storeTotal = 0;

    for (var p = 0; p < PRODUCT_ORDER.length; p++) {
      var pName = PRODUCT_ORDER[p];
      var qty = products[pName] || 0;
      row.push(qty > 0 ? qty : "");
      grandTotal[pName] += qty;
      storeTotal += qty;
    }
    row.push(storeTotal);
    grandTotalQty += storeTotal;
    reportSheet.appendRow(row);
  }

  // 総合計行
  var totalRow = ["合計"];
  for (var p = 0; p < PRODUCT_ORDER.length; p++) {
    var qty = grandTotal[PRODUCT_ORDER[p]];
    totalRow.push(qty > 0 ? qty : "");
  }
  totalRow.push(grandTotalQty);
  reportSheet.appendRow(totalRow);

  // 総合計行の書式
  var lastRowNum = reportSheet.getLastRow();
  var totalRange = reportSheet.getRange(lastRowNum, 1, 1, headerRow.length);
  totalRange.setBackground("#FFF3EE");
  totalRange.setFontWeight("bold");

  // 列幅自動調整
  for (var c = 1; c <= headerRow.length; c++) {
    reportSheet.autoResizeColumn(c);
  }

  // 罫線
  var allRange = reportSheet.getRange(1, 1, lastRowNum, headerRow.length);
  allRange.setBorder(true, true, true, true, true, true);

  // 数量列を中央揃え
  if (lastRowNum > 1) {
    reportSheet.getRange(2, 2, lastRowNum - 1, PRODUCT_ORDER.length + 1)
      .setHorizontalAlignment("center");
  }
}

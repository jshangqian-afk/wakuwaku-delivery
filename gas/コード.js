// === スプレッドシートID ===
var SPREADSHEET_ID = "19vyOw7JDEhSSVeATes1w2aRiS1ssFzy4cn3gLhEuaKM";

// === GoogleドライブフォルダID ===
var DRIVE_FOLDER_ID = "1GJf6Bl8LL-HIYeYCvMJf12jWk308zSHH";

// === 定数 ===
var HISTORY_SHEET_NAME = "出荷履歴";
var PRODUCT_ORDER = [
  "あいかのキムチ210g",
  "匠220g",
  "匠大辛220g",
  "匠一本漬け",
  "匠大辛一本漬け"
];

// === Webアプリ：POSTリクエスト受信 ===
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(HISTORY_SHEET_NAME);
      sheet.appendRow(["日付", "店舗名", "住所", "電話番号", "商品名", "数量", "単価", "金額"]);
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
    }

    // PDF をGoogleドライブに保存（失敗しても続行）
    var driveStatus = "skipped";
    if (data.pdfBase64 && DRIVE_FOLDER_ID) {
      try {
        var pdfBlob = Utilities.newBlob(
          Utilities.base64Decode(data.pdfBase64),
          "application/pdf",
          "納品書_" + data.storeName + "_" + data.date + ".pdf"
        );
        var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        folder.createFile(pdfBlob);
        driveStatus = "ok";
      } catch (driveErr) {
        driveStatus = "error: " + driveErr.message;
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", drive: driveStatus }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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
    if (row[0] instanceof Date) {
      dateVal = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    if (dateVal.substring(0, 7) === yearMonth) {
      filtered.push({
        storeName: row[1].toString(),
        productName: row[4].toString(),
        qty: Number(row[5]) || 0,
        price: Number(row[6]) || 0
      });
    }
  }

  if (filtered.length === 0) {
    SpreadsheetApp.getUi().alert("該当月のデータがありません。");
    return;
  }

  // 商品列を動的に取得（PRODUCT_ORDERを優先順として使い、残りを後ろに追加）
  var productSet = {};
  var priceMap = {};
  for (var i = 0; i < filtered.length; i++) {
    var r = filtered[i];
    productSet[r.productName] = true;
    if (priceMap[r.productName] === undefined) {
      priceMap[r.productName] = r.price;
    }
  }
  var productColumns = [];
  for (var p = 0; p < PRODUCT_ORDER.length; p++) {
    if (productSet[PRODUCT_ORDER[p]]) {
      productColumns.push(PRODUCT_ORDER[p]);
    }
  }
  for (var name in productSet) {
    if (productColumns.indexOf(name) === -1) {
      productColumns.push(name);
    }
  }

  // 店舗ごとに数量を集計（日付は無視）
  var rowMap = {};
  var rowOrder = [];
  for (var i = 0; i < filtered.length; i++) {
    var r = filtered[i];
    var key = r.storeName;
    if (!rowMap[key]) {
      rowMap[key] = { storeName: r.storeName, products: {} };
      rowOrder.push(key);
    }
    if (!rowMap[key].products[r.productName]) {
      rowMap[key].products[r.productName] = 0;
    }
    rowMap[key].products[r.productName] += r.qty;
  }

  // 店舗名順にソート
  rowOrder.sort();

  // 出荷表シート作成（既存なら上書き）
  var parts = yearMonth.split("-");
  var sheetName = parts[0] + "年" + parts[1] + "月_出荷表";
  var reportSheet = ss.getSheetByName(sheetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(sheetName);
  }

  var numProducts = productColumns.length;

  // ヘッダー行: ["店舗名", 商品名..., "合計数量", "合計金額"]
  var headerRow = ["店舗名"];
  for (var p = 0; p < numProducts; p++) {
    headerRow.push(productColumns[p]);
  }
  headerRow.push("合計数量");
  headerRow.push("合計金額");
  reportSheet.appendRow(headerRow);

  // ヘッダー書式
  var headerRange = reportSheet.getRange(1, 1, 1, headerRow.length);
  headerRange.setBackground("#FF6B35");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");

  // 総合計用
  var grandTotal = {};
  for (var p = 0; p < numProducts; p++) {
    grandTotal[productColumns[p]] = 0;
  }
  var grandTotalQty = 0;
  var grandTotalAmount = 0;

  // 店舗ごとにデータ行を出力
  for (var i = 0; i < rowOrder.length; i++) {
    var entry = rowMap[rowOrder[i]];
    var row = [entry.storeName];
    var rowTotalQty = 0;
    var rowTotalAmount = 0;

    for (var p = 0; p < numProducts; p++) {
      var pName = productColumns[p];
      var qty = entry.products[pName] || 0;
      row.push(qty > 0 ? qty : "");
      grandTotal[pName] += qty;
      rowTotalQty += qty;
      rowTotalAmount += qty * (priceMap[pName] || 0);
    }
    row.push(rowTotalQty);
    row.push(rowTotalAmount);
    grandTotalQty += rowTotalQty;
    grandTotalAmount += rowTotalAmount;
    reportSheet.appendRow(row);
  }

  // 単価行: ["", "単価", 商品1の単価, 商品2の単価, ..., "", ""]
  // ※ 1列目が空、2列目が"単価"なので、商品列は3列目から → ヘッダーは["店舗名", 商品...]
  // 仕様に合わせて: ["", "単価", 単価1, 単価2, ..., "", ""]
  // ただしヘッダーが["店舗名", 商品1, 商品2, ..., "合計数量", "合計金額"]なので
  // 単価行の最初の要素は店舗名列に対応 → 仕様通り ""
  // → しかし仕様では3要素目から単価が始まる = 2列目が"単価"ラベル
  // ヘッダーは1列目=店舗名、2列目以降=商品名... なので
  // 単価行: 1列目="単価", 2列目以降=各商品の単価, 最後2列="",""
  var unitPriceRow = ["単価"];
  for (var p = 0; p < numProducts; p++) {
    unitPriceRow.push(priceMap[productColumns[p]] || 0);
  }
  unitPriceRow.push("");
  unitPriceRow.push("");
  reportSheet.appendRow(unitPriceRow);

  // 商品別金額行: ["商品別金額", 商品1の合計金額, 商品2の合計金額, ..., "", 総合計金額]
  var productAmountRow = ["商品別金額"];
  for (var p = 0; p < numProducts; p++) {
    var pName = productColumns[p];
    var amt = grandTotal[pName] * (priceMap[pName] || 0);
    productAmountRow.push(amt > 0 ? amt : "");
  }
  productAmountRow.push("");
  productAmountRow.push(grandTotalAmount);
  reportSheet.appendRow(productAmountRow);

  // 合計行: ["合計", 商品1の合計数量, 商品2の合計数量, ..., 総合計数量, 総合計金額]
  var totalRow = ["合計"];
  for (var p = 0; p < numProducts; p++) {
    var qty = grandTotal[productColumns[p]];
    totalRow.push(qty > 0 ? qty : "");
  }
  totalRow.push(grandTotalQty);
  totalRow.push(grandTotalAmount);
  reportSheet.appendRow(totalRow);

  // 書式設定
  var lastRowNum = reportSheet.getLastRow();
  var totalCols = headerRow.length;

  // 合計行の書式（最終行）
  var totalRange = reportSheet.getRange(lastRowNum, 1, 1, totalCols);
  totalRange.setBackground("#FFF3EE");
  totalRange.setFontWeight("bold");

  // 単価行・商品別金額行の書式
  var unitPriceRowNum = lastRowNum - 2;
  var productAmountRowNum = lastRowNum - 1;
  reportSheet.getRange(unitPriceRowNum, 1, 1, totalCols).setBackground("#F5F5F5").setFontWeight("bold");
  reportSheet.getRange(productAmountRowNum, 1, 1, totalCols).setBackground("#F5F5F5").setFontWeight("bold");

  // 金額列（合計金額列）を#,##0形式に設定
  var amountColIndex = totalCols; // 最後の列が合計金額
  reportSheet.getRange(2, amountColIndex, lastRowNum - 1, 1).setNumberFormat("#,##0");

  // 商品別金額行の商品金額セルも#,##0形式
  reportSheet.getRange(productAmountRowNum, 2, 1, numProducts).setNumberFormat("#,##0");

  // 数量列を中央揃え（商品列＋合計数量列）
  if (lastRowNum > 1) {
    reportSheet.getRange(2, 2, lastRowNum - 1, numProducts + 1)
      .setHorizontalAlignment("center");
  }

  // 列幅自動調整
  for (var c = 1; c <= totalCols; c++) {
    reportSheet.autoResizeColumn(c);
  }

  // 全セルに罫線
  var allRange = reportSheet.getRange(1, 1, lastRowNum, totalCols);
  allRange.setBorder(true, true, true, true, true, true);
}

// === テスト用関数 ===
function testDriveSave() {
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var file = folder.createFile('test.txt', 'テスト保存', MimeType.PLAIN_TEXT);
  Logger.log('保存成功: ' + file.getName());
}

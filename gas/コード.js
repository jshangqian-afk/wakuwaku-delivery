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

    // action分岐: 削除リクエスト
    if (data.action === "delete") {
      return deleteHistoryByUniqueId(data.uniqueId);
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(HISTORY_SHEET_NAME);
      sheet.appendRow(["uniqueID", "日付", "店舗名", "住所", "電話番号", "商品名", "数量", "単価", "金額"]);
    }

    // uniqueIDをフロントから受け取る（delivery.idを流用）。なければGAS側で生成
    var uniqueId = data.uniqueId || Utilities.getUuid();

    // 冪等性チェック: 既に同一uniqueIdが存在すれば追記しない（再出力時の重複防止）
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var alreadyExists = existingIds.some(function(r) { return r[0] === uniqueId; });
      if (alreadyExists) {
        var driveStatusSkip = "skipped";
        if (data.pdfBase64 && DRIVE_FOLDER_ID) {
          try {
            var pdfBlobSkip = Utilities.newBlob(
              Utilities.base64Decode(data.pdfBase64),
              "application/pdf",
              "納品書_" + data.storeName + "_" + data.date + "_" + uniqueId.substring(0, 8) + "_reprint.pdf"
            );
            DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(pdfBlobSkip);
            driveStatusSkip = "ok";
          } catch (driveErr) {
            driveStatusSkip = "error: " + driveErr.message;
          }
        }
        return ContentService
          .createTextOutput(JSON.stringify({ status: "ok", skipped: true, uniqueId: uniqueId, drive: driveStatusSkip }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    var items = data.items || [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      sheet.appendRow([
        uniqueId,
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
          "納品書_" + data.storeName + "_" + data.date + "_" + uniqueId.substring(0, 8) + ".pdf"
        );
        var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        folder.createFile(pdfBlob);
        driveStatus = "ok";
      } catch (driveErr) {
        driveStatus = "error: " + driveErr.message;
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", uniqueId: uniqueId, drive: driveStatus }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// === uniqueIDで履歴を削除 ===
function deleteHistoryByUniqueId(uniqueId) {
  if (!uniqueId) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: "uniqueId is required" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: "シートが見つかりません" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", deleted: 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var deleteCount = 0;

  // 下から削除（行番号がずれないように）
  for (var i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === uniqueId) {
      sheet.deleteRow(i + 2);
      deleteCount++;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", deleted: deleteCount }))
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
  // uniqueID列追加に伴い9列に拡張（row[0]=uniqueID）
  var data = historySheet.getRange(2, 1, lastRow - 1, 9).getValues();

  // 該当月のデータを抽出
  var filtered = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[1].toString();
    if (row[1] instanceof Date) {
      dateVal = Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    if (dateVal.substring(0, 7) === yearMonth) {
      filtered.push({
        storeName: row[2].toString(),
        productName: row[5].toString(),
        qty: Number(row[6]) || 0,
        price: Number(row[7]) || 0
      });
    }
  }

  if (filtered.length === 0) {
    SpreadsheetApp.getUi().alert("該当月のデータがありません。");
    return;
  }

  // 商品列を動的に取得（PRODUCT_ORDERを優先順として使い、残りを後ろに追加）
  // priceMap: { 商品名: [登場した全単価の配列] } — 月内で複数単価が出るケースに対応
  var productSet = {};
  var priceMap = {};
  for (var i = 0; i < filtered.length; i++) {
    var r = filtered[i];
    productSet[r.productName] = true;
    if (!priceMap[r.productName]) {
      priceMap[r.productName] = [];
    }
    if (priceMap[r.productName].indexOf(r.price) === -1) {
      priceMap[r.productName].push(r.price);
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

  // 店舗ごとに数量と金額を集計（金額は実単価ベースで積み上げる）
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
      rowMap[key].products[r.productName] = { qty: 0, amount: 0 };
    }
    rowMap[key].products[r.productName].qty += r.qty;
    rowMap[key].products[r.productName].amount += r.qty * r.price;
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

  // 総合計用（数量と金額を別々に集計）
  var grandTotalQtyByProduct = {};
  var grandTotalAmountByProduct = {};
  for (var p = 0; p < numProducts; p++) {
    grandTotalQtyByProduct[productColumns[p]] = 0;
    grandTotalAmountByProduct[productColumns[p]] = 0;
  }
  var grandTotalQty = 0;
  var grandTotalAmount = 0;

  // 店舗ごとにデータ行を出力（金額は実単価ベースの実額をそのまま使う）
  for (var i = 0; i < rowOrder.length; i++) {
    var entry = rowMap[rowOrder[i]];
    var row = [entry.storeName];
    var rowTotalQty = 0;
    var rowTotalAmount = 0;

    for (var p = 0; p < numProducts; p++) {
      var pName = productColumns[p];
      var prod = entry.products[pName] || { qty: 0, amount: 0 };
      var qty = prod.qty;
      var amount = prod.amount;

      row.push(qty > 0 ? qty : "");
      grandTotalQtyByProduct[pName] += qty;
      grandTotalAmountByProduct[pName] += amount;
      rowTotalQty += qty;
      rowTotalAmount += amount;
    }
    row.push(rowTotalQty);
    row.push(rowTotalAmount);
    grandTotalQty += rowTotalQty;
    grandTotalAmount += rowTotalAmount;
    reportSheet.appendRow(row);
  }

  // 単価行: 1列目="単価"、2列目以降=各商品の単価（複数あれば昇順スラッシュ併記）、最後2列=""
  var unitPriceRow = ["単価"];
  for (var p = 0; p < numProducts; p++) {
    var prices = priceMap[productColumns[p]] || [];
    prices.sort(function(a, b) { return a - b; });
    unitPriceRow.push(prices.length > 0 ? prices.join("/") : 0);
  }
  unitPriceRow.push("");
  unitPriceRow.push("");
  reportSheet.appendRow(unitPriceRow);

  // 商品別金額行: 実集計値をそのまま使う（priceMapで再計算しない）
  var productAmountRow = ["商品別金額"];
  for (var p = 0; p < numProducts; p++) {
    var pName = productColumns[p];
    var amt = grandTotalAmountByProduct[pName];
    productAmountRow.push(amt > 0 ? amt : "");
  }
  productAmountRow.push("");
  productAmountRow.push(grandTotalAmount);
  reportSheet.appendRow(productAmountRow);

  // 合計行: ["合計", 商品1の合計数量, ..., 総合計数量, 総合計金額]
  var totalRow = ["合計"];
  for (var p = 0; p < numProducts; p++) {
    var qty = grandTotalQtyByProduct[productColumns[p]];
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

// === バックフィル関数（既存データのuniqueID列を埋める / 1回限り実行）===
// 検証環境で動作確認後、本番でも1回だけ手動実行する
function backfillUniqueIds() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    Logger.log("シートが見つかりません");
    return;
  }
  var n = sheet.getLastRow() - 1;
  if (n < 1) {
    Logger.log("バックフィル対象なし");
    return;
  }
  var ids = sheet.getRange(2, 1, n, 1).getValues();
  var filledCount = 0;
  for (var i = 0; i < n; i++) {
    if (!ids[i][0]) {
      sheet.getRange(i + 2, 1).setValue("legacy-" + Utilities.getUuid());
      filledCount++;
    }
  }
  Logger.log("バックフィル完了: " + filledCount + " 行");
}

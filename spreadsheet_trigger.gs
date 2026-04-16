// ============================================================
//  利益シミュレーション 自動実行スクリプト
//  設置場所: スプレッドシート > 拡張機能 > Apps Script
// ============================================================

// ▼ デプロイ後のサーバーURLに書き換えてください
var SERVER_URL = "https://estimation-analysis-app-219333891042.asia-northeast1.run.app/trigger/spreadsheet";

// ▼ 各列の位置（1始まり）※シートの列順に合わせて変更可
var COL_JAN      = 1;  // A列: JANコード
var COL_NAME     = 2;  // B列: 商品名
var COL_QTY      = 3;  // C列: 数量 / 在庫
var COL_COST     = 4;  // D列: 仕入れ価格 / 下代

/**
 * 「分析実行」ボタン or メニューから呼ばれる関数。
 * 現在開いているシートのデータをサーバーへ送信する。
 */
function runAnalysis() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getName();

  var response = ui.alert(
    "利益シミュレーション実行",
    "【" + sheetName + "】のデータを分析します。\nよろしいですか？",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  // データ行を収集（1行目はヘッダーとしてスキップ）
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("情報", "データがありません。", ui.ButtonSet.OK);
    return;
  }

  var items = [];
  for (var i = 2; i <= lastRow; i++) {
    var jan  = String(sheet.getRange(i, COL_JAN).getValue()).trim();
    if (!jan || jan === "") continue;  // JANコードが空の行はスキップ

    var name = String(sheet.getRange(i, COL_NAME).getValue()).trim() || "Unknown";
    var qty  = String(sheet.getRange(i, COL_QTY).getValue()).trim();

    // 仕入れ価格: ¥記号・カンマを除去して数値化
    var costRaw = String(sheet.getRange(i, COL_COST).getValue()).trim();
    var cost = parseFloat(costRaw.replace(/[¥,\s]/g, "")) || 0;

    items.push({
      jan_code:     jan,
      product_name: name,
      quantity:     qty,
      cost:         cost,
      sender_name:  sheetName,  // どのシートから送信したか記録
      row_index:    i,           // ステータス更新のための行番号
      sheet_name:   sheetName    // ステータス更新のためのシート名
    });
  }

  if (items.length === 0) {
    ui.alert("情報", "JANコードが入力された行がありません。", ui.ButtonSet.OK);
    return;
  }

  // サーバーへ送信
  try {
    var result = UrlFetchApp.fetch(SERVER_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ items: items }),
      muteHttpExceptions: true
    });

    var body = JSON.parse(result.getContentText());

    if (body.queued > 0) {
      ui.alert(
        "実行開始",
        body.queued + " 件の分析を開始しました。\n数分後に Analysis シートを確認してください。",
        ui.ButtonSet.OK
      );
    } else {
      ui.alert("情報", body.message || "処理対象がありません。", ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert("エラー", "サーバーへの接続に失敗しました。\n" + e.message, ui.ButtonSet.OK);
  }
}

/**
 * スプレッドシート起動時にメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("利益シミュレーション")
    .addItem("▶ 現在のシートを分析", "runAnalysis")
    .addToUi();
}

// ============================================================
//  利益シミュレーション 自動実行スクリプト
//  設置場所: スプレッドシート > 拡張機能 > Apps Script
// ============================================================

// ▼ デプロイ後のサーバーURLに書き換えてください
var SERVER_URL = "https://YOUR_SERVER_URL/trigger/spreadsheet";

/**
 * 「分析実行」ボタンをクリックしたときに呼ばれる関数。
 * ボタンには「runAnalysis」を割り当てる。
 */
function runAnalysis() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    "利益シミュレーション実行",
    "入力シートの未処理行を分析します。\nよろしいですか？",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    var result = UrlFetchApp.fetch(SERVER_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({}),
      muteHttpExceptions: true
    });

    var body = JSON.parse(result.getContentText());

    if (body.queued > 0) {
      ui.alert("実行開始", body.queued + " 件の分析を開始しました。\n数分後にAnalysisシートを確認してください。", ui.ButtonSet.OK);
    } else {
      ui.alert("情報", body.message || "処理対象の行がありません。", ui.ButtonSet.OK);
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
    .addItem("▶ 未処理行を分析", "runAnalysis")
    .addToUi();
}

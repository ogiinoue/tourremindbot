const SLACK_WEBHOOK_URL = "##"; // 取得したSlackのWebhook URL
const TARGET_FORM_ID = "##"; // フォームIDを直接指定
const SHEET_NAME = "##"; // 対象のスプレッドシート名
const NOTIFICATION_COLUMN = 10; // J列（通知フラグ）
const NAME_COLUMN = 4; // d列（名前）
const EXCLUDE_COLUMNS = [0, 8, 9,10]; // 除外する列（A列=0, I列=8, J列=9,K列＝10）
const EXCLUDE_HEADERS = ["列1","スコア"]; // 通知から削除するヘッダー名

function setupTrigger() {
  Logger.log("setupTrigger: 既存のトリガーを確認します。");

  // 現在のプロジェクトに設定されている全トリガーを取得
  var triggers = ScriptApp.getProjectTriggers();
  
  // 同じハンドラー（sendSlackNotification）のトリガーがあれば削除する
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendSlackNotification') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log("setupTrigger: 重複するトリガーを削除しました。");
    }
  }
  
  // 新たにフォーム送信トリガーを作成
  ScriptApp.newTrigger('sendSlackNotification')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  Logger.log("setupTrigger: 新しいフォーム送信トリガーが作成されました。");
}

function escapeMarkdown(text) {
  Logger.log("escapeMarkdown: 入力テキスト: " + text);
  let escapedText = String(text).replace(/_/g, "\\_");
  Logger.log("escapeMarkdown: エスケープ後のテキスト: " + escapedText);
  return escapedText;
}

function sendSlackNotification(e) {
  Logger.log("sendSlackNotification: 関数の実行を開始します。");

  if (!e || !e.namedValues) {
    Logger.log("sendSlackNotification: イベントデータが取得できませんでした。処理を終了します。");
    return;
  }
  Logger.log("sendSlackNotification: イベントデータを確認しました。");

  // 対象スプレッドシートとシートの取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("sendSlackNotification: アクティブなスプレッドシートを取得しました。");
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log("sendSlackNotification: 指定されたシート「" + SHEET_NAME + "」が見つかりません。");
    return;
  }
  Logger.log("sendSlackNotification: シート「" + SHEET_NAME + "」が見つかりました。");

  // フォームURLの取得と対象フォームチェック
  const formUrl = sheet.getFormUrl();
  Logger.log("sendSlackNotification: 取得したフォームURL: " + formUrl);
  if (!formUrl || !formUrl.includes(TARGET_FORM_ID)) {
    Logger.log("sendSlackNotification: 対象外のフォームからの送信のため、処理をスキップします。");
    return;
  }
  Logger.log("sendSlackNotification: フォームIDが一致しました。");

  // ヘッダー行の取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log("sendSlackNotification: ヘッダー行を取得しました: " + headers.join(", "));

  // 最終行のデータの取得
  const lastRow = sheet.getLastRow();
  Logger.log("sendSlackNotification: 最終行の行番号は " + lastRow + " です。");
  const data = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log("sendSlackNotification: 最終行のデータ: " + data.join(", "));

  // 通知済みフラグと名前の取得
  const notificationValue = data[NOTIFICATION_COLUMN - 1]; // J列の値
  const nameValue = data[NAME_COLUMN - 1]; // D列（名前）の値
  Logger.log("sendSlackNotification: 通知済みフラグの値（J列）: " + notificationValue);
  Logger.log("sendSlackNotification: 名前の値（D列）: " + nameValue);

  // 通知済みチェック
  if (notificationValue === "1") {
    Logger.log("sendSlackNotification: この行はすでに通知済みのため、処理をスキップします。");
    return;
  }

  // 名前が空の場合は通知しない
  if (!nameValue) {
    Logger.log("sendSlackNotification: D列（名前）が空白のため、送信をスキップ。行番号: " + lastRow);
    return;
  }

  Logger.log("sendSlackNotification: 通知対象の行を処理中: " + lastRow);

  // メッセージ構築
  let messageParts = [];
  for (let i = 0; i < data.length; i++) {
    if (!EXCLUDE_COLUMNS.includes(i) && !EXCLUDE_HEADERS.includes(headers[i])) { 
      let part = `*${headers[i]}*：${escapeMarkdown(data[i])}`;
      messageParts.push(part);
      Logger.log("sendSlackNotification: メッセージに追加する項目: " + part);
    }
  }
  const message = `*見学会参加フォームの新しい回答*\n` + messageParts.join("\n");
  Logger.log("sendSlackNotification: 送信メッセージ:\n" + message);

  // Slackに通知を送信
  const payload = { 'text': message };
  Logger.log("sendSlackNotification: Slackに送信するペイロード: " + JSON.stringify(payload));
  const response = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  });
  Logger.log("sendSlackNotification: Slackへの通知結果のHTTPステータスコード: " + response.getResponseCode());

  // 通知済みフラグの更新
  sheet.getRange(lastRow, NOTIFICATION_COLUMN).setValue("1");
  Logger.log("sendSlackNotification: 行 " + lastRow + " の通知済みフラグを「1」に設定しました。");
}

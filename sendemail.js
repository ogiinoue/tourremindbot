
function sendEmail() {
  try {
    var spreadsheet = SpreadsheetApp.openById("##");フォームIDを直接指定
    var sheet = spreadsheet.getSheetByName("##");// 対象のスプレッドシート名
    if (!sheet) {
      Logger.log("エラー: 指定されたシートが見つかりません。");
      return;
    }

    var lastRow = sheet.getLastRow();
    Logger.log("処理対象の最終行: " + lastRow);

    // 2行目から最終行までループ（ヘッダーが1行目の場合）
    for (var row = 2; row <= lastRow; row++) {
      // 送信済みフラグ（J列 = 11列目）の取得
      var isSendCell = sheet.getRange(row, 11);
      var isSendValue = isSendCell.getValue();

      if (isSendValue === "" || isSendValue == null) {
        // 名前（D列）とメールアドレス（C列）の取得
        // C列から2列分を取得すると、[メールアドレス, 名前]の順番になる
        var rowData = sheet.getRange(row, 3, 1, 2).getDisplayValues()[0];
        let mailAddress = rowData[0];
        let name = rowData[1];

        // メールアドレスの簡易チェック
        if (!mailAddress.includes("@") || !mailAddress.includes(".")) {
          Logger.log("エラー: 無効なメールアドレス - " + mailAddress + "（行 " + row + "）");
          continue;
        }

        // 見学日（G列 = 7列目）の取得
        let date = sheet.getRange(row, 7).getDisplayValue();

        // 返信メール件名と本文の作成
        let subject = 'ご見学についての確認メール【株式会社〇〇】';
        let body = makeBody(name, date);

        // メール送信オプション設定
        var options = {
          from: 'ndcogi@gmail.com',
          cc: 'nddcogi@gmail.com'
        };

        // Gmailに下書きを作成
        GmailApp.createDraft(mailAddress, subject, body, options);
        Logger.log("行 " + row + " の下書きを作成: " + mailAddress);

        // 送信済みフラグ（J列）を「1」にセット
        isSendCell.setValue(1);
      } else {
        Logger.log("行 " + row + " はすでに下書き作成済みです。");
      }
    }
  } catch (error) {
    Logger.log("エラー: " + error.message);
  }
}


function makeBody(name, dateString) {
  var parsedDate = parseCustomDate(dateString);
  if (!parsedDate) {
    Logger.log("日付のパースに失敗しました: " + dateString);
    parsedDate = new Date();
  }

  let formattedDate = parsedDate.toLocaleString('ja-JP', {
    year: 'numeric', month: 'numeric', day: 'numeric',
    hour: 'numeric', minute: 'numeric', hour12: false
  });

  let body = `${name} 様\n\n`
           + "お世話になっております。\n\n"
           + "株式会社〇〇の〇〇です。\n\n"
           + "先日は見学会のご予約をいただき、誠にありがとうございました。\n\n"
           + "見学会の日程が近づいておりますので、この度ご確認のためご連絡いたします。\n\n"
           + "【見学会日時】\n"
           + formattedDate + "\n\n"
           + "万が一ご不明な点がございましたらお気軽にお問い合わせください。\n\n"
           + "それでは、当日お会いできることを楽しみにしております。\n\n"
           + "【住所】〒〇〇-〇〇 〇〇〇〇〇〇\n\n"
           + "【GoogleMap】〇〇〇〇〇〇〇〇\n\n"
           + "【行き方】最寄りは〇〇駅です。〇〇施設です。\n\n"
           + "--------------------------------------------\n"
           + "株式会社〇〇\n"
           + "〇〇　〇〇\n"
           + "TEL　〇〇\n"
           + "HP　〇〇\n"
           + "--------------------------------------------\n";
  return body;
}

function parseCustomDate(str) {
  var pattern = /(\d{4})年(\d{1,2})月(\d{1,2})日.*?(\d{1,2}:\d{2})/;
  var match = str.match(pattern);

  if (!match) {
    Logger.log("parseCustomDate: パターンにマッチしません: " + str);
    return null;
  }

  var yyyy = match[1];
  var mm   = match[2];
  var dd   = match[3];
  var time = match[4];

  var parsedString = `${yyyy}/${mm}/${dd} ${time}`;
  return new Date(parsedString);
}
function main() {
  // スクリプトプロパティの取得
  const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const MAILING_LIST_ADDRESS = PropertiesService.getScriptProperties().getProperty('MAILING_LIST_ADDRESS');

  // プロパティのバリデーション
  if (!SPREADSHEET_ID) {
    Logger.log('エラー: SPREADSHEET_IDがスクリプトプロパティに設定されていません。');
    return;
  }
  if (!MAILING_LIST_ADDRESS || MAILING_LIST_ADDRESS.includes('<') || MAILING_LIST_ADDRESS.includes('(')) {
    Logger.log('エラー: MAILING_LIST_ADDRESSが正しく設定されていないか、不正な文字が含まれています。例: keio-equestrian-team.googlegroups.com');
    return;
  }
  if (!SHEET_ID) {
    Logger.log('エラー: SHEET_IDがスクリプトプロパティに設定されていません。');
    return;
  }

  // スプレッドシートとシートを取得
  const sheet = getSpreadsheetSheet(SPREADSHEET_ID, SHEET_ID);
  if (!sheet) {
    // getSpreadsheetSheet関数内でログが出力されるので、ここでは何もしないか、追加のログを出しても良い
    return;
  }

  // ヘッダー行が存在しない場合は追加
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['受信日時', '件名', '本文', '送信元']);
  }

  // Gmailラベル「ML」を取得または作成
  const mlLabel = getOrCreateGmailLabel('ML');
  if (!mlLabel) {
    Logger.log('エラー: MLラベルの取得または作成に失敗しました。');
    return;
  }

  // Gmailの検索クエリ
  const query = `in:inbox is:unread list:${MAILING_LIST_ADDRESS}`;

  // Gmailメッセージを処理してスプレッドシートに書き込み、ラベルに移動
  processGmailMessages(query, sheet, mlLabel);

  // Discordにメッセージを送信する
  checkAndNotifyMissingData(sheet); 
}

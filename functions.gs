/**
 * 指定されたスプレッドシートIDとシートIDに基づいてシートオブジェクトを取得します。
 * @param {string} spreadsheetId - スプレッドシートのID。
 * @param {string} sheetId - シートのID（文字列）。
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} 取得したシートオブジェクト、またはエラーの場合はnull。
 */
function getSpreadsheetSheet(spreadsheetId, sheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = null;
    const sheets = spreadsheet.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === Number(sheetId)) {
        sheet = sheets[i];
        break;
      }
    }
    if (!sheet) {
      Logger.log(`エラー: シートID ${sheetId} のシートが見つかりません。`);
      return null;
    }
    return sheet;
  } catch (e) {
    Logger.log(`スプレッドシートまたはシートの取得中にエラーが発生しました: ${e.message}`);
    return null;
  }
}

/**
 * 指定されたGmailクエリに基づいてメールを検索し、スプレッドシートに書き込み、指定されたラベルに移動します。
 * @param {string} query - GmailApp.searchで使用する検索クエリ。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 書き込み先のスプレッドシートのシートオブジェクト。
 * @param {GoogleAppsScript.Gmail.GmailLabel} label - 移動先のGmailラベルオブジェクト。
 */
function processGmailMessages(query, sheet, label) {
  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    Logger.log('新しい対象メールは見つかりませんでした。');
    return;
  }

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i]; // スレッドオブジェクトを取得
    const messages = thread.getMessages(); // スレッド内の個々のメッセージを取得

    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      const subject = message.getSubject();
      const body = message.getPlainBody();
      const date = message.getDate();
      const sender = message.getFrom();

      // スプレッドシートにデータを追加
      sheet.appendRow([date, subject, body, sender]);

      // ログにメールの情報を出力 (デバッグ用)
      Logger.log(`メールを処理しました: 件名「${subject}」, 送信元「${sender}」`);
    }
    
    // ★ 修正箇所: スレッド全体を「ML」ラベルに追加し、同時に受信トレイからアーカイブします。
    // thread.addLabel(label) でラベルを追加
    // thread.moveToArchive() で受信トレイからアーカイブ（既読になり、受信トレイから消えます）
    thread.addLabel(label);
    thread.moveToArchive();
    
    Logger.log(`スレッドを「${label.getName()}」に移動（アーカイブ）しました。`);
  }
}

/**
 * 指定された名前のGmailラベルを取得します。存在しない場合は新しく作成します。
 * @param {string} labelName - 取得または作成するラベルの名前。
 * @returns {GoogleAppsScript.Gmail.GmailLabel|null} 取得または作成されたラベルオブジェクト、またはエラーの場合はnull。
 */
function getOrCreateGmailLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    try {
      label = GmailApp.createLabel(labelName);
      Logger.log(`「${labelName}」ラベルが作成されました。`);
    } catch (e) {
      Logger.log(`エラー: 「${labelName}」ラベルの作成に失敗しました: ${e.message}`);
      return null;
    }
  }
  return label;
}

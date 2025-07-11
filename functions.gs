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

/**
 * スプレッドシートの指定範囲に未記入セルがあるかチェックし、Discordに通知します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - チェックするシートオブジェクト。
 */
function checkAndNotifyMissingData(sheet) {
  const lastRow = sheet.getLastRow();
  const startRow = 2; 

  // A列にデータがない場合は、チェックする行がないと判断します。
  if (lastRow < startRow) {
    Logger.log(`A列にデータがありません。チェック範囲: E${startRow}:A${lastRow} は無効です。`);
    // 未記入のデータがないと判断し、通知は行いません
    return;
  }
  
  const CHECK_RANGE = `E${startRow}:E${lastRow}`; 
  const range = sheet.getRange(CHECK_RANGE);
  const values = range.getValues();

  let missingDataFound = false;
  let missingCells = [];

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      // セルが空か、空白文字のみで構成されているかをチェック
      if (values[r][c] === "" || String(values[r][c]).trim() === "") {
        missingDataFound = true;
        // A1表記でセルの位置を取得
        // getRow()とgetColumn()は範囲の左上のセルの行番号と列番号を返します。
        // rとcはそこからの相対位置なので、足し合わせます。
        const row = range.getRow() + r;
        const column = range.getColumn() + c;
        const cellAddress = sheet.getRange(row, column).getA1Notation();
        missingCells.push(cellAddress);
      }
    }
  }

  const SPREADSHEET_URL = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_URL');
  if (missingDataFound) {
    const spreadsheet = sheet.getParent(); 
    const message = `スプレッドシート「${spreadsheet.getName()}」のシート「${sheet.getName()}」に未記入のセルがあります。\n\n未記入のセル: ${missingCells.join(", ")}\n\n${SPREADSHEET_URL}`; 
    sendDiscordMessage(message);
    Logger.log(message);
  } else {
    const spreadsheet = sheet.getParent();
    const message = `スプレッドシート「${spreadsheet.getName()}」のシート「${sheet.getName()}」に未記入のセルはありませんでした。`;
    Logger.log(message);
  }
}

/**
 * Discord Webhookにメッセージを送信します。
 * DISCORD_WEBHOOK_URLは呼び出し元（main.gsなど）で定義されていることを前提とします。
 * @param {string} message - Discordに送信するメッセージ。
 */
function sendDiscordMessage(message) {

  const DISCORD_WEBHOOK_URL  = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL');

  const payload = {
    content: message,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // エラー時に例外をスローせず、レスポンスでエラー内容を取得できるようにする
  };

  try {
    const response = UrlFetchApp.fetch(DISCORD_WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();
    if (responseCode >= 200 && responseCode < 300) {
      Logger.log("Discordにメッセージを送信しました。");
    } else {
      Logger.log(`Discordへのメッセージ送信に失敗しました。ステータスコード: ${responseCode}, レスポンス: ${response.getContentText()}`);
    }
  } catch (e) {
    Logger.log(`Discordへのメッセージ送信中にエラーが発生しました: ${e.message}`);
  }
}

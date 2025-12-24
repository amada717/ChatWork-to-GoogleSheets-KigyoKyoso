/**
 * チャットワークのWebhookからのPOSTリクエストを処理する
 */ 
/**
 * チャットワークのWebhookからのPOSTリクエストを処理する
 */ 
function doPost(e) {
  // ロックを取得（最大30秒待機）
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
  } catch (error) {
    Logger.log('Could not obtain lock after 30 seconds.');
    return ContentService.createTextOutput('locked');
  }
  
  try {
    // デバッグ用ログ
    Logger.log('=== doPost called ===');
    Logger.log('Timestamp: ' + new Date());
    
    // Webhookからのデータをパース
    const data = JSON.parse(e.postData.contents);
    const room_id = data.webhook_event.room_id;
    const message_id = data.webhook_event.message_id;
    const body = data.webhook_event.body;
    
    Logger.log('Room ID: ' + room_id);
    Logger.log('Message ID: ' + message_id);
    
    const SHEET_NAME = 'room_' + room_id;

    const scriptProperties = PropertiesService.getScriptProperties();
    const CHATWORK_TOKEN = scriptProperties.getProperty('CHATWORK_TOKEN');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);

    // ヘッダー作成（初回のみ）
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'message_id', 'subject', 'purpose', 'body', 'created_at', 'updated_at'
      ]);
    }

    // 重複チェック: 既に処理済みか確認（Lock内で再確認）
    const existingRow = findRowByColumn(sheet, 'message_id', message_id);
    
    if (existingRow) {
      Logger.log('Message already exists - skipping (row: ' + existingRow + ')');
      return ContentService.createTextOutput('already exists');
    }

    // テンプレートデータをパース
    const templateData = parseTemplateMessage(body);

    // テンプレート形式でない場合はスキップ
    if (!templateData) {
      Logger.log('Not a template message - skipping');
      return ContentService.createTextOutput('not a template message');
    }

    // 送信日時
    const send_time = data.webhook_event.send_time
      ? new Date(data.webhook_event.send_time * 1000)
      : new Date();

    // データオブジェクトを作成
    const dataObj = {
      'message_id': message_id,
      'subject': templateData.subject,
      'purpose': templateData.purpose,
      'body': templateData.body,
      'created_at': send_time,
      'updated_at': send_time
    };

    // 新規追加
    Logger.log('Appending new row');
    appendRowByHeaders(sheet, dataObj);

    Logger.log('=== doPost completed successfully ===');
    return ContentService.createTextOutput('ok');

  } catch (error) {
    Logger.log('=== Error in doPost ===');
    Logger.log('Error: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return ContentService.createTextOutput('error: ' + error.toString());
  } finally {
    // ロックを解放
    lock.releaseLock();
  }
}


/**
 * チャットワークメッセージからテンプレートデータを抽出
 */
function parseTemplateMessage(body) {
  // <件名>と<目的>の抽出
  const subjectMatch = body.match(/<件名>(.+?)(?=\n|<|$)/s);
  const purposeMatch = body.match(/<目的>(.+?)(?=\n|\[|$)/s);
  
  // どちらか一方でも欠けている場合はnullを返す
  if (!subjectMatch || !purposeMatch) return null;
  
  const subject = subjectMatch[1].trim();
  const purpose = purposeMatch[1].trim();
  
  // 本文の抽出（[hr]の後）
  const bodyMatch = body.match(/\[hr\]\s*([\s\S]+)$/);
  const bodyText = bodyMatch ? bodyMatch[1].trim() : '';
  
  return {
    subject: subject,
    purpose: purpose,
    body: bodyText
  };
}


/**
 * ヘッダー名からカラムインデックスを取得
 */
function getColumnIndex(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(headerName);
  return index !== -1 ? index + 1 : null; // 1始まりで返す
}


/**
 * ヘッダー名をキーとしてデータを保存
 */
function appendRowByHeaders(sheet, dataObj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = headers.map(header => dataObj[header] || '');
  sheet.appendRow(rowData);
}


/**
 * ヘッダー名をキーとして既存行を更新
 */
function updateRowByHeaders(sheet, rowIndex, dataObj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = headers.map(header => dataObj[header] || '');
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
}


/**
 * 特定カラムの値で行を検索
 */
function findRowByColumn(sheet, columnName, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnIndex = headers.indexOf(columnName);
  
  if (columnIndex === -1) return null;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][columnIndex] == value) {
      return i + 1; // 行番号を返す（1始まり）
    }
  }
  return null;
}


/**
 * 全角英数字を半角に変換する
 */ 
function toHalfWidth(str) {
  return str.replace(/[Ā-ー]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
}


/**
 * チャットワークのWebhookからのGETリクエストを処理する（テスト用）
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}


/**
 * 指定されたルームIDのメッセージをシートから取得して返す
 */
function getRoomMessages(roomId) {
  const sheetName = 'room_' + roomId;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}


/**
 * ルームIDとルーム名の一覧を取得する
 */
function getRoomList() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const roomList = [];
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith('room_')) {
      const roomId = name.replace('room_', '');
      const data = sheet.getDataRange().getValues();
      const roomName = data.length > 1 ? 'Room ' + roomId : 'unknown';
      roomList.push({ roomId, roomName });
    }
  });
  
  return roomList;
}
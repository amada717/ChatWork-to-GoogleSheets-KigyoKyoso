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
    const account_id = data.webhook_event.account_id;
    const body = data.webhook_event.body;
    
    Logger.log('Room ID: ' + room_id);
    Logger.log('Message ID: ' + message_id);
    Logger.log('Account ID: ' + account_id);
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const CHATWORK_TOKEN = scriptProperties.getProperty('CHATWORK_TOKEN');
    const SHEET_NAME = 'room_' + room_id;
    const LOG_SHEET_NAME = 'log_room_' + room_id;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // メインシート（編集可能）
    const sheet = spreadsheet.getSheetByName(SHEET_NAME)
      || spreadsheet.insertSheet(SHEET_NAME);

    // ログシート（編集履歴保存用）
    const logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME)
      || spreadsheet.insertSheet(LOG_SHEET_NAME);

    // メインシートのヘッダー作成（初回のみ）
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'message_id', 'from_account', 'to_accounts', 'subject', 'purpose', 'body', 'created_at', 'updated_at'
      ]);
    }

    // ログシートのヘッダー作成（初回のみ）
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow([
        'log_id', 'message_id', 'from_account', 'to_accounts', 'subject', 'purpose', 'body', 'action_type', 'action_at'
      ]);
    }

    // 投稿者名を取得
    const fromAccount = getAccountName(CHATWORK_TOKEN, room_id, account_id);
    
    // 宛先を抽出
    const toAccounts = extractToAccounts(body, CHATWORK_TOKEN, room_id);

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

    // ChatworkリンクURL生成
    const chatworkLink = `https://www.chatwork.com/#!rid${room_id}-${message_id}`;

    // 既存メッセージのチェック
    const existingRow = findRowByColumn(sheet, 'message_id', chatworkLink);

    // ログID生成（タイムスタンプ + message_id）
    const log_id = new Date().getTime() + '_' + message_id;

    if (existingRow) {
      // 既存の場合は更新
      Logger.log('Message exists - updating row: ' + existingRow);
      
      const dataObj = {
        'message_id': chatworkLink, // リンク形式で保存
        'from_account': fromAccount,
        'to_accounts': toAccounts,
        'subject': templateData.subject,
        'purpose': templateData.purpose,
        'body': templateData.body,
        'created_at': sheet.getRange(existingRow, getColumnIndex(sheet, 'created_at')).getValue(), // 元のcreated_atを保持
        'updated_at': new Date() // 更新日時を現在時刻に
      };
      
      updateRowByHeaders(sheet, existingRow, dataObj);

      // ログシートに更新履歴を追加
      const logDataObj = {
        'log_id': log_id,
        'message_id': chatworkLink, // リンク形式で保存
        'from_account': fromAccount,
        'to_accounts': toAccounts,
        'subject': templateData.subject,
        'purpose': templateData.purpose,
        'body': templateData.body,
        'action_type': 'updated',
        'action_at': new Date()
      };
      
      appendRowByHeaders(logSheet, logDataObj);
      
    } else {
      // 新規の場合は追加
      Logger.log('New message - appending row');
      
      const dataObj = {
        'message_id': chatworkLink, // リンク形式で保存
        'from_account': fromAccount,
        'to_accounts': toAccounts,
        'subject': templateData.subject,
        'purpose': templateData.purpose,
        'body': templateData.body,
        'created_at': send_time,
        'updated_at': send_time
      };
      
      appendRowByHeaders(sheet, dataObj);

      // ログシートに新規作成履歴を追加
      const logDataObj = {
        'log_id': log_id,
        'message_id': chatworkLink, // リンク形式で保存
        'from_account': fromAccount,
        'to_accounts': toAccounts,
        'subject': templateData.subject,
        'purpose': templateData.purpose,
        'body': templateData.body,
        'action_type': 'created',
        'action_at': send_time
      };
      
      appendRowByHeaders(logSheet, logDataObj);
    }

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
 * アカウントIDから名前を取得
 */
function getAccountName(token, roomId, accountId) {
  try {
    const url = `https://api.chatwork.com/v2/rooms/${roomId}/members`;
    const options = {
      method: 'get',
      headers: { 'X-ChatWorkToken': token }
    };
    const response = UrlFetchApp.fetch(url, options);
    const members = JSON.parse(response.getContentText());
    const member = members.find(m => m.account_id === accountId);
    return member ? member.name : 'ID:' + accountId;
  } catch (error) {
    Logger.log('Error getting account name: ' + error.toString());
    return 'ID:' + accountId;
  }
}


/**
 * メッセージ本文からTO情報を抽出
 */
function extractToAccounts(body, token, roomId) {
  // [To:1234567] 形式を抽出
  const toPattern = /\[To:(\d+)\]/g;
  const matches = [...body.matchAll(toPattern)];
  
  if (matches.length === 0) {
    return '指定なし';
  }
  
  // アカウントIDから名前を取得
  const accountIds = matches.map(match => parseInt(match[1]));
  const names = accountIds.map(id => getAccountName(token, roomId, id));
  
  return names.join(', ');
}


/**
 * チャットワークメッセージからテンプレートデータを抽出
 */
function parseTemplateMessage(body) {
  // [To:xxx]を除去してからパース
  const cleanBody = body.replace(/\[To:\d+\][^\n]*/g, '').trim();
  
  // <件名>と<目的>の抽出
  const subjectMatch = cleanBody.match(/<件名>(.+?)(?=\n|<|$)/s);
  const purposeMatch = cleanBody.match(/<目的>(.+?)(?=\n|\[|$)/s);
  
  // どちらか一方でも欠けている場合はnullを返す
  if (!subjectMatch || !purposeMatch) return null;
  
  // 全角英数字を半角に変換
  const subject = toHalfWidth(subjectMatch[1].trim());
  const purpose = toHalfWidth(purposeMatch[1].trim());
  
  // 本文の抽出（[hr]の後）
  const bodyMatch = cleanBody.match(/\[hr\]\s*([\s\S]+)$/);
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
  return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
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
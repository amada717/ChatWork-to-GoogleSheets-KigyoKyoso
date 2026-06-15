/**
 * チャットワークのWebhookからのPOSTリクエストを処理する
 */ 
const V2_MAIN_HEADERS = [
  'メッセージID', '投稿者', '宛先', '会社', '期限日', '指示', '内容', '本文', '作成日時', '更新日時', 'ステータス'
];

const V2_LOG_HEADERS = [
  'ログID', 'メッセージID', '投稿者', '宛先', '会社', '期限日', '指示', '内容', '本文', '操作種別', '操作日時'
];

const STATUS_HEADER = 'ステータス';
const STATUS_OPTIONS = ['未着手', '進行中', '完了'];
const DEFAULT_STATUS = '未着手';
const STATUS_COLORS = {
  '未着手': '#ffffff',
  '進行中': '#fce5cd',
  '完了': '#d9d9d9'
};

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
    const SHEET_NAME = 'room_v2_' + room_id;
    const LOG_SHEET_NAME = 'log_room_v2_' + room_id;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // メインシート（編集可能）
    const sheet = spreadsheet.getSheetByName(SHEET_NAME)
      || spreadsheet.insertSheet(SHEET_NAME);

    // ログシート（編集履歴保存用）
    const logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME)
      || spreadsheet.insertSheet(LOG_SHEET_NAME);

    ensureHeaders(sheet, V2_MAIN_HEADERS);
    ensureHeaders(logSheet, V2_LOG_HEADERS);
    applyStatusFeatures(sheet);

    // 投稿者名を取得
    const fromAccount = getAccountName(CHATWORK_TOKEN, room_id, account_id);
    
    // 宛先を抽出
    const toAccounts = extractToAccounts(body, CHATWORK_TOKEN, room_id);

    // メッセージデータをパース
    const parsedData = parseTemplateMessage(body);

    // 送信日時
    const send_time = data.webhook_event.send_time
      ? new Date(data.webhook_event.send_time * 1000)
      : new Date();

    // ChatworkリンクURL生成
    const chatworkLink = `https://www.chatwork.com/#!rid${room_id}-${message_id}`;

    const dataObj = {
      'メッセージID': chatworkLink,
      '投稿者': fromAccount,
      '宛先': toAccounts,
      '会社': parsedData.company,
      '期限日': parsedData.dueDate,
      '指示': parsedData.instruction,
      '内容': parsedData.content,
      '本文': parsedData.body,
      '作成日時': send_time,
      '更新日時': send_time,
      'ステータス': DEFAULT_STATUS
    };

    // メインシートでの既存メッセージのチェック
    const existingRow = findRowByColumn(sheet, 'メッセージID', chatworkLink);

    // ログID生成（タイムスタンプ + message_id）
    const log_id = new Date().getTime() + '_' + message_id;

    // 会社シートは #会社# が空でない場合のみ作成
    const companySheetName = getCompanySheetName(parsedData.company);
    const companySheet = companySheetName
      ? (spreadsheet.getSheetByName(companySheetName) || spreadsheet.insertSheet(companySheetName))
      : null;

    if (companySheet) {
      ensureHeaders(companySheet, V2_MAIN_HEADERS);
      applyStatusFeatures(companySheet);
    }

    if (existingRow) {
      // 既存の場合: 内容変更チェック
      Logger.log('Message exists - checking for changes');
      
      // 既存データを取得
      const existingFromAccount = getDisplayValueByHeader(sheet, existingRow, '投稿者');
      const existingToAccounts = getDisplayValueByHeader(sheet, existingRow, '宛先');
      const existingCompany = getDisplayValueByHeader(sheet, existingRow, '会社');
      const existingDueDate = getDisplayValueByHeader(sheet, existingRow, '期限日');
      const existingInstruction = getDisplayValueByHeader(sheet, existingRow, '指示');
      const existingContent = getDisplayValueByHeader(sheet, existingRow, '内容');
      const existingBody = getDisplayValueByHeader(sheet, existingRow, '本文');
      
      // 内容が変更されているかチェック
      const isChanged = (
        existingFromAccount !== fromAccount ||
        existingToAccounts !== toAccounts ||
        existingCompany !== parsedData.company ||
        existingDueDate !== parsedData.dueDate ||
        existingInstruction !== parsedData.instruction ||
        existingContent !== parsedData.content ||
        existingBody !== parsedData.body
      );

      dataObj['作成日時'] = getCellValueByHeader(sheet, existingRow, '作成日時') || send_time;
      dataObj['更新日時'] = isChanged
        ? new Date()
        : getCellValueByHeader(sheet, existingRow, '更新日時') || send_time;
      dataObj[STATUS_HEADER] = getDisplayValueByHeader(sheet, existingRow, STATUS_HEADER) || DEFAULT_STATUS;

      const companySynced = syncCompanySheet(companySheet, chatworkLink, dataObj);
      
      if (!isChanged) {
        // 内容が同じ場合はスキップ（Chatworkの重複Webhook対策）
        if (companySynced) {
          Logger.log('No main changes, but company sheet synced');
          return ContentService.createTextOutput('company synced');
        }
        Logger.log('No changes detected - skipping duplicate webhook');
        return ContentService.createTextOutput('no changes');
      }
      
      // 内容が変更されている場合は更新
      Logger.log('Changes detected - updating row: ' + existingRow);
      
      updateRowByHeaders(sheet, existingRow, dataObj);

      syncCompanySheet(companySheet, chatworkLink, dataObj);

      // ログシートに更新履歴を追加
      const logDataObj = {
        'ログID': log_id,
        'メッセージID': chatworkLink,
        '投稿者': fromAccount,
        '宛先': toAccounts,
        '会社': parsedData.company,
        '期限日': parsedData.dueDate,
        '指示': parsedData.instruction,
        '内容': parsedData.content,
        '本文': parsedData.body,
        '操作種別': 'updated',
        '操作日時': new Date()
      };
      
      appendRowByHeaders(logSheet, logDataObj);
    } else {
      // 新規の場合は追加
      Logger.log('New message - appending row');

      appendRowByHeaders(sheet, dataObj);
      syncCompanySheet(companySheet, chatworkLink, dataObj);

      // ログシートに新規作成履歴を追加
      const logDataObj = {
        'ログID': log_id,
        'メッセージID': chatworkLink,
        '投稿者': fromAccount,
        '宛先': toAccounts,
        '会社': parsedData.company,
        '期限日': parsedData.dueDate,
        '指示': parsedData.instruction,
        '内容': parsedData.content,
        '本文': parsedData.body,
        '操作種別': 'created',
        '操作日時': send_time
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
  const separatorIndex = cleanBody.indexOf('###');
  const metadataPart = separatorIndex === -1
    ? cleanBody
    : cleanBody.slice(0, separatorIndex);
  const bodyText = separatorIndex === -1
    ? ''
    : cleanBody.slice(separatorIndex + 3).replace(/^\r?\n/, '');
  
  return {
    company: extractTagValues(metadataPart, '会社'),
    dueDate: extractTagValues(metadataPart, '期限日'),
    instruction: extractTagValues(metadataPart, '指示'),
    content: extractTagValues(metadataPart, '内容'),
    body: bodyText
  };
}


/**
 * タグ値を抽出（同一タグが複数あれば改行連結）
 */
function extractTagValues(text, tagName) {
  const regex = new RegExp(`#${escapeRegExp(tagName)}#([^\n]*)`, 'g');
  const values = [];
  let match;

  while ((match = regex.exec(text)) !== null) {
    const value = match[1].trim();
    if (value) values.push(value);
  }

  return values.join('\n');
}


/**
 * 正規表現エスケープ
 */
function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


/**
 * ヘッダーを初回作成
 */
function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  // 既存シートには不足ヘッダーだけ末尾追加する
  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missingHeaders = headers.filter(header => existingHeaders.indexOf(header) === -1);
  if (missingHeaders.length > 0) {
    sheet.getRange(1, sheet.getLastColumn() + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }
}


/**
 * 会社シート名を生成（禁止文字置換、最大50文字）
 */
function getCompanySheetName(companyName) {
  const normalized = (companyName || '').trim();
  if (!normalized) return '';

  const safeName = normalized
    .replace(/[\\\/\?\*\[\]:]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();

  if (!safeName) return '';

  return (`company_${safeName}`).slice(0, 50);
}


/**
 * 会社シートへ同期（存在時のみ）
 */
function syncCompanySheet(companySheet, messageId, dataObj) {
  if (!companySheet) return false;

  const existingRow = findRowByColumn(companySheet, 'メッセージID', messageId);
  if (existingRow) {
    updateRowByHeaders(companySheet, existingRow, dataObj);
    return false;
  }

  appendRowByHeaders(companySheet, dataObj);
  return true;
}


/**
 * ヘッダー名でセル値を取得（ヘッダー不存在なら空文字）
 */
function getCellValueByHeader(sheet, rowIndex, headerName) {
  const columnIndex = getColumnIndex(sheet, headerName);
  if (!columnIndex) return '';
  return sheet.getRange(rowIndex, columnIndex).getValue();
}


/**
 * ヘッダー名で表示文字列を取得（比較用）
 */
function getDisplayValueByHeader(sheet, rowIndex, headerName) {
  const columnIndex = getColumnIndex(sheet, headerName);
  if (!columnIndex) return '';
  return sheet.getRange(rowIndex, columnIndex).getDisplayValue();
}


/**
 * ステータス列のプルダウンと行色ルールを適用
 */
function applyStatusFeatures(sheet) {
  const statusColIndex = getColumnIndex(sheet, STATUS_HEADER);
  if (!statusColIndex) return;

  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  const statusRange = sheet.getRange(2, statusColIndex, maxRows, 1);
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(validationRule);

  const statusColLetter = columnToLetter(statusColIndex);
  const dataRange = sheet.getRange(2, 1, maxRows, sheet.getLastColumn());

  const keepRules = sheet.getConditionalFormatRules().filter(rule => {
    const condition = rule.getBooleanCondition();
    if (!condition) return true;
    if (condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return true;
    const criteriaValues = condition.getCriteriaValues();
    const formula = criteriaValues && criteriaValues[0] ? String(criteriaValues[0]) : '';
    return !(
      formula === `=$${statusColLetter}2="完了"` ||
      formula === `=$${statusColLetter}2="未着手"` ||
      formula === `=$${statusColLetter}2="進行中"`
    );
  });

  keepRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${statusColLetter}2="完了"`)
      .setBackground(STATUS_COLORS['完了'])
      .setRanges([dataRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${statusColLetter}2="未着手"`)
      .setBackground(STATUS_COLORS['未着手'])
      .setRanges([dataRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${statusColLetter}2="進行中"`)
      .setBackground(STATUS_COLORS['進行中'])
      .setRanges([dataRange])
      .build()
  );

  sheet.setConditionalFormatRules(keepRules);
}


/**
 * 編集時に room_v2 と company のステータスを同期
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const editedRow = e.range.getRow();
    const editedCol = e.range.getColumn();

    if (editedRow <= 1) return;
    if (!(sheetName.startsWith('room_v2_') || sheetName.startsWith('company_'))) return;

    const statusColIndex = getColumnIndex(sheet, STATUS_HEADER);
    if (!statusColIndex || editedCol !== statusColIndex) return;

    const newStatus = e.value || '';
    const messageId = getDisplayValueByHeader(sheet, editedRow, 'メッセージID');
    if (!messageId) return;

    const spreadsheet = sheet.getParent();

    if (sheetName.startsWith('room_v2_')) {
      const companyName = getDisplayValueByHeader(sheet, editedRow, '会社');
      const companySheetName = getCompanySheetName(companyName);
      if (!companySheetName) return;

      const companySheet = spreadsheet.getSheetByName(companySheetName);
      if (!companySheet) return;

      syncStatusByMessageId(companySheet, messageId, newStatus);
      return;
    }

    // company_* 側で編集された場合は room_v2_* へ同期
    const roomId = extractRoomIdFromMessageLink(messageId);
    if (!roomId) return;

    const roomSheet = spreadsheet.getSheetByName('room_v2_' + roomId);
    if (!roomSheet) return;

    syncStatusByMessageId(roomSheet, messageId, newStatus);
  } catch (error) {
    Logger.log('Error in onEdit: ' + error.toString());
  }
}


/**
 * メッセージIDで行を特定し、ステータスを更新
 */
function syncStatusByMessageId(targetSheet, messageId, status) {
  const rowIndex = findRowByColumn(targetSheet, 'メッセージID', messageId);
  if (!rowIndex) return;

  const statusColIndex = getColumnIndex(targetSheet, STATUS_HEADER);
  if (!statusColIndex) return;

  const currentStatus = targetSheet.getRange(rowIndex, statusColIndex).getDisplayValue();
  if (currentStatus === status) return;

  targetSheet.getRange(rowIndex, statusColIndex).setValue(status);
}


/**
 * メッセージリンクから room_id を抽出
 */
function extractRoomIdFromMessageLink(messageLink) {
  const match = String(messageLink || '').match(/rid(\d+)-\d+/);
  return match ? match[1] : '';
}


/**
 * カラム番号をA1形式の列文字へ変換
 */
function columnToLetter(columnNumber) {
  let num = columnNumber;
  let letter = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    num = Math.floor((num - mod) / 26);
  }
  return letter;
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
 * ヘッダー名をキーとして既存行を更新（dataObj にないキーの列は触らない）
 */
function updateRowByHeaders(sheet, rowIndex, dataObj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach((header, i) => {
    if (header in dataObj) {
      sheet.getRange(rowIndex, i + 1).setValue(dataObj[header]);
    }
  });
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
  const sheetName = 'room_v2_' + roomId;
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
    if (name.startsWith('room_v2_')) {
      const roomId = name.replace('room_v2_', '');
      const data = sheet.getDataRange().getValues();
      const roomName = data.length > 1 ? 'Room ' + roomId : 'unknown';
      roomList.push({ roomId, roomName });
    }
  });
  
  return roomList;
}
/**
 * チャットワークのWebhookからのPOSTリクエストを処理する
 */ 
function doPost(e) {
  // デバッグ用ログ
  Logger.log('doPost called');
  Logger.log(JSON.stringify(e));

  try {
    // Webhookからのデータをパース
    const data = JSON.parse(e.postData.contents);
    const room_id = data.webhook_event.room_id;
    const SHEET_NAME = 'room_' + room_id;

    const scriptProperties = PropertiesService.getScriptProperties();
    const CHATWORK_TOKEN = scriptProperties.getProperty('CHATWORK_TOKEN');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);

    // ヘッダー作成（初回のみ）
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'message_id', 'message_type', 'company', 'deadline', 
        'subject', 'importance', 'purpose', 'body', 'tags', 
        'details_json', 'created_at', 'updated_at'
      ]);
    }

    // メッセージ情報を取得
    const message_id = data.webhook_event.message_id;
    const body = data.webhook_event.body;

    // テンプレートデータをパース
    const templateData = parseTemplateMessage(body);

    // テンプレート形式でない場合はスキップ
    if (!templateData) {
      return ContentService.createTextOutput('not a template message');
    }

    // ルーム名取得
    let room_name = '';
    try {
      const roomInfoUrl = `https://api.chatwork.com/v2/rooms/${room_id}`;
      const roomInfoOptions = {
        method: 'get',
        headers: { 'X-ChatWorkToken': CHATWORK_TOKEN }
      };
      const roomInfoResponse = UrlFetchApp.fetch(roomInfoUrl, roomInfoOptions);
      const roomInfo = JSON.parse(roomInfoResponse.getContentText());
      room_name = roomInfo.name;
    } catch (error) {
      Logger.log('ルーム名取得エラー: ' + error.toString());
      room_name = 'unknown';
    }

    // アカウント名取得
    let account = '';
    try {
      const account_id = data.webhook_event.account_id;
      const accountUrl = `https://api.chatwork.com/v2/rooms/${room_id}/members`;
      const accountOptions = {
        method: 'get',
        headers: { 'X-ChatWorkToken': CHATWORK_TOKEN }
      };
      const accountResponse = UrlFetchApp.fetch(accountUrl, accountOptions);
      const members = JSON.parse(accountResponse.getContentText());
      const member = members.find(m => m.account_id === account_id);
      account = member ? member.name : 'unknown';
    } catch (error) {
      Logger.log('アカウント名取得エラー: ' + error.toString());
      account = 'unknown';
    }

    // 送信日時
    const send_time = data.webhook_event.send_time
      ? new Date(data.webhook_event.send_time * 1000)
      : new Date();

    // 既存メッセージの更新チェック
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    let updated = false;

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] == message_id) {
        // 既存行を更新
        sheet.getRange(i + 1, 1, 1, 12).setValues([[
          message_id,
          templateData.message_type,
          templateData.company,
          templateData.deadline,
          templateData.subject,
          templateData.importance,
          templateData.purpose,
          templateData.body,
          templateData.tags,
          JSON.stringify(templateData.details),
          send_time,
          new Date()
        ]]);
        updated = true;
        break;
      }
    }

    // 新規追加
    if (!updated) {
      sheet.appendRow([
        message_id,
        templateData.message_type,
        templateData.company,
        templateData.deadline,
        templateData.subject,
        templateData.importance,
        templateData.purpose,
        templateData.body,
        templateData.tags,
        JSON.stringify(templateData.details),
        send_time,
        send_time
      ]);
    }

    return ContentService.createTextOutput('ok');

  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return ContentService.createTextOutput('error: ' + error.toString());
  }
}


/**
 * チャットワークメッセージからテンプレートデータを抽出
 */
function parseTemplateMessage(body) {
  // [info][title]依頼[/title] 形式からメッセージ種別を抽出
  const typeMatch = body.match(/\[title\](.*?)\[\/title\]/);
  if (!typeMatch) return null;
  
  const messageType = typeMatch[1];
  
  // 各項目を正規表現で抽出
  const extractField = (fieldName) => {
    const regex = new RegExp(`${fieldName}[::：]\\s*(.+?)(?=\\n|\\[|$)`, 's');
    const match = body.match(regex);
    return match ? match[1].trim() : '';
  };
  
  // 共通項目の抽出（デフォルト値として「指定なし」を設定）
  const templateData = {
    message_type: messageType,
    company: extractField('企業名') || '指定なし',
    deadline: extractField('期限') || '指定なし',
    subject: extractField('件名'),
    importance: extractField('重要度') || '指定なし',
    purpose: extractField('目的'),
    body: '',
    tags: '',
    details: {}
  };

  // 本文の抽出（3番目の[hr]と4番目の[hr]の間）
  const hrSplit = body.split('[hr]');
  if (hrSplit.length >= 4) {
    // 4番目の[hr]の後（本文部分）
    templateData.body = hrSplit[3].trim();
  }

  // タグの抽出（最後の[hr]の後、[/info]の前）
  const tagMatch = body.match(/\[hr\]\s*#(.+?)\s*\[\/info\]/s);
  if (tagMatch) {
    // #Web制作, #新規事業, #見積もり依頼 → Web制作,新規事業,見積もり依頼
    templateData.tags = tagMatch[1]
      .split(',')
      .map(tag => tag.replace(/#/g, '').trim())
      .filter(tag => tag !== '')
      .join(',');
  }
  
  // メッセージタイプ別の詳細項目を抽出
  switch (messageType) {
    case '依頼':
      templateData.details = {
        expected_response: extractField('期待する返答'),
        background: extractField('背景/前提'),
        constraints: extractField('制約事項/補足')
      };
      break;
    
    case '報告':
      templateData.details = {
        report_period: extractField('実施期間'),
        action_taken: extractField('実施事項'),
        result: extractField('結果/進捗'),
        next_action: extractField('次工程の予定'),
        issue: extractField('トラブル/課題')
      };
      break;
    
    case '相談':
      templateData.details = {
        current_issue: extractField('現状の課題'),
        options_considered: extractField('検討済みの選択肢'),
        request: extractField('求める意見/決定')
      };
      break;
    
    case '連絡':
      templateData.details = {
        contact_type: extractField('連絡種別'),
        before_change: extractField('変更前の情報'),
        action_required: extractField('対応要否'),
        after_action: extractField('対応要の場合')
      };
      break;
    
    case '確認':
      templateData.details = {
        check_target: extractField('確認対象'),
        criteria: extractField('判断基準'),
        check_reason: extractField('依頼理由')
      };
      break;
    
    case '承認':
      templateData.details = {
        approval_target: extractField('承認対象'),
        rationale: extractField('申請の根拠'),
        impact: extractField('金額/影響'),
        post_action: extractField('承認後の行動')
      };
      break;
    
    case '質問':
      templateData.details = {
        question_target: extractField('質問対象'),
        pre_check: extractField('試したこと/確認箇所'),
        expected_answer: extractField('期待する回答')
      };
      break;
    
    case '案内':
      templateData.details = {
        audience: extractField('周知対象'),
        urgency: extractField('対応要否/緊急度')
      };
      break;
    
    case '注意':
      templateData.details = {
        audience: extractField('周知対象'),
        urgency: extractField('対応要否/緊急度')
      };
      break;
  }
  
  return templateData;
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
      const roomName = data.length > 1 ? data[1][1] : 'unknown';
      roomList.push({ roomId, roomName });
    }
  });
  
  return roomList;
}
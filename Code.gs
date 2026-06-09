// ============================================================
//  LINE スケジュールBot  -  Code.gs
//  構成: LINE Messaging API + GAS + Google スプレッドシート
// ============================================================

// ===== 設定 =====
const CONFIG = {
  LINE_CHANNEL_ACCESS_TOKEN: 'YOUR_CHANNEL_ACCESS_TOKEN', // ← LINE Developersから取得
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',                  // ← スプレッドシートのID
  SHEET_SCHEDULE: 'schedules',   // 予定データシート
  SHEET_SESSION:  'sessions',    // 会話セッションシート
  REMINDER_HOUR:  7,             // リマインダー送信時刻（時）
};

// ============================================================
//  Web API  ― doGet / doPost 共通ルーター
//  カレンダーUIから fetch() で呼ぶ REST-like JSON API
//
//  【エンドポイント一覧】
//  GET  ?action=list&roomKey=XXX              → 全予定取得
//  GET  ?action=list&roomKey=XXX&year=2025&month=6 → 月絞り込み
//  POST action=add    body:{roomKey,title,date,start,end,venue,memo,allDay,color}
//  POST action=update body:{id,title,date,start,end,venue,memo,allDay,color}
//  POST action=delete body:{id}
// ============================================================

function doGet(e) {
  return handleApiRequest('GET', e);
}

// ===== エントリーポイント（LINEからのWebhook受信 + Web API POST）=====
function doPost(e) {
  try {
    const contents = e.postData && e.postData.contents;
    if (!contents) return jsonResponse({ status: 'ok' });

    const body = JSON.parse(contents);

    // LINE Webhookは必ず body.events 配列を持つ
    // Web APIリクエストは body.action を持つ
    if (body.action) {
      return handleApiRequest('POST', e);
    }

    // --- LINE Webhook ---
    const events = body.events || [];
    events.forEach(event => {
      if (event) handleEvent(event);
    });

  } catch (err) {
    Logger.log('doPost Error: ' + err);
  }
  return jsonResponse({ status: 'ok' });
}

// ===== API ルーター =====
function handleApiRequest(method, e) {
  try {
    if (method === 'GET') {
      const p      = e.parameter || {};
      const action = p.action || 'list';
      if (action === 'list') {
        const roomKey = p.roomKey || '';
        const year    = p.year  ? parseInt(p.year)  : null;
        const month   = p.month ? parseInt(p.month) : null;
        const data    = apiGetSchedules(roomKey, year, month);
        return jsonResponse({ status: 'ok', events: data });
      }
      return jsonResponse({ status: 'error', message: 'unknown action' });
    }

    if (method === 'POST') {
      const body   = JSON.parse(e.postData.contents);
      const action = body.action;

      if (action === 'add') {
        const ev = apiAddSchedule(body);
        return jsonResponse({ status: 'ok', event: ev });
      }
      if (action === 'update') {
        apiUpdateSchedule(body);
        return jsonResponse({ status: 'ok' });
      }
      if (action === 'delete') {
        deleteSchedule(body.id);
        return jsonResponse({ status: 'ok' });
      }
      return jsonResponse({ status: 'error', message: 'unknown action' });
    }
  } catch (err) {
    Logger.log('API Error: ' + err);
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== API 実装 =====
function apiGetSchedules(roomKey, year, month) {
  const sheet  = getSheet(CONFIG.SHEET_SCHEDULE);
  const data   = sheet.getDataRange().getValues();
  const header = data[0];
  const idx    = h => header.indexOf(h);

  return data.slice(1)
    .filter(row => {
      if (row[idx('done')] === true || row[idx('done')] === 'TRUE') return false;
      if (roomKey && row[idx('roomKey')] !== roomKey) return false;
      if (year && month) {
        const d = row[idx('date')];           // 'yyyy/MM/dd'
        const parts = d.split('/');
        if (parseInt(parts[0]) !== year)  return false;
        if (parseInt(parts[1]) !== month) return false;
      }
      return true;
    })
    .map(row => ({
      id:     String(row[idx('id')]),
      title:  row[idx('title')],
      date:   row[idx('date')].replace(/\//g, '-'), // yyyy-MM-dd に変換
      start:  row[idx('time')],
      end:    row[idx('endTime')] || '',
      venue:  row[idx('venue')]  || '',
      memo:   row[idx('memo')]   || '',
      allDay: row[idx('allDay')] === true || row[idx('allDay')] === 'TRUE',
      color:  parseInt(row[idx('color')]) || 0,
    }));
}

function apiAddSchedule(body) {
  const sheet = getSheet(CONFIG.SHEET_SCHEDULE);
  const id    = generateId();
  const date  = (body.date || '').replace(/-/g, '/'); // yyyy-MM-dd → yyyy/MM/dd
  sheet.appendRow([
    id,
    body.roomKey || 'web',
    'web',
    body.title   || '',
    date,
    body.start   || '',
    body.end     || '',
    body.venue   || '',
    body.memo    || '',
    false,
    new Date(),
    body.allDay  || false,
    body.color   || 0,
  ]);
  return { id: String(id), ...body, date: body.date };
}

function apiUpdateSchedule(body) {
  const sheet  = getSheet(CONFIG.SHEET_SCHEDULE);
  const data   = sheet.getDataRange().getValues();
  const header = data[0];
  const idx    = h => header.indexOf(h);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx('id')]) === String(body.id)) {
      const r = i + 1;
      const date = (body.date || '').replace(/-/g, '/');
      sheet.getRange(r, idx('title')  + 1).setValue(body.title  || '');
      sheet.getRange(r, idx('date')   + 1).setValue(date);
      sheet.getRange(r, idx('time')   + 1).setValue(body.start  || '');
      sheet.getRange(r, idx('endTime')+ 1).setValue(body.end    || '');
      sheet.getRange(r, idx('venue')  + 1).setValue(body.venue  || '');
      sheet.getRange(r, idx('memo')   + 1).setValue(body.memo   || '');
      sheet.getRange(r, idx('allDay') + 1).setValue(body.allDay || false);
      sheet.getRange(r, idx('color')  + 1).setValue(body.color  || 0);
      return;
    }
  }
}

// ===== イベント振り分け =====
function handleEvent(event) {
  if (!event || !event.type) return;
  if (event.type !== 'message' || !event.message || event.message.type !== 'text') return;

  const replyToken = event.replyToken;
  const userId     = event.source.userId;
  const groupId    = event.source.groupId || event.source.roomId || null;
  const roomKey    = groupId || userId;   // グループ or 個人を一意に識別
  const text       = event.message.text.trim();

  // セッション確認（登録・編集の途中かどうか）
  const session = getSession(roomKey);

  if (session) {
    handleSession(replyToken, roomKey, userId, text, session);
  } else {
    handleCommand(replyToken, roomKey, userId, text);
  }
}

// ===== コマンド処理 =====
function handleCommand(replyToken, roomKey, userId, text) {
  const cmd = text.toLowerCase();

  if (cmd === '予定登録' || cmd === '登録') {
    startRegistration(replyToken, roomKey);

  } else if (cmd === '予定一覧' || cmd === '一覧') {
    showScheduleList(replyToken, roomKey);

  } else if (cmd === '今日' || cmd === '今日の予定') {
    showTodaySchedule(replyToken, roomKey);

  } else if (cmd === '明日' || cmd === '明日の予定') {
    showScheduleByDate(replyToken, roomKey, getTomorrowStr());

  } else if (cmd.startsWith('削除')) {
    startDeletion(replyToken, roomKey, text);

  } else if (cmd.startsWith('編集')) {
    startEdit(replyToken, roomKey, text);

  } else if (cmd === 'ヘルプ' || cmd === 'help') {
    sendHelp(replyToken);

  } else {
    // 日付入力（例: 6/15, 2025/6/15）
    const date = parseDate(text);
    if (date) {
      showScheduleByDate(replyToken, roomKey, date);
    } else {
      sendHelp(replyToken);
    }
  }
}

// ============================================================
//  予定登録フロー
// ============================================================
function startRegistration(replyToken, roomKey) {
  setSession(roomKey, { step: 'title' });
  reply(replyToken, '📅 予定を登録します。\n\n予定のタイトルを入力してください。\n（例: 会議、運動会など）');
}

function handleSession(replyToken, roomKey, userId, text, session) {
  if (text === 'キャンセル') {
    clearSession(roomKey);
    reply(replyToken, '❌ 操作をキャンセルしました。');
    return;
  }

  switch (session.step) {
    case 'title':
      session.title = text;
      session.step  = 'date';
      setSession(roomKey, session);
      reply(replyToken, `📅 「${text}」\n\n日付を入力してください。\n（例: 6/20、2025/6/20）`);
      break;

    case 'date': {
      const date = parseDate(text);
      if (!date) {
        reply(replyToken, '⚠️ 日付の形式が正しくありません。\n例: 6/20 または 2025/6/20');
        return;
      }
      session.date = date;
      session.step = 'time';
      setSession(roomKey, session);
      reply(replyToken, `📅 日付: ${formatDateDisplay(date)}\n\n時刻を入力してください。\n（例: 14:00）\n終日の場合は「終日」と入力してください。`);
      break;
    }

    case 'time':
      session.time = (text === '終日') ? '終日' : text;
      session.step = 'venue';
      setSession(roomKey, session);
      reply(replyToken, `⏰ 時刻: ${session.time}

会場を入力してください。
（不要な場合は「なし」と入力）`);
      break;

    case 'venue':
      session.venue = (text === 'なし') ? '' : text;
      session.step  = 'memo';
      setSession(roomKey, session);
      reply(replyToken, `🏠 会場: ${session.venue || 'なし'}\n\nメモを入力してください。\n（不要な場合は「なし」と入力）`);
      break;

    case 'memo':
      session.memo = (text === 'なし') ? '' : text;
      clearSession(roomKey);
      saveSchedule(roomKey, userId, session);
      const msg = buildConfirmMsg(session);
      reply(replyToken, msg);
      break;

    // 編集フロー
    case 'edit_field': {
      const id = session.editId;
      if (text === '1') {
        session.step = 'edit_title'; setSession(roomKey, session);
        reply(replyToken, '新しいタイトルを入力してください。');
      } else if (text === '2') {
        session.step = 'edit_date';  setSession(roomKey, session);
        reply(replyToken, '新しい日付を入力してください。\n（例: 6/20）');
      } else if (text === '3') {
        session.step = 'edit_time';  setSession(roomKey, session);
        reply(replyToken, '新しい時刻を入力してください。\n（例: 14:00 または 終日）');
      } else if (text === '4') {
        session.step = 'edit_venue'; setSession(roomKey, session);
        reply(replyToken, '新しい会場を入力してください。');
      } else if (text === '5') {
        session.step = 'edit_memo';  setSession(roomKey, session);
        reply(replyToken, '新しいメモを入力してください。');
      } else {
        reply(replyToken, '1〜5の番号を入力してください。');
      }
      break;
    }
    case 'edit_title':
      updateScheduleField(session.editId, 'title', text);
      clearSession(roomKey);
      reply(replyToken, `✅ タイトルを「${text}」に変更しました。`);
      break;
    case 'edit_date': {
      const d = parseDate(text);
      if (!d) { reply(replyToken, '⚠️ 日付の形式が正しくありません。'); return; }
      updateScheduleField(session.editId, 'date', d);
      clearSession(roomKey);
      reply(replyToken, `✅ 日付を「${formatDateDisplay(d)}」に変更しました。`);
      break;
    }
    case 'edit_time':
      updateScheduleField(session.editId, 'time', text);
      clearSession(roomKey);
      reply(replyToken, `✅ 時刻を「${text}」に変更しました。`);
      break;
    case 'edit_venue':
      updateScheduleField(session.editId, 'venue', text);
      clearSession(roomKey);
      reply(replyToken, `✅ 会場を「${text}」に変更しました。`);
      break;
    case 'edit_memo':
      updateScheduleField(session.editId, 'memo', text);
      clearSession(roomKey);
      reply(replyToken, `✅ メモを更新しました。`);
      break;
  }
}

// ============================================================
//  予定一覧表示
// ============================================================
function showScheduleList(replyToken, roomKey) {
  const schedules = getUpcomingSchedules(roomKey, 10);
  if (schedules.length === 0) {
    reply(replyToken, '📭 登録されている予定はありません。\n「予定登録」で追加できます。');
    return;
  }
  let msg = '📋 直近の予定一覧\n' + '─'.repeat(16) + '\n';
  schedules.forEach((s, i) => {
    msg += `[${s.id}] ${formatDateDisplay(s.date)} ${s.time}\n`;
    msg += `　${s.title}`;
    if (s.venue) msg += `\n　📍 ${s.venue}`;
    if (s.memo) msg += `\n　📝 ${s.memo}`;
    msg += '\n\n';
  });
  msg += '─'.repeat(16) + '\n削除: 削除 [番号]\n編集: 編集 [番号]';
  reply(replyToken, msg);
}

function showTodaySchedule(replyToken, roomKey) {
  const today = getTodayStr();
  showScheduleByDate(replyToken, roomKey, today);
}

function showScheduleByDate(replyToken, roomKey, dateStr) {
  const schedules = getSchedulesByDate(roomKey, dateStr);
  const label = formatDateDisplay(dateStr);
  if (schedules.length === 0) {
    reply(replyToken, `📭 ${label}の予定はありません。`);
    return;
  }
  let msg = `📅 ${label}の予定\n` + '─'.repeat(16) + '\n';
  schedules.forEach(s => {
    msg += `⏰ ${s.time}　${s.title}`;
    if (s.venue) msg += `\n　📍 ${s.venue}`;
    if (s.memo) msg += `\n　📝 ${s.memo}`;
    msg += '\n\n';
  });
  reply(replyToken, msg.trimEnd());
}

// ============================================================
//  予定削除
// ============================================================
function startDeletion(replyToken, roomKey, text) {
  const id = extractId(text);
  if (!id) {
    reply(replyToken, '削除する予定の番号を指定してください。\n例: 削除 3\n\n番号は「予定一覧」で確認できます。');
    return;
  }
  const s = getScheduleById(id);
  if (!s || s.roomKey !== roomKey) {
    reply(replyToken, '⚠️ 該当する予定が見つかりませんでした。');
    return;
  }
  deleteSchedule(id);
  reply(replyToken, `🗑️ 削除しました。\n\n[${id}] ${formatDateDisplay(s.date)} ${s.time}\n${s.title}`);
}

// ============================================================
//  予定編集
// ============================================================
function startEdit(replyToken, roomKey, text) {
  const id = extractId(text);
  if (!id) {
    reply(replyToken, '編集する予定の番号を指定してください。\n例: 編集 3\n\n番号は「予定一覧」で確認できます。');
    return;
  }
  const s = getScheduleById(id);
  if (!s || s.roomKey !== roomKey) {
    reply(replyToken, '⚠️ 該当する予定が見つかりませんでした。');
    return;
  }
  setSession(roomKey, { step: 'edit_field', editId: id });
  reply(replyToken,
    `✏️ 編集する項目を番号で選んでください。\n\n` +
    `[${id}] ${formatDateDisplay(s.date)} ${s.time}\n${s.title}\n\n` +
    `1. タイトル\n2. 日付\n3. 時刻\n4. 会場\n5. メモ\n\n` +
    `（キャンセルは「キャンセル」と入力）`
  );
}

// ============================================================
//  ヘルプ
// ============================================================
function sendHelp(replyToken) {
  const msg =
    '📌 スケジュールBotの使い方\n' +
    '─'.repeat(16) + '\n' +
    '【予定を登録】\n「予定登録」\n\n' +
    '【予定を確認】\n「予定一覧」 → 直近10件\n「今日」 → 本日の予定\n「明日」 → 明日の予定\n「6/20」 → 日付を指定\n\n' +
    '【予定を削除】\n「削除 番号」\n例: 削除 3\n\n' +
    '【予定を編集】\n「編集 番号」\n例: 編集 3\n\n' +
    '─'.repeat(16) + '\n' +
    '⏰ リマインダーは毎朝7時に\n翌日の予定をお知らせします。';
  reply(replyToken, msg);
}

// ============================================================
//  リマインダー（毎朝トリガー実行）
// ============================================================
function sendDailyReminder() {
  const tomorrow = getTomorrowStr();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_SCHEDULE);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idxDate    = header.indexOf('date');
  const idxTitle   = header.indexOf('title');
  const idxTime    = header.indexOf('time');
  const idxVenue   = header.indexOf('venue');
  const idxMemo    = header.indexOf('memo');
  const idxRoomKey = header.indexOf('roomKey');
  const idxDone    = header.indexOf('done');

  // roomKey ごとにまとめる
  const roomMap = {};
  data.slice(1).forEach(row => {
    if (row[idxDone] === true || row[idxDone] === 'TRUE') return;
    if (row[idxDate] !== tomorrow) return;
    const key = row[idxRoomKey];
    if (!roomMap[key]) roomMap[key] = [];
    roomMap[key].push({
      title: row[idxTitle],
      time:  row[idxTime],
      venue: row[idxVenue] || '',
      memo:  row[idxMemo],
    });
  });

  Object.keys(roomMap).forEach(roomKey => {
    const list = roomMap[roomKey];
    let msg = `🔔 明日 ${formatDateDisplay(tomorrow)} の予定\n` + '─'.repeat(16) + '\n';
    list.forEach(s => {
      msg += `⏰ ${s.time}　${s.title}`;
      if (s.venue) msg += `\n　📍 ${s.venue}`;
      if (s.memo)  msg += `\n　📝 ${s.memo}`;
      msg += '\n';
    });
    // roomKey がグループIDかユーザーIDかで送信先を変える
    if (roomKey.startsWith('C') || roomKey.startsWith('R')) {
      pushToGroup(roomKey, msg);
    } else {
      pushToUser(roomKey, msg);
    }
  });
}

// ============================================================
//  LINE API 送信
// ============================================================
function reply(replyToken, message) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: message }],
  };
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

function pushToUser(userId, message) {
  _push(userId, message);
}

function pushToGroup(groupId, message) {
  _push(groupId, message);
}

function _push(to, message) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: to,
    messages: [{ type: 'text', text: message }],
  };
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

// ============================================================
//  スプレッドシート操作
// ============================================================
function getSheet(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === CONFIG.SHEET_SCHEDULE) {
      sheet.appendRow(['id', 'roomKey', 'userId', 'title', 'date', 'time', 'endTime', 'venue', 'memo', 'done', 'createdAt', 'allDay', 'color']);
    } else if (name === CONFIG.SHEET_SESSION) {
      sheet.appendRow(['roomKey', 'sessionJson', 'updatedAt']);
    }
  }
  return sheet;
}

function saveSchedule(roomKey, userId, s) {
  const sheet = getSheet(CONFIG.SHEET_SCHEDULE);
  const id    = generateId();
  sheet.appendRow([id, roomKey, userId, s.title, s.date, s.time, '', s.venue || '', s.memo || '', false, new Date(), false, 0]);
}

function getUpcomingSchedules(roomKey, limit) {
  const sheet = getSheet(CONFIG.SHEET_SCHEDULE);
  const data  = sheet.getDataRange().getValues();
  const header = data[0];
  const today  = getTodayStr();
  const results = [];

  data.slice(1).forEach((row, i) => {
    if (row[header.indexOf('done')] === true || row[header.indexOf('done')] === 'TRUE') return;
    if (row[header.indexOf('roomKey')] !== roomKey) return;
    if (row[header.indexOf('date')] < today) return;
    results.push({
      rowIndex: i + 2,
      id:    row[header.indexOf('id')],
      title: row[header.indexOf('title')],
      date:  row[header.indexOf('date')],
      time:  row[header.indexOf('time')],
      memo:  row[header.indexOf('memo')],
    });
  });

  results.sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
  return results.slice(0, limit);
}

function getSchedulesByDate(roomKey, dateStr) {
  const sheet = getSheet(CONFIG.SHEET_SCHEDULE);
  const data  = sheet.getDataRange().getValues();
  const header = data[0];
  const results = [];

  data.slice(1).forEach(row => {
    if (row[header.indexOf('done')] === true || row[header.indexOf('done')] === 'TRUE') return;
    if (row[header.indexOf('roomKey')] !== roomKey) return;
    if (row[header.indexOf('date')] !== dateStr) return;
    results.push({
      id:    row[header.indexOf('id')],
      title: row[header.indexOf('title')],
      date:  row[header.indexOf('date')],
      time:  row[header.indexOf('time')],
      memo:  row[header.indexOf('memo')],
    });
  });

  results.sort((a, b) => a.time.localeCompare(b.time));
  return results;
}

function getScheduleById(id) {
  const sheet = getSheet(CONFIG.SHEET_SCHEDULE);
  const data  = sheet.getDataRange().getValues();
  const header = data[0];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][header.indexOf('id')]) === String(id)) {
      return {
        rowIndex: i + 1,
        id:      data[i][header.indexOf('id')],
        roomKey: data[i][header.indexOf('roomKey')],
        title:   data[i][header.indexOf('title')],
        date:    data[i][header.indexOf('date')],
        time:    data[i][header.indexOf('time')],
        memo:    data[i][header.indexOf('memo')],
      };
    }
  }
  return null;
}

function deleteSchedule(id) {
  const sheet = getSheet(CONFIG.SHEET_SCHEDULE);
  const data  = sheet.getDataRange().getValues();
  const header = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][header.indexOf('id')]) === String(id)) {
      sheet.getRange(i + 1, header.indexOf('done') + 1).setValue(true);
      return;
    }
  }
}

function updateScheduleField(id, field, value) {
  const sheet = getSheet(CONFIG.SHEET_SCHEDULE);
  const data  = sheet.getDataRange().getValues();
  const header = data[0];
  const colIdx = header.indexOf(field) + 1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][header.indexOf('id')]) === String(id)) {
      sheet.getRange(i + 1, colIdx).setValue(value);
      return;
    }
  }
}

// ============================================================
//  セッション管理
// ============================================================
function getSession(roomKey) {
  const sheet = getSheet(CONFIG.SHEET_SESSION);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === roomKey) {
      try { return JSON.parse(data[i][1]); } catch(e) { return null; }
    }
  }
  return null;
}

function setSession(roomKey, sessionObj) {
  const sheet = getSheet(CONFIG.SHEET_SESSION);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === roomKey) {
      sheet.getRange(i + 1, 2).setValue(JSON.stringify(sessionObj));
      sheet.getRange(i + 1, 3).setValue(new Date());
      return;
    }
  }
  sheet.appendRow([roomKey, JSON.stringify(sessionObj), new Date()]);
}

function clearSession(roomKey) {
  const sheet = getSheet(CONFIG.SHEET_SESSION);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === roomKey) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

// ============================================================
//  ユーティリティ
// ============================================================
function getTodayStr() {
  const now = new Date();
  return Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function getTomorrowStr() {
  const d = new Date();
  d.setDate(d.getDate() + 1);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function parseDate(text) {
  // 対応形式: 6/20, 06/20, 2025/6/20, 2025-6-20, 2025-06-20
  const now  = new Date();
  const year = now.getFullYear();

  let m;
  if ((m = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/))) {
    return `${m[1]}/${String(m[2]).padStart(2,'0')}/${String(m[3]).padStart(2,'0')}`;
  }
  if ((m = text.match(/^(\d{1,2})[\/\-](\d{1,2})$/))) {
    return `${year}/${String(m[1]).padStart(2,'0')}/${String(m[2]).padStart(2,'0')}`;
  }
  return null;
}

function formatDateDisplay(dateStr) {
  // yyyy/MM/dd → M月d日(曜日)
  const parts = dateStr.split('/');
  const d     = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  const days  = ['日','月','火','水','木','金','土'];
  return `${parseInt(parts[1])}月${parseInt(parts[2])}日(${days[d.getDay()]})`;
}

function generateId() {
  const sheet = getSheet(CONFIG.SHEET_SCHEDULE);
  const last  = sheet.getLastRow();
  return last; // 行番号をIDとして使用（シンプル）
}

function extractId(text) {
  const m = text.match(/\d+/);
  return m ? m[0] : null;
}

function buildConfirmMsg(session) {
  let msg = '✅ 予定を登録しました！\n\n';
  msg += `📅 ${formatDateDisplay(session.date)}\n`;
  msg += `⏰ ${session.time}\n`;
  msg += `📌 ${session.title}`;
  if (session.venue) msg += `\n📍 ${session.venue}`;
  if (session.memo)  msg += `\n📝 ${session.memo}`;
  return msg;
}

// ============================================================
//  初期設定用：トリガー登録（一度だけ実行）
// ============================================================
function setupTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // 毎朝7時にリマインダー送信
  ScriptApp.newTrigger('sendDailyReminder')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.REMINDER_HOUR)
    .create();
  Logger.log('トリガーを設定しました（毎朝 ' + CONFIG.REMINDER_HOUR + '時）');
}

// ============================================================
//  初期設定用：シートを手動で作成
// ============================================================
function initSheets() {
  getSheet(CONFIG.SHEET_SCHEDULE);
  getSheet(CONFIG.SHEET_SESSION);
  Logger.log('シートを初期化しました。');
}

// ============================================================
//  Web API — カレンダーUI連携
//  GETリクエストで action パラメータにより処理を分岐
//  POSTリクエストはLINE Webhookと共用（action パラメータで判定）
// ============================================================

/**
 * GETエンドポイント
 * ?action=getEvents[&roomKey=xxx][&month=2025-06]
 * ?action=deleteEvent&id=xxx&roomKey=xxx
 */
function doGet(e) {
  const params  = e.parameter;
  const action  = params.action || '';
  const roomKey = params.roomKey || 'default';

  let result;
  try {
    switch (action) {
      case 'getEvents':
        result = apiGetEvents(roomKey, params.month || null);
        break;
      case 'deleteEvent':
        result = apiDeleteEvent(params.id, roomKey);
        break;
      default:
        result = { error: 'unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POSTエンドポイント（既存のLINE Webhookと共用）
 * body.action === 'createEvent' → 予定新規登録
 * body.action === 'updateEvent' → 予定更新
 * それ以外 → LINE Webhook処理
 */
const _origDoPost = doPost; // 既存doPostを退避
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'parse error' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Web API リクエスト判定
  if (body.action === 'createEvent' || body.action === 'updateEvent') {
    let result;
    try {
      result = body.action === 'createEvent'
        ? apiCreateEvent(body)
        : apiUpdateEvent(body);
    } catch (err) {
      result = { error: err.toString() };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // LINE Webhook 処理（既存ロジック）
  try {
    const events = body.events;
    if (Array.isArray(events)) events.forEach(event => handleEvent(event));
  } catch (err) {
    Logger.log('doPost LINE Error: ' + err);
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- API実装 ----

/**
 * 予定一覧取得
 * month: 'yyyy-MM' 形式で絞り込み（省略時は全件）
 */
function apiGetEvents(roomKey, month) {
  const sheet = getSheet(CONFIG.SHEET_SCHEDULE);
  const data  = sheet.getDataRange().getValues();
  const header = data[0];
  const idx = {
    id:      header.indexOf('id'),
    roomKey: header.indexOf('roomKey'),
    userId:  header.indexOf('userId'),
    title:   header.indexOf('title'),
    date:    header.indexOf('date'),
    time:    header.indexOf('time'),
    endTime: header.indexOf('endTime'),
    memo:    header.indexOf('memo'),
    color:   header.indexOf('color'),
    allDay:  header.indexOf('allDay'),
    done:    header.indexOf('done'),
  };

  const result = [];
  data.slice(1).forEach(row => {
    if (row[idx.done] === true || row[idx.done] === 'TRUE') return;
    if (row[idx.roomKey] !== roomKey) return;
    const date = row[idx.date];
    if (month && !String(date).startsWith(month.replace('-', '/'))) return;
    result.push({
      id:      String(row[idx.id]),
      title:   row[idx.title],
      date:    row[idx.date],
      start:   row[idx.time]    || '',
      end:     idx.endTime >= 0 ? (row[idx.endTime] || '') : '',
      memo:    row[idx.memo]    || '',
      color:   row[idx.color]   != null ? Number(row[idx.color]) : 0,
      allDay:  row[idx.allDay] === true || row[idx.allDay] === 'TRUE' || row[idx.time] === '終日',
    });
  });
  return { events: result };
}

/**
 * 予定新規作成
 */
function apiCreateEvent(body) {
  const sheet   = getSheet(CONFIG.SHEET_SCHEDULE);
  const roomKey = body.roomKey || 'default';
  const id      = generateId();
  // endTime 列が存在しない場合に追加
  ensureEndTimeColumn(sheet);
  sheet.appendRow([
    id,
    roomKey,
    body.userId || 'web',
    body.title  || '',
    body.date   || '',
    body.start  || (body.allDay ? '終日' : ''),
    body.end    || '',
    body.memo   || '',
    false,
    new Date(),
    body.color  != null ? body.color : 0,
    body.allDay ? true : false,
  ]);
  return { status: 'created', id: String(id) };
}

/**
 * 予定更新
 */
function apiUpdateEvent(body) {
  const sheet  = getSheet(CONFIG.SHEET_SCHEDULE);
  const data   = sheet.getDataRange().getValues();
  const header = data[0];
  ensureEndTimeColumn(sheet);

  const idIdx    = header.indexOf('id');
  const fields   = ['title', 'date', 'time', 'endTime', 'memo', 'color', 'allDay'];
  const bodyMap  = {
    title:   body.title,
    date:    body.date,
    time:    body.start  || (body.allDay ? '終日' : ''),
    endTime: body.end    || '',
    memo:    body.memo   || '',
    color:   body.color  != null ? body.color : 0,
    allDay:  body.allDay ? true : false,
  };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(body.id)) {
      fields.forEach(f => {
        const ci = header.indexOf(f);
        if (ci >= 0) sheet.getRange(i + 1, ci + 1).setValue(bodyMap[f]);
      });
      return { status: 'updated' };
    }
  }
  return { error: 'not found' };
}

/**
 * 予定削除（論理削除）
 */
function apiDeleteEvent(id, roomKey) {
  const sheet  = getSheet(CONFIG.SHEET_SCHEDULE);
  const data   = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx   = header.indexOf('id');
  const rkIdx   = header.indexOf('roomKey');
  const doneIdx = header.indexOf('done');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(id) && data[i][rkIdx] === roomKey) {
      sheet.getRange(i + 1, doneIdx + 1).setValue(true);
      return { status: 'deleted' };
    }
  }
  return { error: 'not found' };
}

/**
 * endTime / color / allDay 列がなければ追加
 */
function ensureEndTimeColumn(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const needed = ['endTime', 'color', 'allDay'];
  needed.forEach(col => {
    if (!header.includes(col)) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col);
    }
  });
}

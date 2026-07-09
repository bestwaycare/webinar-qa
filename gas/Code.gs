// ===================================================
//  ウェビナーQ&A システム - Google Apps Script
// ===================================================

const SPREADSHEET_ID   = '1LKEgMvXR25_HyD9tui5AQPSG0ZIyxmJ8sFMQ0Pzczb4';
const SHEET_NAME       = '質問一覧';
const INSTRUCTOR_EMAIL = 'bestway.mabuchi@gmail.com';

const props = () => PropertiesService.getScriptProperties();

function getTitle()        { return props().getProperty('WEBINAR_TITLE') || 'ウェビナータイトル'; }
function getVideoUrl()     { return props().getProperty('VIDEO_URL') || ''; }
function getSessionStart() { return props().getProperty('SESSION_START') || ''; }
function getBroadcastEnded() { return props().getProperty('BROADCAST_ENDED') === 'true'; }

const COL = {
  ID: 1, TIMESTAMP: 2, NAME: 3, EMAIL: 4,
  QUESTION: 5, STATUS: 6, REPLY: 7, REPLIED_AT: 8,
};

// ---------------------------------------------------
//  POST: メッセージ受信
// ---------------------------------------------------
function doPost(e) {
  try {
    const { name, email, question } = JSON.parse(e.postData.contents);
    if (!name || !question) return json({ status: 'error', message: '必須項目が不足しています' });

    const sheet = getSheet();
    const id    = 'Q' + new Date().getTime().toString(36).toUpperCase();
    const ts    = formatDate(new Date());
    const mail  = email || '';

    sheet.appendRow([id, ts, name, mail, question, '未回答', '', '']);

    // 参加通知以外のメッセージのみ講師に通知
    if (question !== '(参加しました)') {
      notifyInstructor({ id, name, question, ts });
    }

    return json({ status: 'ok', id });
  } catch (err) {
    return json({ status: 'error', message: err.message });
  }
}

// ---------------------------------------------------
//  GET: アクション振り分け
// ---------------------------------------------------
function doGet(e) {
  const action = e.parameter.action || 'list';
  try {
    if (action === 'list')        return json({ status: 'ok', questions: getQuestions() });
    if (action === 'listByEmail') return json({ status: 'ok', questions: getByEmail(e.parameter.email) });
    if (action === 'reply')       return json(sendReply(e.parameter.id, e.parameter.reply));
    if (action === 'getTitle')    return json({ status: 'ok', title: getTitle(), sessionStart: getSessionStart(), broadcastEnded: getBroadcastEnded() });
    if (action === 'setTitle')    return json(setTitleProp(e.parameter.title));
    if (action === 'setBroadcastStatus') return json(setBroadcastStatusProp(e.parameter.status));
    if (action === 'getVideoUrl') return json({ status: 'ok', url: getVideoUrl() });
    if (action === 'setVideoUrl') return json(setVideoUrlProp(e.parameter.url));
    return json({ status: 'error', message: '不明なアクション' });
  } catch (err) {
    return json({ status: 'error', message: err.message });
  }
}

// ---------------------------------------------------
//  内部関数
// ---------------------------------------------------

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID','受信日時','名前','メール','質問内容','ステータス','回答内容','回答日時']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRows() {
  const sheet = getSheet();
  const vals  = sheet.getDataRange().getValues();
  return vals.length <= 1 ? [] : vals.slice(1);
}

function rowToObj(row) {
  return {
    id: row[0], timestamp: row[1], name: row[2],
    email: row[3], question: row[4], status: row[5],
    reply: row[6], repliedAt: row[7],
  };
}

function getQuestions() { return getRows().map(rowToObj); }

function getByEmail(email) {
  if (!email) return [];
  return getRows().filter(r => r[COL.EMAIL - 1] === email).map(rowToObj);
}

function sendReply(id, reply) {
  if (!id || !reply) return { status: 'error', message: 'id と reply が必要です' };
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(id)) continue;
    const name = rows[i][COL.NAME - 1];
    const now  = formatDate(new Date());
    sheet.getRange(i + 1, COL.STATUS).setValue('回答済');
    sheet.getRange(i + 1, COL.REPLY).setValue(reply);
    sheet.getRange(i + 1, COL.REPLIED_AT).setValue(now);
    return { status: 'ok', message: `${name} さんへ返信しました` };
  }
  return { status: 'error', message: '該当IDが見つかりません' };
}

function notifyInstructor({ id, name, question, ts }) {
  GmailApp.sendEmail(
    INSTRUCTOR_EMAIL,
    `【${getTitle()}】新しいメッセージ: ${name}さん`,
    `新しいメッセージが届きました。\n\nID: ${id}\n日時: ${ts}\n名前: ${name}\n\n【メッセージ】\n${question}`
  );
}

function setTitleProp(title) {
  if (!title) return { status: 'error', message: 'title が必要です' };
  props().setProperty('WEBINAR_TITLE', title);
  props().setProperty('SESSION_START', new Date().toISOString());
  props().setProperty('BROADCAST_ENDED', 'false');
  return { status: 'ok', title };
}

function setBroadcastStatusProp(status) {
  props().setProperty('BROADCAST_ENDED', status === 'ended' ? 'true' : 'false');
  return { status: 'ok' };
}

function setVideoUrlProp(url) {
  if (!url) return { status: 'error', message: 'url が必要です' };
  props().setProperty('VIDEO_URL', url);
  return { status: 'ok', url };
}

function formatDate(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

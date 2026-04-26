/******************************************************
 * AT 學員管理 Web App - Google Apps Script
 * 穩定重修版：Google Sheet 同步 + 直接用網址建立 Google Calendar 提醒
 * Version: calendar-direct-v1
 ******************************************************/

const SHEET_NAME = 'students';
const HEADERS = [
  'id', 'name', 'uid', 'phone', 'loginType', 'lineId', 'need',
  'teacherName', 'teacherUrl', 'classTime', 'status', 'createdAt'
];

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = String(params.action || 'list').trim();

    if (action === 'list') {
      return jsonOutput_({ ok: true, students: listStudents_() });
    }

    // 最穩定的日曆建立方式：前端直接用 window.open 開這個網址。
    // 參數：name, classTime, phone, lineId, teacherName
    if (action === 'addCalendarEventsSimple' || action === 'addCalendarEvents') {
      return jsonOutput_(addCalendarEventsSimple_(params));
    }

    // 獨立測試：在 10 分鐘後建立一個事件
    if (action === 'calendarTest') {
      const cal = CalendarApp.getDefaultCalendar();
      const start = new Date(Date.now() + 10 * 60 * 1000);
      const end = new Date(start.getTime() + 10 * 60 * 1000);
      const ev = cal.createEvent('AT管理工具網址測試', start, end, {
        description: '如果你看到這個事件，代表 Apps Script 可以寫入 Google 日曆。'
      });
      return jsonOutput_({ ok: true, message: '測試事件已建立', eventId: ev.getId() });
    }

    return jsonOutput_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOutput_({
      ok: false,
      error: String(err && err.message ? err.message : err),
      stack: String(err && err.stack ? err.stack : '')
    });
  }
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const body = JSON.parse(raw);
    const action = String(body.action || '').trim();

    if (action === 'save') {
      saveStudents_(body.students || []);
      return jsonOutput_({ ok: true, saved: Array.isArray(body.students) ? body.students.length : 0 });
    }

    return jsonOutput_({ ok: false, error: 'Unknown POST action: ' + action });
  } catch (err) {
    return jsonOutput_({
      ok: false,
      error: String(err && err.message ? err.message : err),
      stack: String(err && err.stack ? err.stack : '')
    });
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsHeader = HEADERS.some((h, i) => String(firstRow[i] || '') !== h);
  if (needsHeader) sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sheet;
}

function listStudents_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return rows
    .filter(row => row.some(cell => String(cell || '').trim() !== ''))
    .map(row => {
      const obj = {};
      HEADERS.forEach((h, i) => obj[h] = row[i] === undefined || row[i] === null ? '' : String(row[i]));
      return obj;
    })
    .filter(s => s.id);
}

function saveStudents_(students) {
  const sheet = getSheet_();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (!Array.isArray(students) || students.length === 0) return;
  const rows = students.map(s => HEADERS.map(h => String(s[h] || '')));
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
}

function parseClassTime_(classTimeText) {
  const text = String(classTimeText || '').trim();
  if (!text) throw new Error('缺少 classTime');

  const m = text.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6] || 0));
  }

  const d = new Date(text);
  if (isNaN(d.getTime())) throw new Error('classTime 格式無法解析：' + text);
  return d;
}

function addCalendarEventsSimple_(params) {
  const name = String(params.name || '').trim() || '未命名學員';
  const phone = String(params.phone || '').trim();
  const lineId = String(params.lineId || '').trim();
  const teacherName = String(params.teacherName || '').trim();
  const classTimeText = String(params.classTime || '').trim();

  const classTime = parseClassTime_(classTimeText);
  const now = new Date();
  const cal = CalendarApp.getDefaultCalendar();
  const tz = Session.getScriptTimeZone();

  const baseDesc = [
    'AT 學員管理工具自動建立',
    '學員：' + name,
    phone ? '手機：' + phone : '',
    lineId ? 'LINE：' + lineId : '',
    teacherName ? '老師：' + teacherName : '',
    '體驗課時間：' + Utilities.formatDate(classTime, tz, 'yyyy/MM/dd HH:mm')
  ].filter(Boolean).join('\n');

  const planned = [
    {
      key: 'preDay',
      title: '[AT課前一天] 提醒 ' + name + ' 上課',
      start: new Date(classTime.getTime() - 24 * 60 * 60 * 1000),
      duration: 15,
      desc: baseDesc + '\n\n待辦：傳送課前提醒訊息。'
    },
    {
      key: 'pre30',
      title: '[AT課前30分] ' + name + ' 即將上課',
      start: new Date(classTime.getTime() - 30 * 60 * 1000),
      duration: 10,
      desc: baseDesc + '\n\n待辦：注意 LINE 訊息，確認學員是否順利進教室。'
    },
    {
      key: 'post30',
      title: '[AT課後30分] 關懷 ' + name,
      start: new Date(classTime.getTime() + 30 * 60 * 1000),
      duration: 15,
      desc: baseDesc + '\n\n待辦：關心體驗課狀況。'
    },
    {
      key: 'log24',
      title: '[AT 24H Log] 跟進 ' + name,
      start: new Date(classTime.getTime() + 23 * 60 * 60 * 1000),
      duration: 30,
      desc: baseDesc + '\n\n待辦：完成 24H Log 與後續跟進。'
    }
  ];

  let created = 0;
  let skippedPast = 0;
  let skippedDuplicate = 0;
  const createdTitles = [];
  const skippedTitles = [];

  planned.forEach(item => {
    const end = new Date(item.start.getTime() + item.duration * 60 * 1000);

    // 已經過去的提醒不建立，避免日曆沒有通知造成混亂。
    if (item.start.getTime() <= now.getTime()) {
      skippedPast++;
      skippedTitles.push(item.title + '（時間已過）');
      return;
    }

    // 避免重複建立：同一時間前後 2 分鐘、同一標題就略過。
    const searchStart = new Date(item.start.getTime() - 2 * 60 * 1000);
    const searchEnd = new Date(end.getTime() + 2 * 60 * 1000);
    const exists = cal.getEvents(searchStart, searchEnd, { search: item.title })
      .some(ev => ev.getTitle() === item.title && Math.abs(ev.getStartTime().getTime() - item.start.getTime()) < 2 * 60 * 1000);

    if (exists) {
      skippedDuplicate++;
      skippedTitles.push(item.title + '（已存在）');
      return;
    }

    const ev = cal.createEvent(item.title, item.start, end, { description: item.desc });
    try {
      ev.removeAllReminders();
      ev.addPopupReminder(0);
    } catch (err) {
      try { ev.addPopupReminder(1); } catch (e2) {}
    }
    created++;
    createdTitles.push(item.title);
  });

  return {
    ok: true,
    message: 'Google 日曆處理完成',
    name: name,
    classTime: classTimeText,
    created: created,
    skippedPast: skippedPast,
    skippedDuplicate: skippedDuplicate,
    createdTitles: createdTitles,
    skippedTitles: skippedTitles
  };
}

// 第一次啟用自動加入 Google 日曆時，手動執行一次，讓 Google 授權 CalendarApp。
function authorizeCalendar() {
  const cal = CalendarApp.getDefaultCalendar();
  const start = new Date();
  const end = new Date(start.getTime() + 5 * 60 * 1000);
  const event = cal.createEvent('AT管理工具授權測試', start, end);
  event.deleteEvent();
}

// 手動測試：建立一個 10 分鐘後的事件。
function createRealCalendarTest() {
  const cal = CalendarApp.getDefaultCalendar();
  const start = new Date(Date.now() + 10 * 60 * 1000);
  const end = new Date(start.getTime() + 10 * 60 * 1000);
  cal.createEvent('AT管理工具日曆測試', start, end, {
    description: '如果你看到這個事件，代表 CalendarApp 權限正常。'
  });
}

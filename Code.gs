/**
 * CAP Cadet Check-In backend.
 *
 * Log: Date | Time | CAPID
 * Date tabs: CAPID | Time
 */

const TZ = 'America/New_York';
const LOG_HEADERS = ['Date', 'Time', 'CAPID'];
const DATE_HEADERS = ['CAPID', 'Time'];

/* ---------- SETUP ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove the old station-code setting.
  props_().deleteProperty('STATION_KEY');

  cleanupLegacyNameFields_(ss);

  const log = logSheet_();

  // Remove the old Roster tab.
  const oldRoster = ss.getSheetByName('Roster');
  if (oldRoster && ss.getSheets().length > 1) {
    ss.deleteSheet(oldRoster);
  }

  moveSheet_(log, 1);

  // Create dated tabs from existing Log records.
  const rows = log.getDataRange().getValues().slice(1);
  const byDate = {};

  rows.forEach(function(row) {
    const date = String(row[0] || '');
    const time = String(row[1] || '');
    const capid = String(row[2] || '');

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !capid) {
      return;
    }

    if (!byDate[date]) {
      byDate[date] = [];
    }

    byDate[date].push([capid, time]);
  });

  Object.keys(byDate).sort().forEach(function(date) {
    syncDateSheet_(date, byDate[date]);
  });

  ensureTrigger_();
}


/* ---------- LEGACY CLEANUP ---------- */

function cleanupLegacyNameFields_(ss) {
  const log = ss.getSheetByName('Log');

  if (
    log &&
    log.getMaxColumns() >= 4 &&
    String(log.getRange(1, 4).getValue()).toLowerCase() === 'name'
  ) {
    log.deleteColumn(4);
  }

  ss.getSheets().forEach(function(sh) {
    const name = sh.getName();

    if (!/^\d{1,2}-\d{1,2}-\d{2}$/.test(name)) {
      return;
    }

    if (
      sh.getMaxColumns() >= 3 &&
      String(sh.getRange(2, 3).getValue()).toLowerCase() === 'name'
    ) {
      sh.getRange('A1:C1').breakApart();
      sh.deleteColumn(3);
      sh.getRange('A1:B1').merge();
      sh.getRange('A1:B1')
        .setFontWeight('bold')
        .setFontSize(16)
        .setHorizontalAlignment('center');

      sh.setFrozenRows(2);
    }
  });
}


/* ---------- WEB APP ---------- */

function doGet() {
  return json_({
    ok: true,
    service: 'cap-checkin'
  });
}

function doPost(e) {
  let req;

  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({
      ok: false,
      error: 'Bad request'
    });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    switch (req.action) {
      case 'checkin':
        return json_(checkin_(req));

      case 'getConfig':
        return json_(getConfig_(req));

      case 'saveConfig':
        return json_(saveConfig_(req));

      case 'sendTest':
        return json_(sendTest_(req));

      default:
        return json_({
          ok: false,
          error: 'Unknown action'
        });
    }
  } catch (err) {
    return json_({
      ok: false,
      error: String(err)
    });
  } finally {
    lock.releaseLock();
  }
}


/* ---------- CHECK-IN ---------- */

function checkin_(req) {
  const capid = String(req.capid || '').replace(/\D/g, '');

  if (capid.length < 4 || capid.length > 8) {
    return {
      ok: false,
      error: 'Unreadable card, scan again'
    };
  }

  let when = req.ts ? new Date(req.ts) : new Date();

  if (isNaN(when.getTime())) {
    when = new Date();
  }

  const today = fmt_(when, 'yyyy-MM-dd');
  const time = fmt_(when, 'HH:mm:ss');

  const log = logSheet_();
  const rows = log.getDataRange().getValues();

  // Prevent a CAPID from checking in twice on the same date.
  for (let i = 1; i < rows.length; i++) {
    if (
      String(rows[i][0]) === today &&
      String(rows[i][2]) === capid
    ) {
      return {
        ok: true,
        duplicate: true,
        capid: capid,
        time: String(rows[i][1])
      };
    }
  }

  const row = log.getLastRow() + 1;

  log.getRange(row, 1, 1, 3)
    .setNumberFormat('@')
    .setValues([[today, time, capid]]);

  appendDateRow_(today, capid, time);

  return {
    ok: true,
    duplicate: false,
    capid: capid,
    time: time
  };
}


/* ---------- CONFIGURATION ---------- */

function getConfig_(req) {
  const p = props_();

  if (!p.getProperty('ADMIN_PIN')) {
    return {
      ok: true,
      firstRun: true,
      email: '',
      time: '20:00'
    };
  }

  if (!auth_(req.pin)) {
    return {
      ok: false,
      error: 'Wrong admin PIN'
    };
  }

  return {
    ok: true,
    email: p.getProperty('EMAIL') || '',
    time: p.getProperty('SEND_TIME') || '20:00',
    lastSent: p.getProperty('LAST_SENT') || ''
  };
}

function saveConfig_(req) {
  const p = props_();
  const hasPin = !!p.getProperty('ADMIN_PIN');

  if (hasPin) {
    if (!auth_(req.pin)) {
      return {
        ok: false,
        error: 'Wrong admin PIN'
      };
    }
  } else if (!req.newPin || String(req.newPin).length < 4) {
    return {
      ok: false,
      error: 'Choose an admin PIN of at least 4 characters'
    };
  }

  const emails = String(req.email || '')
    .split(/[,;\s]+/)
    .filter(Boolean);

  if (!emails.length) {
    return {
      ok: false,
      error: 'Enter at least one email address'
    };
  }

  for (const email of emails) {
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return {
        ok: false,
        error: 'Invalid email: ' + email
      };
    }
  }

  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(String(req.time || ''))) {
    return {
      ok: false,
      error: 'Invalid time'
    };
  }

  if (req.newPin) {
    p.setProperty('ADMIN_PIN', String(req.newPin));
  }

  p.setProperty('EMAIL', emails.join(','));
  p.setProperty('SEND_TIME', String(req.time));

  ensureTrigger_();

  return {
    ok: true
  };
}

function sendTest_(req) {
  if (!auth_(req.pin)) {
    return {
      ok: false,
      error: 'Wrong admin PIN'
    };
  }

  const email = props_().getProperty('EMAIL');

  if (!email) {
    return {
      ok: false,
      error: 'Save an email address first'
    };
  }

  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const rows = todayRows_(today);

  sendReport_(email, today, rows, true);

  return {
    ok: true,
    count: rows.length
  };
}


/* ---------- SCHEDULED EMAIL ---------- */

function tick() {
  const p = props_();
  const email = p.getProperty('EMAIL');

  if (!email) return;

  const now = new Date();
  const today = fmt_(now, 'yyyy-MM-dd');

  if (p.getProperty('LAST_SENT') === today) {
    return;
  }

  const parts = (p.getProperty('SEND_TIME') || '20:00')
    .split(':')
    .map(Number);

  const nowMin =
    Number(fmt_(now, 'H')) * 60 +
    Number(fmt_(now, 'm'));

  const sendMin = parts[0] * 60 + parts[1];

  if (nowMin < sendMin) {
    return;
  }

  const rows = todayRows_(today);

  if (!rows.length) {
    return;
  }

  sendReport_(email, today, rows, false);
  p.setProperty('LAST_SENT', today);
}

function sendReport_(to, date, rows, isTest) {
  const esc = function(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  };

  const trs = rows.map(function(row) {
    return '<tr><td>' + esc(row[1]) +
      '</td><td>' + esc(row[2]) + '</td></tr>';
  }).join('');

  const html =
    '<p>' +
    (isTest ? '<b>[TEST]</b> ' : '') +
    '<b>' + rows.length + '</b> checked in on ' +
    esc(date) + '.</p>' +
    '<table border="1" cellpadding="6" cellspacing="0" ' +
    'style="border-collapse:collapse">' +
    '<tr style="background:#eee">' +
    '<th>Time</th><th>CAPID</th></tr>' +
    trs +
    '</table>' +
    '<p style="color:#666;font-size:12px">CSV copy attached.</p>';

  const csv =
    'Date,Time,CAPID\n' +
    rows.map(function(row) {
      return [row[0], row[1], row[2]].join(',');
    }).join('\n');

  MailApp.sendEmail({
    to: to,
    subject:
      (isTest ? '[TEST] ' : '') +
      'CAP Attendance ' + date +
      ' (' + rows.length + ' checked in)',
    htmlBody: html,
    body: rows.length + ' checked in on ' +
      date + '. See attached CSV.',
    attachments: [
      Utilities.newBlob(
        csv,
        'text/csv',
        'attendance-' + date + '.csv'
      )
    ],
    name: 'CAP Check-In'
  });
}


/* ---------- COMMON HELPERS ---------- */

function props_() {
  return PropertiesService.getScriptProperties();
}

function auth_(pin) {
  const stored = props_().getProperty('ADMIN_PIN');
  return !!stored && String(pin) === stored;
}

function fmt_(date, pattern) {
  return Utilities.formatDate(date, TZ, pattern);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * FIX FOR moveSheet_ IS NOT DEFINED
 *
 * Apps Script Sheet objects do not have setIndex().
 * Set the active sheet, then move it using the Spreadsheet object.
 */
function moveSheet_(sheet, position) {
  if (!sheet) return;

  const ss = sheet.getParent();

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(position);
}


/* ---------- MASTER LOG ---------- */

function logSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Log');

  if (!sh) {
    sh = ss.insertSheet('Log', 0);
    sh.getRange('A:C').setNumberFormat('@');
    sh.appendRow(LOG_HEADERS);
  }

  sh.setFrozenRows(1);
  sh.getRange('A:C').setNumberFormat('@');

  if (sh.getIndex() !== 1) {
    moveSheet_(sh, 1);
  }

  return sh;
}

function todayRows_(today) {
  return logSheet_()
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(function(row) {
      return String(row[0]) === today;
    })
    .sort(function(a, b) {
      return String(a[1]).localeCompare(String(b[1]));
    });
}


/* ---------- DATED SHEETS ---------- */

function displayDate_(date) {
  const d = new Date(date + 'T12:00:00Z');
  return Utilities.formatDate(d, TZ, 'M/d/yy');
}

function dateSheetName_(date) {
  const d = new Date(date + 'T12:00:00Z');
  return Utilities.formatDate(d, TZ, 'M-d-yy');
}

function dateSheet_(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = dateSheetName_(date);

  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);

    sh.getRange('A:C').setNumberFormat('@');

    sh.getRange('A1:B1')
      .merge()
      .setValue(displayDate_(date));

    sh.getRange('A2:B2').setValues([DATE_HEADERS]);

    sh.getRange('A1:B1')
      .setFontWeight('bold')
      .setFontSize(16)
      .setHorizontalAlignment('center');

    sh.getRange('A2:B2').setFontWeight('bold');
    sh.setFrozenRows(2);

    moveSheet_(sh, Math.max(2, ss.getSheets().length));

    sh.autoResizeColumns(1, 2);
  }

  return sh;
}

function appendDateRow_(date, capid, time) {
  const sh = dateSheet_(date);
  const existing = sh.getDataRange().getValues();

  for (let i = 2; i < existing.length; i++) {
    if (String(existing[i][0]) === capid) {
      return;
    }
  }

  const row = sh.getLastRow() + 1;

  sh.getRange(row, 1, 1, 2)
    .setNumberFormat('@')
    .setValues([[capid, time]]);

  sh.autoResizeColumns(1, 2);
}

function syncDateSheet_(date, rows) {
  const sh = dateSheet_(date);
  const existing = sh.getDataRange().getValues();
  const have = {};

  for (let i = 2; i < existing.length; i++) {
    if (existing[i][0]) {
      have[String(existing[i][0])] = true;
    }
  }

  const toAdd = rows.filter(function(row) {
    return row[0] && !have[String(row[0])];
  });

  if (toAdd.length) {
    sh.getRange(
      sh.getLastRow() + 1,
      1,
      toAdd.length,
      2
    )
      .setNumberFormat('@')
      .setValues(toAdd);
  }

  sh.autoResizeColumns(1, 2);
}


/* ---------- EMAIL TRIGGER ---------- */

function ensureTrigger_() {
  const exists = ScriptApp.getProjectTriggers()
    .some(function(trigger) {
      return trigger.getHandlerFunction() === 'tick';
    });

  if (!exists) {
    ScriptApp.newTrigger('tick')
      .timeBased()
      .everyMinutes(10)
      .create();
  }
}
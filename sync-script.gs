/** Presente sync backend v3 — paste into Apps Script of a private Google Sheet.  */
/** FIRST TIME:  1) Change TOKEN below.  2) Deploy > New deployment > Web app,    */
/**    Execute as: Me / Access: Anyone.  3) Put URL + token in the app.           */
/** UPDATING an existing deployment: replace all code with this, save, then       */
/**    Deploy > Manage deployments > pencil > Version: New version > Deploy.      */
/**    The URL stays the same.                                                    */
/** EMAILS: set SEND_EMAILS = true to email students when they are scanned.       */
/**    Quota: ~100 recipients/day on gmail.com, 1500/day on Workspace accounts.   */
/** ONLINE CHECK-IN: the app creates expiring sessions here; students hit the     */
/**    'chk' endpoint with only a session key — never your TOKEN.                 */
var TOKEN = 'change-me-please';
var SEND_EMAILS = false;
var EMAIL_FROM_NAME = 'Presente Attendance';
var TZ = 'Asia/Manila';   // change only if you teach in another timezone

var EV_HEAD = ['uid','ts','date','classId','studentId','name','status','via','time','device','srvTs'];
var SESS_HEAD = ['key','classId','className','date','createdTs','expiresTs','marked'];

function sheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ev = ss.getSheetByName('Events') || ss.insertSheet('Events');
  if (ev.getLastRow() === 0) ev.appendRow(EV_HEAD);
  var ro = ss.getSheetByName('Roster') || ss.insertSheet('Roster');
  var me = ss.getSheetByName('Meta') || ss.insertSheet('Meta');
  if (me.getLastRow() === 0) me.appendRow(['rosterTs', 0]);
  var se = ss.getSheetByName('Sessions') || ss.insertSheet('Sessions');
  if (se.getLastRow() === 0) se.appendRow(SESS_HEAD);
  return { ev: ev, ro: ro, me: me, se: se };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
function word_(st) {
  return st === 'P' ? 'Present' : st === 'L' ? 'Late' :
         st === 'A' ? 'Absent' : st === 'E' ? 'Excused' : st;
}
function nowDate_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowTime_() { return Utilities.formatDate(new Date(), TZ, 'h:mm:ss a'); }
function key_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

function rosterMap_(s) {
  var map = {};
  var rl = s.ro.getLastRow();
  if (rl > 0) {
    var rows = s.ro.getRange(1, 1, rl, 6).getValues();
    for (var i = 0; i < rows.length; i++) {
      var sid = String(rows[i][2] || '');
      if (sid) map[sid] = { clsId: String(rows[i][0] || ''), cls: String(rows[i][1] || ''),
        name: String(rows[i][3] || ''), em: String(rows[i][5] || '') };
    }
  }
  return map;
}
function className_(s, cid) {
  var rl = s.ro.getLastRow();
  if (rl > 0) {
    var rows = s.ro.getRange(1, 1, rl, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === cid) return String(rows[i][1] || '');
    }
  }
  return '';
}
function findSession_(s, key) {
  var last = s.se.getLastRow();
  if (last < 2) return null;
  var rows = s.se.getRange(2, 1, last - 1, SESS_HEAD.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === key) return { row: i + 2, v: rows[i] };
  }
  return null;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var s;
  if (p.chk) {                       /* public: session info for the check-in page */
    s = sheets_();
    var f = findSession_(s, String(p.chk));
    if (!f) return json_({ ok: false, err: 'unknown' });
    var exp = Number(f.v[5]);
    return json_({ ok: true, cls: String(f.v[2] || ''), exp: exp,
      now: Date.now(), open: Date.now() < exp });
  }
  if (p.t !== TOKEN) return json_({ ok: false, err: 'bad token' });
  s = sheets_();
  var since = Number(p.since || 0);
  var clientRts = Number(p.rts || 0);
  var rosterTs = Number(s.me.getRange(1, 2).getValue() || 0);
  var out = { ok: true, now: Date.now(), rosterTs: rosterTs, ev: [] };
  var last = s.ev.getLastRow();
  if (last > 1) {
    var rows = s.ev.getRange(2, 1, last - 1, EV_HEAD.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (Number(r[10]) > since) {
        out.ev.push({ u: String(r[0]), ts: Number(r[1]), d: String(r[2]),
          c: String(r[3]), s: String(r[4]), n: String(r[5]), st: String(r[6]),
          via: String(r[7]), tm: String(r[8]), dev: String(r[9]) });
      }
    }
  }
  if (rosterTs > clientRts) {
    var rl = s.ro.getLastRow();
    out.roster = rl > 0 ? s.ro.getRange(1, 1, rl, 6).getValues() : [];
  }
  return json_(out);
}

function checkin_(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var s = sheets_();
    var f = findSession_(s, String(body.k || ''));
    if (!f) return json_({ ok: false, err: 'unknown' });
    if (Date.now() > Number(f.v[5])) return json_({ ok: false, err: 'expired' });
    var sid = String(body.sid || '').trim();
    var rmap = rosterMap_(s);
    var info = rmap[sid];
    if (!info) return json_({ ok: false, err: 'unknown' });
    var marked = String(f.v[6] || '');
    var hits = marked ? marked.split(',') : [];
    for (var i = 0; i < hits.length; i++) {
      var pair = hits[i].split('@');
      if (pair[0] === sid) return json_({ ok: false, err: 'already', name: info.name, tm: pair[1] || '' });
    }
    var tm = nowTime_();
    var uid = 'sf' + Date.now().toString(36) + key_().slice(0, 6);
    s.ev.appendRow([uid, Date.now(), String(f.v[3] || nowDate_()), String(f.v[1]), sid,
      info.name, 'P', 'self', tm, 'online', Date.now()]);
    hits.push(sid + '@' + tm);
    s.se.getRange(f.row, 7).setValue(hits.join(','));
    return json_({ ok: true, name: info.name, cls: String(f.v[2] || ''), tm: tm });
  } finally {
    lock.releaseLock();
  }
}

function newSession_(body) {
  var s = sheets_();
  var cid = String(body.c || '');
  if (!cid) return json_({ ok: false, err: 'no class' });
  var mins = Math.max(5, Math.min(240, Number(body.mins) || 30));
  var key = key_();
  var now = Date.now();
  var exp = now + mins * 60000;
  s.se.appendRow([key, cid, className_(s, cid), nowDate_(), now, exp, '']);
  return json_({ ok: true, key: key, exp: exp });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, err: 'bad json' }); }
  if (body.op === 'chk') return checkin_(body);          /* student endpoint — no TOKEN */
  if (body.t !== TOKEN) return json_({ ok: false, err: 'bad token' });
  if (body.op === 'newsession') return newSession_(body);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var s = sheets_();
    var saved = 0, emailed = 0;
    var newEvents = [];
    if (body.ev && body.ev.length) {
      var have = {};
      var last = s.ev.getLastRow();
      if (last > 1) {
        var uids = s.ev.getRange(2, 1, last - 1, 1).getValues();
        for (var i = 0; i < uids.length; i++) have[String(uids[i][0])] = 1;
      }
      var now = Date.now();
      var add = [];
      for (var j = 0; j < body.ev.length; j++) {
        var v = body.ev[j];
        if (have[v.u]) continue;
        add.push([v.u, v.ts, v.d, v.c, v.s, v.n, v.st, v.via, v.tm, v.dev || '', now]);
        have[v.u] = 1;
        newEvents.push(v);
      }
      if (add.length) {
        s.ev.getRange(s.ev.getLastRow() + 1, 1, add.length, EV_HEAD.length).setValues(add);
        saved = add.length;
      }
    }
    if (body.roster && Number(body.roster.ts) > Number(s.me.getRange(1, 2).getValue() || 0)) {
      s.ro.clearContents();
      var rows = body.roster.rows || [];
      if (rows.length) s.ro.getRange(1, 1, rows.length, 6).setValues(rows);
      s.me.getRange(1, 2).setValue(Number(body.roster.ts));
    }
    if (SEND_EMAILS && newEvents.length) {
      var quota = 0;
      try { quota = MailApp.getRemainingDailyQuota(); } catch (qe) { quota = 0; }
      if (quota > 0) {
        var rmap = rosterMap_(s);
        for (var k = 0; k < newEvents.length && quota > 0; k++) {
          var ev2 = newEvents[k];
          if (ev2.st !== 'P' && ev2.st !== 'L') continue;
          if (ev2.via !== 'scan' && ev2.via !== 'tap') continue;
          var info = rmap[String(ev2.s)];
          if (!info || !info.em || info.em.indexOf('@') < 1) continue;
          try {
            MailApp.sendEmail({
              to: info.em,
              name: EMAIL_FROM_NAME,
              subject: 'Attendance: ' + ev2.n + ' — ' + word_(ev2.st) + ' (' + ev2.d + ')',
              body: 'Hi,\n\n' + ev2.n + ' was marked ' + word_(ev2.st) +
                    (info.cls ? ' in ' + info.cls : '') +
                    ' on ' + ev2.d + ' at ' + ev2.tm + '.\n\n' +
                    'This is an automated confirmation from Presente. Please do not reply.'
            });
            emailed++; quota--;
          } catch (mailErr) { quota = 0; }
        }
      }
    }
    return json_({ ok: true, now: Date.now(), saved: saved, emailed: emailed });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Scale Up — bookings backend (Google Apps Script)
 *
 * What it does
 *  - POST  : receives a booking from the website, validates it, blocks double-booking,
 *            saves it to the "Bookings" sheet, adds it to your Google Calendar,
 *            and emails you all the details (lands in your Gmail inbox, marked important).
 *  - GET   : ?action=slots  → which call times are already taken + how many client places are taken,
 *            so the website greys out booked times and shows the real "places left" counter.
 *  - Daily : 8:00 AM (Cairo) email with today's calls, only on days that have calls.
 *  - Admin : scaleupreach.com/admin → change the website video (upload a file to your Drive, or paste a
 *            YouTube / Drive link), its caption, and the "places taken" counter. Password is emailed to you by setup().
 *
 * Setup (one time): run setup() from the editor, approve the permissions, then
 * Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → copy the URL
 * into CONFIG.API_URL in index.html.
 */

const CFG = {
  OWNER_TZ: 'Africa/Cairo',
  CALL_MINUTES: 20,
  BOOKINGS: 'Bookings',
  SETTINGS: 'Settings',
  MAX_DAYS_AHEAD: 21,
  MIN_HOURS_AHEAD: 6,
};

const HEADERS = ['Received (Cairo)', 'Call time (Cairo)', 'Call time (client)', 'Client timezone', 'Name', 'WhatsApp',
  'Channel', 'Best videos', 'Socials', 'Sells', 'Posting', 'Goal', 'Status', 'Start UTC', 'Calendar event'];

/* ---------- one-time setup ---------- */
function setup() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CFG.BOOKINGS);
  if (!sh) sh = ss.insertSheet(CFG.BOOKINGS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0B0A09').setFontColor('#F3EFE8');
  }
  let st = ss.getSheetByName(CFG.SETTINGS);
  if (!st) {
    st = ss.insertSheet(CFG.SETTINGS);
    st.getRange(1, 1, 4, 2).setValues([
      ['setting', 'value'],
      ['notify_email', Session.getEffectiveUser().getEmail()],
      ['slots_taken', 2],
      ['slots_total', 5],
    ]);
    st.getRange(1, 1, 1, 2).setFontWeight('bold');
    st.setColumnWidth(1, 160); st.setColumnWidth(2, 280);
  }
  ['vsl', 'vsl_caption'].forEach(k => { if (!st.getDataRange().getValues().some(r => r[0] === k)) st.appendRow([k, '']); });

  // admin password: made once, emailed to you
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PASS')) {
    const pass = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    props.setProperty('ADMIN_PASS', pass);
    GmailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Scale Up — admin password',
      'Admin panel: https://scaleupreach.com/admin\nPassword: ' + pass + '\n\nTo change it: run resetAdminPassword() in Apps Script.',
      { name: 'Scale Up Bookings' });
  }
  DriveApp.getRootFolder(); // asks for Drive permission now (video uploads go to your Drive)

  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);

  // daily digest trigger (replace any old one)
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyDigest').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyDigest').timeBased().atHour(8).everyDays(1).inTimezone(CFG.OWNER_TZ).create();

  Logger.log('Sheet: ' + ss.getUrl());
  return ss.getUrl();
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('Scale Up — Bookings');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function settings_() {
  const st = getSpreadsheet_().getSheetByName(CFG.SETTINGS);
  const out = {};
  st.getDataRange().getValues().slice(1).forEach(([k, v]) => { if (k) out[String(k).trim()] = v; });
  return {
    email: String(out.notify_email || Session.getEffectiveUser().getEmail()).trim(),
    taken: Number(out.slots_taken) || 0,
    total: Number(out.slots_total) || 5,
    vsl: String(out.vsl || '').trim(),
    caption: String(out.vsl_caption || '').trim(),
  };
}

function setSetting_(key, value) {
  const st = getSpreadsheet_().getSheetByName(CFG.SETTINGS);
  const rows = st.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][0]).trim() === key) { st.getRange(i + 1, 2).setValue(value); return; }
  st.appendRow([key, value]);
}

/* ---------- web endpoints ---------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'slots';
  if (action !== 'slots') return json_({ ok: false, error: 'unknown_action' });
  const s = settings_();
  return json_({ ok: true, booked: bookedStarts_(), slotsTaken: s.taken, slotsTotal: s.total, vsl: s.vsl, vslCaption: s.caption });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  if (d.action && String(d.action).indexOf('admin_') === 0) return admin_(d);

  // bots: hidden field must stay empty
  if (d.company_website) return json_({ ok: true });

  const v = validate_(d);
  if (v) return json_({ ok: false, error: 'invalid', field: v });

  // same phone can't spam (10 min)
  const cache = CacheService.getScriptCache();
  const key = 'p_' + d.phone.replace(/\D/g, '');
  if (cache.get(key)) return json_({ ok: false, error: 'too_many' });

  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (err) { return json_({ ok: false, error: 'busy' }); }
  try {
    const start = new Date(d.startUtc);
    if (bookedStarts_().indexOf(start.toISOString()) !== -1) return json_({ ok: false, error: 'slot_taken' });

    const end = new Date(start.getTime() + CFG.CALL_MINUTES * 60000);
    const s = settings_();
    const ownerTime = Utilities.formatDate(start, CFG.OWNER_TZ, 'EEE d MMM yyyy · h:mm a');
    const clientTime = Utilities.formatDate(start, d.tz, 'EEE d MMM · h:mm a');
    const received = Utilities.formatDate(new Date(), CFG.OWNER_TZ, 'yyyy-MM-dd HH:mm');

    let eventId = '';
    try {
      const ev = CalendarApp.getDefaultCalendar().createEvent('Scale Up call — ' + d.name, start, end, {
        description: plainDetails_(d, ownerTime, clientTime),
      });
      ev.removeAllReminders(); ev.addPopupReminder(30); ev.addPopupReminder(5);
      eventId = ev.getId();
    } catch (err) { console.error('calendar', err); }

    getSpreadsheet_().getSheetByName(CFG.BOOKINGS).appendRow([
      received, ownerTime, clientTime, d.tz, d.name, "'" + d.phone, d.channel, d.videos, d.socials || '',
      d.sell, d.freq, d.goal, 'New', start.toISOString(), eventId,
    ]);

    sendEmail_(s.email, d, ownerTime, clientTime);
    cache.put(key, '1', 600);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */
const TZS = ['Asia/Riyadh', 'Asia/Dubai', 'Africa/Cairo'];
const SELL = ['كورس', 'كوتشينج', 'خدمة', 'منتج', 'غير كده'];
const FREQ = ['كل أسبوع', 'كل أسبوعين', 'مرة في الشهر', 'لسه هبدأ'];
const GOAL = ['عملاء جداد', 'شهرة أكتر', 'مبيعات كورس'];

function validate_(d) {
  const str = (x, max) => typeof x === 'string' && x.trim().length > 0 && x.length <= max;
  const url = x => /^https?:\/\/\S+\.\S+/i.test(String(x || '').trim());
  if (!str(d.name, 80) || d.name.trim().length < 2) return 'name';
  const digits = String(d.phone || '').replace(/\D/g, '');
  if (!/^\+\d{1,4} \d{7,12}$/.test(String(d.phone || '')) || digits.length < 8 || digits.length > 15) return 'phone';
  if (!url(d.channel) || d.channel.length > 300) return 'channel';
  if (!str(d.videos, 1500) || !String(d.videos).split(/\s+/).some(url)) return 'videos';
  if (d.socials && String(d.socials).length > 300) return 'socials';
  if (SELL.indexOf(d.sell) === -1) return 'sell';
  if (FREQ.indexOf(d.freq) === -1) return 'freq';
  if (GOAL.indexOf(d.goal) === -1) return 'goal';
  if (TZS.indexOf(d.tz) === -1) return 'tz';
  const start = new Date(d.startUtc), now = Date.now();
  if (isNaN(start)) return 'time';
  if (start.getTime() < now + CFG.MIN_HOURS_AHEAD * 3600000) return 'time';
  if (start.getTime() > now + CFG.MAX_DAYS_AHEAD * 86400000) return 'time';
  if (start.getUTCMinutes() !== 0) return 'time';
  return null;
}

function bookedStarts_() {
  const sh = getSpreadsheet_().getSheetByName(CFG.BOOKINGS);
  if (!sh || sh.getLastRow() < 2) return [];
  const statusCol = HEADERS.indexOf('Status'), startCol = HEADERS.indexOf('Start UTC');
  const now = Date.now();
  return sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues()
    .filter(r => String(r[statusCol]).toLowerCase() !== 'cancelled' && r[startCol])
    .map(r => new Date(r[startCol]))
    .filter(dt => !isNaN(dt) && dt.getTime() > now)
    .map(dt => dt.toISOString());
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function plainDetails_(d, ownerTime, clientTime) {
  return [
    'الاسم: ' + d.name,
    'واتساب: ' + d.phone + '  →  https://wa.me/' + d.phone.replace(/\D/g, ''),
    'الميعاد (مصر): ' + ownerTime,
    'الميعاد (عند العميل): ' + clientTime + ' — ' + d.tz,
    'القناة: ' + d.channel,
    'أحسن فيديوهات:\n' + d.videos,
    'حسابات: ' + (d.socials || '—'),
    'بيبيع: ' + d.sell + ' · بينزّل: ' + d.freq + ' · عايز: ' + d.goal,
  ].join('\n');
}

function sendEmail_(to, d, ownerTime, clientTime) {
  const wa = 'https://wa.me/' + d.phone.replace(/\D/g, '');
  const links = String(d.videos).split(/\s+/).filter(x => /^https?:\/\//i.test(x))
    .map(u => '<a href="' + esc_(u) + '" style="color:#FF4A2B">' + esc_(u) + '</a>').join('<br>');
  const row = (k, v) => '<tr><td style="padding:8px 0;color:#7a746b;white-space:nowrap;vertical-align:top">' + k +
    '</td><td style="padding:8px 16px 8px 0;color:#0B0A09">' + v + '</td></tr>';
  const html =
    '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;max-width:560px;margin:auto;background:#F3EFE8;padding:24px;border-radius:16px">' +
    '<div style="font-size:13px;color:#7a746b">Scale Up · حجز جديد</div>' +
    '<h2 style="margin:8px 0 4px;color:#0B0A09">' + esc_(d.name) + ' حجز مكالمة</h2>' +
    '<div style="font-size:18px;color:#FF4A2B;font-weight:bold;margin-bottom:16px">' + esc_(ownerTime) + ' (بتوقيت مصر)</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:15px">' +
    row('الواتساب', '<a href="' + wa + '" style="color:#FF4A2B">' + esc_(d.phone) + '</a>') +
    row('عند العميل', esc_(clientTime) + ' · ' + esc_(d.tz)) +
    row('القناة', '<a href="' + esc_(d.channel) + '" style="color:#FF4A2B">' + esc_(d.channel) + '</a>') +
    row('أحسن فيديوهات', links) +
    row('حسابات', esc_(d.socials || '—')) +
    row('بيبيع', esc_(d.sell)) + row('بينزّل', esc_(d.freq)) + row('عايز', esc_(d.goal)) +
    '</table>' +
    '<a href="' + wa + '" style="display:inline-block;margin-top:20px;background:#0B0A09;color:#F3EFE8;text-decoration:none;padding:12px 22px;border-radius:999px;font-weight:bold">كلّمه على الواتساب</a>' +
    '<div style="margin-top:20px;font-size:12px;color:#7a746b">الميعاد اتضاف لجوجل كالندر بتاعك. كل الحجوزات في شيت «Scale Up — Bookings».</div>' +
    '</div>';
  GmailApp.sendEmail(to, '🟠 حجز جديد — ' + d.name + ' · ' + ownerTime, plainDetails_(d, ownerTime, clientTime), {
    htmlBody: html, name: 'Scale Up Bookings',
  });
}

function dailyDigest() {
  const sh = getSpreadsheet_().getSheetByName(CFG.BOOKINGS);
  if (!sh || sh.getLastRow() < 2) return;
  const today = Utilities.formatDate(new Date(), CFG.OWNER_TZ, 'yyyy-MM-dd');
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues().filter(r => {
    const st = new Date(r[HEADERS.indexOf('Start UTC')]);
    return !isNaN(st) && String(r[HEADERS.indexOf('Status')]).toLowerCase() !== 'cancelled' &&
      Utilities.formatDate(st, CFG.OWNER_TZ, 'yyyy-MM-dd') === today;
  });
  if (!rows.length) return;
  const list = rows.map(r => '• ' + r[1] + ' — ' + r[4] + ' (' + r[5] + ')').join('\n');
  GmailApp.sendEmail(settings_().email, '📞 مكالمات النهارده (' + rows.length + ')', 'عندك النهارده:\n\n' + list, { name: 'Scale Up Bookings' });
}

/* ---------- admin panel (scaleupreach.com/admin) ---------- */
function admin_(d) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('admin_fails') || 0);
  if (fails >= 8) return json_({ ok: false, error: 'locked' }); // 8 wrong tries → wait 15 min
  const pass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS');
  if (!pass || String(d.pass || '') !== pass) {
    cache.put('admin_fails', String(fails + 1), 900);
    return json_({ ok: false, error: 'wrong_pass' });
  }
  const s = settings_();
  switch (d.action) {
    case 'admin_login':
      return json_({ ok: true, taken: s.taken, total: s.total, vsl: s.vsl, caption: s.caption, email: s.email });

    case 'admin_save': {
      if (d.taken !== undefined) {
        const t = Math.round(Number(d.taken));
        if (!(t >= 0 && t <= s.total)) return json_({ ok: false, error: 'bad_taken' });
        setSetting_('slots_taken', t);
      }
      if (d.caption !== undefined) setSetting_('vsl_caption', String(d.caption).slice(0, 80));
      if (d.vsl !== undefined) {
        const v = normVsl_(d.vsl);
        if (v === null) return json_({ ok: false, error: 'bad_link' });
        if (v.indexOf('drive:') === 0) shareDrive_(v.slice(6));
        setSetting_('vsl', v);
      }
      return json_({ ok: true, vsl: settings_().vsl });
    }

    // big files go straight to your Google Drive in 5 MB pieces
    case 'admin_upload_start': {
      const size = Number(d.size);
      if (!(size > 0 && size <= 2 * 1024 * 1024 * 1024)) return json_({ ok: false, error: 'bad_size' });
      const mime = /^video\//.test(d.mime || '') ? d.mime : 'video/mp4';
      const name = 'Scale Up website video — ' + Utilities.formatDate(new Date(), CFG.OWNER_TZ, 'yyyy-MM-dd HH:mm') +
        ' — ' + String(d.name || 'video').slice(0, 80);
      const r = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', {
        method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': mime, 'X-Upload-Content-Length': String(size) },
        payload: JSON.stringify({ name: name, mimeType: mime }),
      });
      const loc = r.getHeaders()['Location'] || r.getHeaders()['location'];
      if (!loc) return json_({ ok: false, error: 'drive_' + r.getResponseCode() });
      const id = Utilities.getUuid();
      cache.put('up_' + id, JSON.stringify({ loc: loc, size: size }), 21600);
      return json_({ ok: true, id: id });
    }

    case 'admin_upload_chunk': {
      const up = JSON.parse(cache.get('up_' + d.id) || 'null');
      if (!up) return json_({ ok: false, error: 'upload_expired' });
      const bytes = Utilities.base64Decode(d.data);
      const from = Number(d.offset), to = from + bytes.length - 1;
      const r = UrlFetchApp.fetch(up.loc, {
        method: 'put', muteHttpExceptions: true, contentType: 'application/octet-stream',
        headers: { 'Content-Range': 'bytes ' + from + '-' + to + '/' + up.size }, payload: bytes,
      });
      const code = r.getResponseCode();
      if (code === 308) return json_({ ok: true, done: false });
      if (code === 200 || code === 201) {
        const fileId = JSON.parse(r.getContentText()).id;
        shareDrive_(fileId);
        setSetting_('vsl', 'drive:' + fileId);
        cache.remove('up_' + d.id);
        return json_({ ok: true, done: true, vsl: 'drive:' + fileId });
      }
      return json_({ ok: false, error: 'drive_' + code });
    }
  }
  return json_({ ok: false, error: 'unknown_action' });
}

// accepts: YouTube link, Google Drive file link, direct https .mp4 link, or empty (= back to the built-in video)
function normVsl_(x) {
  x = String(x || '').trim();
  if (!x || x === 'default') return '';
  if (/^drive:[\w-]{20,}$/.test(x)) return x;
  let m = x.match(/(?:youtube\.com\/(?:watch\?v=|shorts\/|embed\/|live\/)|youtu\.be\/)([\w-]{11})/);
  if (m) return 'yt:' + m[1];
  m = x.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:export=\w+&)?id=)([\w-]{20,})/);
  if (m) return 'drive:' + m[1];
  if (/^https:\/\/\S+\.(mp4|webm|mov)(\?\S*)?$/i.test(x)) return x;
  return null;
}

// visitors must be able to watch it without a Google account
function shareDrive_(fileId) {
  try { DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (err) { console.error('share', err); }
}

function resetAdminPassword() {
  PropertiesService.getScriptProperties().deleteProperty('ADMIN_PASS');
  setup();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* Run this once from the editor to check the email arrives in your inbox. */
function testEmail() {
  sendEmail_(settings_().email, {
    name: 'تجربة', phone: '(test — no number)', channel: 'https://youtube.com/@example', videos: 'https://youtu.be/example',
    socials: '', sell: 'كوتشينج', freq: 'كل أسبوع', goal: 'عملاء جداد', tz: 'Asia/Riyadh',
  }, 'تجربة', 'تجربة');
}

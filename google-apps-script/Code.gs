/**
 * يوم المبرمج العراقي 2026 — نظام تسجيل الشخصيات VIP
 * الخلفية البرمجية (Google Apps Script + Google Sheets).
 *
 * طريقة التركيب: راجع README.md في جذر المستودع خطوة بخطوة.
 *
 * أمان:
 * - كل عمليات الإدارة (adminList / adminUpdateStatus / adminUpdateVip / adminDeleteVip)
 *   تتطلب كلمة مرور تُقارن بقيمة محفوظة في Script Properties (ADMIN_PASSWORD)،
 *   ولا تُحفظ أي كلمة مرور داخل الكود أو المستودع.
 * - عملية verify وaddVip عامتان لكن لا تُرجعان بيانات إدارية حساسة.
 */

var SHEET_NAME = 'VIPs';
var EVENT_YEAR = '2026';
var HEADERS = [
  'vip_id', 'vip_code', 'full_name', 'workplace', 'position', 'contact_method',
  'status', 'priority', 'entered_by_name', 'entered_by_position', 'created_at', 'updated_at', 'attendance_time',
  'needs_letter', 'attended_day1', 'attended_day1_time', 'attended_day2', 'attended_day2_time'
];
var VALID_STATUSES = ['مدعو', 'تم التأكيد', 'حضر', 'لم يحضر'];
var VALID_PRIORITIES = ['عالية', 'متوسطة', 'عادية'];
var DEFAULT_PRIORITY = 'عادية';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    // ترحيل: إذا أُضيفت أعمدة جديدة بالكود (مثل priority) بعد إنشاء الجدول، كمّل عناوينها بدون فقدان البيانات الحالية
    var existingCols = sheet.getLastColumn();
    if (existingCols < HEADERS.length) {
      sheet.getRange(1, existingCols + 1, 1, HEADERS.length - existingCols)
        .setValues([HEADERS.slice(existingCols)]);
    }
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function checkAdmin_(password) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return !!real && !!password && password === real;
}

/**
 * رمز VIP عشوائي (وليس متسلسلاً) حتى لا يكشف ترتيب أو أولوية التسجيل
 * (مثال: VIP-2026-0001 يوحي بأن صاحبه أول المسجَّلين). يتحقق من عدم
 * التكرار مقابل الرموز الموجودة فعلاً بالجدول قبل اعتماده.
 */
function randomVipCode_(sheet) {
  var existing = {};
  readAllRows_(sheet).forEach(function (r) { existing[String(r.vip_code)] = true; });
  var code, attempts = 0;
  do {
    var n = Math.floor(Math.random() * 90000) + 10000; // 5 أرقام عشوائية: 10000–99999
    code = 'VIP-' + EVENT_YEAR + '-' + n;
    attempts++;
  } while (existing[code] && attempts < 50);
  return code;
}

function readAllRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values.map(function (row, i) {
    var obj = {};
    HEADERS.forEach(function (h, idx) { obj[h] = row[idx]; });
    obj._row = i + 2; // 1-based sheet row number
    return obj;
  });
}

function findRowByCode_(sheet, code) {
  var rows = readAllRows_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].vip_code) === String(code)) return rows[i];
  }
  return null;
}

/* ---------- Fuzzy name matching (for duplicate-invite checking) ---------- */

function normalizeArabicText_(s) {
  return String(s || '')
    .replace(/[ً-ْٰـ]/g, '')
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/\s+/g, ' ')
    .trim();
}

function levenshtein_(a, b) {
  var m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  var prev = [];
  for (var j = 0; j <= n; j++) prev[j] = j;
  for (var i = 1; i <= m; i++) {
    var cur = [i];
    for (var j = 1; j <= n; j++) {
      cur[j] = a.charAt(i - 1) === b.charAt(j - 1)
        ? prev[j - 1]
        : 1 + Math.min(prev[j - 1], prev[j], cur[j - 1]);
    }
    prev = cur;
  }
  return prev[n];
}

/**
 * تشابه على مستوى الكلمات: نسبة كلمات الاسم الأقصر اللي لها كلمة مقاربة
 * بالاسم الآخر. يلتقط اختلاف ترتيب الكلمات أو نقصان/زيادة لقب أو اسم
 * أب لم يلتقطه التشابه الحرفي للسلسلة كاملة (مثال: "احمد كريم علي" مقابل
 * "الدكتور احمد كريم علي حسين").
 */
function tokenSimilarity_(a, b) {
  var ta = normalizeArabicText_(a).split(' ').filter(function (w) { return w.length > 0; });
  var tb = normalizeArabicText_(b).split(' ').filter(function (w) { return w.length > 0; });
  if (ta.length === 0 || tb.length === 0) return 0;
  var matched = 0;
  ta.forEach(function (wa) {
    var best = 0;
    tb.forEach(function (wb) {
      var sim = wa === wb ? 1 : 1 - (levenshtein_(wa, wb) / Math.max(wa.length, wb.length));
      if (sim > best) best = sim;
    });
    if (best >= 0.75) matched++;
  });
  return matched / Math.min(ta.length, tb.length);
}

/** تشابه بين 0 و1: مطابقة جزئية (بادئة/احتواء) تُعامل كتطابق شبه تام، وإلا أعلى قيمة بين تقارب السلسلة كاملة وتقارب الكلمات */
function nameSimilarity_(a, b) {
  var na = normalizeArabicText_(a), nb = normalizeArabicText_(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  if (na.indexOf(nb) !== -1 || nb.indexOf(na) !== -1) return 0.95;
  var maxLen = Math.max(na.length, nb.length);
  var wholeSim = maxLen === 0 ? 0 : 1 - (levenshtein_(na, nb) / maxLen);
  var tokSim = tokenSimilarity_(a, b);
  return Math.max(wholeSim, tokSim);
}

var DUP_SIMILARITY_THRESHOLD = 0.6;

/**
 * يستخدمها العضو (بدون كلمة مرور) قبل إضافة شخصية VIP، للتأكد إذا كانت
 * مدعوّة مسبقاً من عضو آخر. المطابقة تقريبية (اسم مقارب) وليست حرفية،
 * حتى تلتقط الاختلافات البسيطة بالكتابة أو الألقاب.
 */
function handleCheckVip_(body) {
  var query = String(body.name || '').trim();
  if (normalizeArabicText_(query).length < 3) return { ok: true, matches: [] };

  var sheet = getSheet_();
  var rows = readAllRows_(sheet);
  var matches = rows
    .map(function (r) { return { row: r, score: nameSimilarity_(query, r.full_name) }; })
    .filter(function (s) { return s.score >= DUP_SIMILARITY_THRESHOLD; })
    .sort(function (a, b) { return b.score - a.score; })
    .slice(0, 5)
    .map(function (s) {
      return { full_name: s.row.full_name, workplace: s.row.workplace, entered_by_name: s.row.entered_by_name };
    });

  return { ok: true, matches: matches };
}

/* ---------- Public actions ---------- */

function handleAddVip_(body) {
  var fullName = String(body.full_name || '').trim();
  var workplace = String(body.workplace || '').trim();
  var position = String(body.position || '').trim();
  var contactMethod = String(body.contact_method || '').trim();
  var needsLetter = body.needs_letter ? 'نعم' : 'لا';
  var enteredByName = String(body.entered_by_name || '').trim();
  var enteredByPosition = String(body.entered_by_position || '').trim();

  if (!fullName) return { error: 'missing_full_name' };
  if (!workplace) return { error: 'missing_workplace' };
  if (!enteredByName || !enteredByPosition) return { error: 'missing_entrant' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var seq = sheet.getLastRow(); // header occupies row 1; used only as an internal id, never shown
    var vipCode = randomVipCode_(sheet);
    var now = new Date().toISOString();

    sheet.appendRow([
      seq, vipCode, fullName, workplace, position, contactMethod,
      'مدعو', DEFAULT_PRIORITY, enteredByName, enteredByPosition, now, now, '', needsLetter
    ]);

    return { ok: true, vip_id: seq, vip_code: vipCode, created_at: now };
  } finally {
    lock.releaseLock();
  }
}

/**
 * يستخدمها العضو (بدون كلمة مرور) لتحديد أولوية الحضور بين أكثر من شخصية
 * VIP أضافها بنفس الجلسة، بسبب محدودية المقاعد المخصصة. لا يعرض أو يعدّل
 * أي بيانات غير الأولوية نفسها.
 */
function handleSetPriorities_(body) {
  var items = Array.isArray(body.items) ? body.items : [];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var rows = readAllRows_(sheet);
    var priorityCol = HEADERS.indexOf('priority') + 1;
    var updatedCol = HEADERS.indexOf('updated_at') + 1;
    var now = new Date().toISOString();
    var updated = 0;

    items.forEach(function (item) {
      var priority = String(item.priority || '').trim();
      if (VALID_PRIORITIES.indexOf(priority) === -1) return;
      var row = rows.filter(function (r) { return r.vip_code === item.vip_code; })[0];
      if (!row) return;
      sheet.getRange(row._row, priorityCol).setValue(priority);
      sheet.getRange(row._row, updatedCol).setValue(now);
      updated++;
    });

    return { ok: true, updated: updated };
  } finally {
    lock.releaseLock();
  }
}

function handleVerify_(code) {
  code = String(code || '').trim();
  if (!code) return { found: false };
  var sheet = getSheet_();
  var row = findRowByCode_(sheet, code);
  if (!row) return { found: false };
  return {
    found: true,
    vip_code: row.vip_code,
    full_name: row.full_name,
    workplace: row.workplace,
    position: row.position || ''
  };
}

/* ---------- Admin actions (password required) ---------- */

function handleAdminList_(body) {
  if (!checkAdmin_(body.password)) return { error: 'unauthorized' };
  var sheet = getSheet_();
  var rows = readAllRows_(sheet).map(function (r) {
    return {
      vip_id: r.vip_id,
      vip_code: r.vip_code,
      full_name: r.full_name,
      workplace: r.workplace,
      position: r.position,
      contact_method: r.contact_method,
      needs_letter: r.needs_letter === 'نعم' ? 'نعم' : 'لا',
      status: r.status,
      priority: r.priority || DEFAULT_PRIORITY,
      entered_by_name: r.entered_by_name,
      entered_by_position: r.entered_by_position,
      created_at: r.created_at,
      updated_at: r.updated_at,
      attendance_time: r.attendance_time,
      attended_day1: r.attended_day1 === 'نعم' ? 'نعم' : 'لا',
      attended_day1_time: r.attended_day1_time || '',
      attended_day2: r.attended_day2 === 'نعم' ? 'نعم' : 'لا',
      attended_day2_time: r.attended_day2_time || ''
    };
  });
  return { ok: true, rows: rows };
}

/**
 * يستخدمها الإدمن لتأكيد أو إلغاء حضور شخصية VIP ليوم معيّن من يومي الفعالية
 * (9-10 أو 10-10) بشكل مستقل عن الحالة العامة (status). مصممة للاستخدام
 * السريع عند باب الفعالية.
 */
function handleAdminMarkAttendance_(body) {
  if (!checkAdmin_(body.password)) return { error: 'unauthorized' };
  var code = String(body.vip_code || '').trim();
  var day = String(body.day || '').trim();
  if (day !== '1' && day !== '2') return { error: 'invalid_day' };
  var attended = !!body.attended;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var row = findRowByCode_(sheet, code);
    if (!row) return { error: 'not_found' };

    var now = new Date().toISOString();
    var flagCol = HEADERS.indexOf('attended_day' + day) + 1;
    var timeCol = HEADERS.indexOf('attended_day' + day + '_time') + 1;
    var updatedCol = HEADERS.indexOf('updated_at') + 1;

    sheet.getRange(row._row, flagCol).setValue(attended ? 'نعم' : 'لا');
    sheet.getRange(row._row, timeCol).setValue(attended ? now : '');
    sheet.getRange(row._row, updatedCol).setValue(now);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function handleAdminUpdateStatus_(body) {
  if (!checkAdmin_(body.password)) return { error: 'unauthorized' };
  var code = String(body.vip_code || '').trim();
  var status = String(body.status || '').trim();
  if (VALID_STATUSES.indexOf(status) === -1) return { error: 'invalid_status' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var row = findRowByCode_(sheet, code);
    if (!row) return { error: 'not_found' };

    var now = new Date().toISOString();
    var statusCol = HEADERS.indexOf('status') + 1;
    var updatedCol = HEADERS.indexOf('updated_at') + 1;
    var attendCol = HEADERS.indexOf('attendance_time') + 1;

    sheet.getRange(row._row, statusCol).setValue(status);
    sheet.getRange(row._row, updatedCol).setValue(now);
    if (status === 'حضر') {
      sheet.getRange(row._row, attendCol).setValue(now);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function handleAdminUpdateVip_(body) {
  if (!checkAdmin_(body.password)) return { error: 'unauthorized' };
  var code = String(body.vip_code || '').trim();

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var row = findRowByCode_(sheet, code);
    if (!row) return { error: 'not_found' };

    var fields = ['full_name', 'workplace', 'position', 'contact_method'];
    fields.forEach(function (f) {
      if (typeof body[f] === 'string') {
        var col = HEADERS.indexOf(f) + 1;
        sheet.getRange(row._row, col).setValue(body[f].trim());
      }
    });
    if (typeof body.priority === 'string' && VALID_PRIORITIES.indexOf(body.priority.trim()) !== -1) {
      sheet.getRange(row._row, HEADERS.indexOf('priority') + 1).setValue(body.priority.trim());
    }
    if (typeof body.needs_letter === 'boolean') {
      sheet.getRange(row._row, HEADERS.indexOf('needs_letter') + 1).setValue(body.needs_letter ? 'نعم' : 'لا');
    }
    var updatedCol = HEADERS.indexOf('updated_at') + 1;
    sheet.getRange(row._row, updatedCol).setValue(new Date().toISOString());
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function handleAdminDeleteVip_(body) {
  if (!checkAdmin_(body.password)) return { error: 'unauthorized' };
  var code = String(body.vip_code || '').trim();

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var row = findRowByCode_(sheet, code);
    if (!row) return { error: 'not_found' };
    sheet.deleteRow(row._row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- HTTP entry points ---------- */

function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'verify') return jsonOut_(handleVerify_(e.parameter.code || ''));
    return jsonOut_({ error: 'unknown_action' });
  } catch (err) {
    return jsonOut_({ error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e.postData && e.postData.contents) || '{}');
    var action = body.action;
    if (action === 'addVip') return jsonOut_(handleAddVip_(body));
    if (action === 'checkVip') return jsonOut_(handleCheckVip_(body));
    if (action === 'verify') return jsonOut_(handleVerify_(body.code || ''));
    if (action === 'setPriorities') return jsonOut_(handleSetPriorities_(body));
    if (action === 'adminList') return jsonOut_(handleAdminList_(body));
    if (action === 'adminUpdateStatus') return jsonOut_(handleAdminUpdateStatus_(body));
    if (action === 'adminMarkAttendance') return jsonOut_(handleAdminMarkAttendance_(body));
    if (action === 'adminUpdateVip') return jsonOut_(handleAdminUpdateVip_(body));
    if (action === 'adminDeleteVip') return jsonOut_(handleAdminDeleteVip_(body));
    return jsonOut_({ error: 'unknown_action' });
  } catch (err) {
    return jsonOut_({ error: String(err) });
  }
}

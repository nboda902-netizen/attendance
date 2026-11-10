// نظام الحضور – الجزء الخلفي (Google Apps Script) – النسخة 3: أسماء كل موقع + كود الموظف + الزوار

// 1) ارجع اكتب الرقم السري بتاعك هنا (نفس اللي كنت حاططه قبل كده)
const PIN = '1234';
// 2) سيبه فاضي لو السكريبت مفتوح من جوه الشيت (الإضافات ← Apps Script).
//    لو فاتح السكريبت من script.google.com حط رقم الشيت بين العلامتين.
const SHEET_ID = '';

const FOLDER_KEY = 'PHOTOS';
// سماح بسيط لدقة الـ GPS (متر). لو عايز الفحص حرفي خلّيه 0
const ACC_TOL = 100;
const TZ = 'Africa/Cairo';
const SITES = ['صواري – 22', 'صواري – 17 و18', 'صواري – المصطفى', 'صواري – 35', 'مروج – 40', 'مروج – 37', 'جكرندا'];
const DAYS = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('الحضور')
    .addItem('تقفيل الشهر', 'buildClosing')
    .addToUi();
}

// شغّلها مرة واحدة، ولو شغّلتها تاني مفيش مشكلة (بتضيف الناقص بس)
function setup() {
  const ss = ss_();
  const mk = (n, h) => {
    let s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getLastRow() == 0) s.appendRow(h);
    return s;
  };
  const st = mk('الطاقم', ['الكود', 'الاسم', 'الوظيفة', 'الموقع الأساسي']);
  const ms = mk('المواقع', ['الموقع', 'خط العرض', 'خط الطول', 'النطاق بالمتر']);
  if (ms.getLastRow() == 1) SITES.forEach(n => addRow_(ms, { 'الموقع': n, 'النطاق بالمتر': 300 }));
  const lg = mk('التسجيلات', ['الوقت', 'الموقع', 'الاسم', 'النوع', 'المسافة (م)', 'الصورة', 'بواسطة', 'الكود']);
  if (head_(lg).indexOf('الكود') < 0) lg.getRange(1, lg.getLastColumn() + 1).setValue('الكود');
  mk('الزوار', ['الوقت', 'الموقع', 'النوع', 'الاسم', 'الجهة', 'جاي لمين / الغرض', 'وقت الخروج', 'سجّله']);
  // عمود الكود في الطاقم
  let hd = st.getRange(1, 1, 1, st.getLastColumn()).getValues()[0].map(String);
  if (hd.indexOf('الكود') < 0) { st.getRange(1, st.getLastColumn() + 1).setValue('الكود'); hd.push('الكود'); }
  // قايمة منسدلة لعمود الموقع الأساسي عشان الاسم ميتكتبش غلط
  const si = hd.indexOf('الموقع الأساسي') + 1;
  if (si > 0) {
    st.getRange(2, si, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(ms.getRange(2, colOf_(ms, 'الموقع') || 1, 19, 1), true).setAllowInvalid(false).build());
  }
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty(FOLDER_KEY)) p.setProperty(FOLDER_KEY, DriveApp.createFolder('صور الحضور').getId());
  setupClosing();
}

// بتعمل ورقة الإعدادات (لو مش موجودة)
function setupClosing() {
  const ss = ss_();
  const ex = ss.getSheetByName('الإعدادات');
  if (ex) {
    // شيت قديم: نضيف صف "يوم بداية الفترة" ونوضّح عنوان الشهر
    ex.getRange('A1').setValue('شهر نهاية الفترة (سنة-شهر)');
    if (ex.getRange('A6').getValue() === '') ex.getRange('A6').setValue('يوم بداية الفترة (1 = شهر ميلادي، 26 = من 26 لـ 25)').setFontWeight('bold');
    if (ex.getRange('B6').getValue() === '') ex.getRange('B6').setValue(26);
    return;
  }
  const s = ex || ss.insertSheet('الإعدادات');
  s.getRange('A1:A6').setValues([['شهر نهاية الفترة (سنة-شهر)'], ['بداية الدوام (الساعة)'], ['السماح بالتأخير (دقيقة)'], ['ساعات الدوام اليومي'], ['الإجازة الأسبوعية'], ['يوم بداية الفترة (1 = شهر ميلادي، 26 = من 26 لـ 25)']]).setFontWeight('bold');
  s.getRange('B6').setValue(26);
  s.getRange('B1').setNumberFormat('@').setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM'));
  s.getRange('B2:B5').setValues([[8], [15], [8], ['الجمعة']]);
  s.setRightToLeft(true);
  s.autoResizeColumn(1);
}

// ===== قراءة وكتابة الأوراق بأسماء الأعمدة: ترتيب الأعمدة مش مهم، ومسموح تضيف أعمدة =====
function head_(sh) { return sh.getLastColumn() < 1 ? [] : sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim()); }
function colOf_(sh, name) { return head_(sh).indexOf(name) + 1; }   // رقم العمود (0 لو مش موجود)
function table_(sh) {
  const v = sh.getDataRange().getValues(), h = (v[0] || []).map(x => String(x).trim());
  return v.slice(1).map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => { if (k) o[k] = r[j]; }); return o; });
}
function addRow_(sh, obj) {
  const h = head_(sh);
  Object.keys(obj).forEach(k => { if (h.indexOf(k) < 0) { h.push(k); sh.getRange(1, h.length).setValue(k); } });
  sh.appendRow(h.map(k => obj[k] === undefined ? '' : obj[k]));
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// قراءة الطاقم بأسماء الأعمدة (ترتيب الأعمدة مش مهم)
// تخزين مؤقت لقايمة الطاقم والمواقع (دقيقتين) عشان التسجيل يبقى أسرع. التعديل اليدوي في الشيت بيمسحه لوحده.
function cached_(key, sec, fn) {
  const c = CacheService.getScriptCache(), h = c.get(key);
  if (h) return JSON.parse(h);
  const v = fn();
  try { c.put(key, JSON.stringify(v), sec); } catch (e) {}
  return v;
}
function onEdit() { try { CacheService.getScriptCache().removeAll(['staff', 'sites']); } catch (e) {} }
function sites_(ss) { return cached_('sites', 120, () => table_(ss.getSheetByName('المواقع'))); }
function staff_(ss) { return cached_('staff', 120, () => staffRaw_(ss)); }
// اختياري: ضيفها كـ Trigger كل 5 دقايق عشان السيرفر ميبردش (Cold start)
function keepWarm() { UrlFetchApp.fetch(ScriptApp.getService().getUrl()); }

function staffRaw_(ss) {
  const v = ss.getSheetByName('الطاقم').getDataRange().getValues();
  const h = v[0].map(x => String(x).trim());
  let a = h.indexOf('الاسم'); if (a < 0) a = 0;
  const c = h.indexOf('الكود'), r = h.indexOf('الوظيفة'), s = h.indexOf('الموقع الأساسي');
  return v.slice(1).filter(x => x[a]).map(x => ({
    n: String(x[a]).trim(),
    c: c < 0 ? '' : String(x[c]).trim(),
    r: r < 0 ? '' : x[r],
    s: s < 0 ? '' : String(x[s]).trim()
  }));
}

// بيرجّع الأسماء (من غير الوظيفة) والمواقع للصفحة
function doGet() {
  const ss = ss_();
  const staff = staff_(ss).map(x => ({ n: x.n, c: x.c, s: x.s }));
  const sites = sites_(ss).filter(r => r['الموقع']).map(r => r['الموقع']);
  return out({ staff: staff, sites: sites });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'nearest') return nearest_(d);
    if (d.action === 'pin') return out(d.pin === PIN ? { ok: true } : { ok: false, msg: 'الرقم السري غلط' });
    if (d.action && d.action.indexOf('visit') === 0) return visit_(d);
    return att_(d);
  } catch (err) {
    return out({ ok: false, msg: 'خطأ: ' + err.message });
  }
}

// تسجيل حضور / انصراف الطاقم (بصورة)
function att_(d) {
  const lock = LockService.getScriptLock(), t0 = Date.now();
  try {
    // الصورة شرط للتسجيل: من غيرها السيرفر يرفض
    if (!d.photo || String(d.photo).length < 2000 || String(d.photo).indexOf('data:image') !== 0) return out({ ok: false, msg: 'الصورة مطلوبة للتسجيل' });
    const ss = ss_();
    const site = sites_(ss).find(r => r['الموقع'] == d.site);
    if (!site) return out({ ok: false, msg: 'الموقع غير معروف، امسح الـ QR بتاع الموقع تاني' });
    if (d.type !== 'حضور' && d.type !== 'انصراف') return out({ ok: false, msg: 'نوع غير صحيح' });
    const person = staff_(ss).find(x => x.n === d.name);
    if (!person) return out({ ok: false, msg: 'الاسم غير مسجل في الطاقم' });
    const sec = !!d.pin;
    if (sec && d.pin !== PIN) return out({ ok: false, msg: 'الرقم السري غلط' });
    let dist = '';
    if (!sec) {
      if (!site['خط العرض'] || !site['خط الطول']) return out({ ok: false, msg: 'إحداثيات الموقع لسه مش متسجلة في الشيت' });
      dist = Math.round(meters(+d.lat, +d.lng, +site['خط العرض'], +site['خط الطول']));
      if (isNaN(dist) || dist - Math.min(+d.acc || 0, ACC_TOL) > (+site['النطاق بالمتر'] || 300)) return out({ ok: false, msg: 'أنت بعيد عن الموقع (' + dist + ' متر)' });
    }
    lock.waitLock(20000);
    const now = new Date();
    const url = savePhoto_(d.name, now, d.photo);
    const sh = ss.getSheetByName('التسجيلات');
    addRow_(sh, { 'الوقت': now, 'الموقع': d.site, 'الاسم': d.name, 'النوع': d.type, 'المسافة (م)': dist, 'الصورة': url, 'بواسطة': sec ? 'الأمن' : 'ذاتي', 'الكود': person.c });
    return out({ ok: true, time: Utilities.formatDate(now, TZ, 'HH:mm'), row: sh.getLastRow(), ms: Date.now() - t0 });
  } catch (err) {
    return out({ ok: false, msg: 'خطأ: ' + err.message });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function savePhoto_(name, when, dataUrl) {
  const stamp = Utilities.formatDate(when, TZ, 'yyyy-MM-dd_HH-mm');
  const blob = Utilities.newBlob(Utilities.base64Decode(dataUrl.split(',')[1]), 'image/jpeg', name + '_' + stamp + '.jpg');
  return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty(FOLDER_KEY)).createFile(blob).getUrl();
}

// الزوار والإدارة (بدون صورة) – للأمن بس بالرقم السري
function visit_(d) {
  if (d.pin !== PIN) return out({ ok: false, msg: 'الرقم السري غلط' });
  const lock = LockService.getScriptLock();
  try {
    const ss = ss_(), sh = ss.getSheetByName('الزوار'), now = new Date();
    const f = (t, p) => Utilities.formatDate(t, TZ, p);
    if (d.action === 'visit_in') {
      if (!d.name || !String(d.name).trim()) return out({ ok: false, msg: 'اكتب اسم الزائر' });
      lock.waitLock(20000);
      addRow_(sh, { 'الوقت': now, 'الموقع': d.site, 'النوع': d.type || 'زائر', 'الاسم': String(d.name).trim(), 'الجهة': d.entity || '', 'جاي لمين / الغرض': d.who || '', 'وقت الخروج': '', 'سجّله': 'الأمن' });
      return out({ ok: true, time: f(now, 'HH:mm') });
    }
    if (d.action === 'visit_out') {
      const row = +d.row;
      if (!(row >= 2 && row <= sh.getLastRow())) return out({ ok: false, msg: 'سجل غير موجود' });
      lock.waitLock(20000);
      const cs = colOf_(sh, 'الموقع'), co = colOf_(sh, 'وقت الخروج');
      if (!cs || !co) return out({ ok: false, msg: 'عناوين ورقة الزوار ناقصة' });
      if (sh.getRange(row, cs).getValue() !== d.site) return out({ ok: false, msg: 'السجل ده من موقع تاني' });
      if (sh.getRange(row, co).getValue() !== '') return out({ ok: false, msg: 'الخروج متسجل قبل كده' });
      sh.getRange(row, co).setValue(now);
      return out({ ok: true, time: f(now, 'HH:mm') });
    }
    // visits: سجل النهاردة + اللي لسه جوه من آخر يومين
    const t = table_(sh), res = [], today = f(now, 'yyyy-MM-dd');
    for (let i = t.length - 1; i >= 0 && res.length < 60; i--) {
      const r = t[i];
      if (r['الموقع'] !== d.site || !(r['الوقت'] instanceof Date)) continue;
      const day = f(r['الوقت'], 'yyyy-MM-dd'), open = !(r['وقت الخروج'] instanceof Date);
      if (!(day === today || (open && now - r['الوقت'] < 172800000))) continue;
      res.push({ row: r._row, t: r['النوع'], n: r['الاسم'], e: r['الجهة'], w: r['جاي لمين / الغرض'], i: f(r['الوقت'], 'HH:mm'), o: open ? '' : f(r['وقت الخروج'], 'HH:mm'), d: day === today ? '' : day });
    }
    return out({ ok: true, list: res });
  } catch (err) {
    return out({ ok: false, msg: 'خطأ: ' + err.message });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function meters(a, b, c, d) {
  const R = 6371000, t = Math.PI / 180, x = (c - a) * t, y = (d - b) * t;
  const h = Math.sin(x / 2) ** 2 + Math.cos(a * t) * Math.cos(c * t) * Math.sin(y / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(h));
}

// ================= تقفيل الشهر =================
function pad_(n) { return ('0' + n).slice(-2); }
function hm_(m) { return pad_(Math.floor(m / 60)) + ':' + pad_(m % 60); }
function r2_(x) { return Math.round(x * 100) / 100; }

function readCfg_(ss) {
  const s = ss.getSheetByName('الإعدادات');
  if (!s) throw new Error('شغّل setupClosing الأول');
  const v = s.getRange('B1:B6').getValues().map(r => r[0]);
  let m = v[0];
  if (m instanceof Date) m = Utilities.formatDate(m, TZ, 'yyyy-MM');
  m = String(m).trim();
  if (!/^\d{4}-\d{1,2}$/.test(m)) throw new Error('الشهر في ورقة الإعدادات لازم يبقى بالشكل 2026-10');
  let sd = Math.floor(+v[5]);
  if (!(sd >= 1 && sd <= 28)) sd = 1;
  return { month: m, start: +v[1], grace: +v[2], hours: +v[3], off: String(v[4]).trim(), startDay: sd };
}

function write_(ss, name, head, rows) {
  const s = ss.getSheetByName(name) || ss.insertSheet(name);
  s.clear();
  s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#0b3b3a').setFontColor('#ffffff');
  if (rows.length) s.getRange(2, 1, rows.length, head.length).setValues(rows);
  s.setFrozenRows(1);
  s.setRightToLeft(true);
  s.autoResizeColumns(1, head.length);
  return s;
}

function buildClosing() {
  const ss = ss_(), cfg = readCfg_(ss);
  const parts = cfg.month.split('-').map(Number), Y = parts[0], M = parts[1];
  const mp = Y + '-' + pad_(M);
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  // الفترة: لو يوم البداية 26 يبقى شهر 11 = من 26/10 لـ 25/11 (الشهر بيتسمّى بشهر النهاية)
  const sd = cfg.startDay, period = [];
  const a = sd > 1 ? new Date(Y, M - 2, sd) : new Date(Y, M - 1, 1);
  const b = sd > 1 ? new Date(Y, M - 1, sd - 1) : new Date(Y, M, 0);
  for (let t = new Date(a); t <= b; t = new Date(t.getFullYear(), t.getMonth(), t.getDate() + 1))
    period.push({ s: t.getFullYear() + '-' + pad_(t.getMonth() + 1) + '-' + pad_(t.getDate()), w: t.getDay() });
  const startStr = period[0].s, endStr = period[period.length - 1].s;
  const label = sd > 1 ? ' (من ' + pad_(a.getDate()) + '.' + pad_(a.getMonth() + 1) + ' إلى ' + pad_(b.getDate()) + '.' + pad_(b.getMonth() + 1) + ')' : '';
  const offIdx = DAYS.indexOf(cfg.off);
  const limit = cfg.start * 60 + cfg.grace;
  const staff = staff_(ss);

  // تجميع التسجيلات: أول حضور وآخر انصراف لكل شخص في كل يوم
  const g = {}, unv = {};
  table_(ss.getSheetByName('التسجيلات')).forEach(r => {
    if (!(r['الوقت'] instanceof Date)) return;
    const d = Utilities.formatDate(r['الوقت'], TZ, 'yyyy-MM-dd');
    if (d < startStr || d > endStr) return;
    const m = +Utilities.formatDate(r['الوقت'], TZ, 'HH') * 60 + +Utilities.formatDate(r['الوقت'], TZ, 'mm');
    // شبكة أمان: أي تسجيل ذاتي من غير صورة (مثلاً من نسخة قديمة) مش بيتحسب ويتعلّم للمراجعة
    if (!(r['الصورة'] || String(r['بواسطة']).trim() !== 'ذاتي')) { (unv[r['الاسم']] = unv[r['الاسم']] || []).push({ d: d, t: r['النوع'], m: m, s: r['الموقع'] }); return; }
    const k = r['الاسم'] + '|' + d;
    const e = g[k] = g[k] || { in: null, out: null, site: '', photo: '' };
    if (r['النوع'] === 'حضور') {
      if (e.in === null || m < e.in) { e.in = m; e.site = r['الموقع']; e.photo = r['الصورة']; }
    } else if (r['النوع'] === 'انصراف') {
      if (e.out === null || m > e.out) e.out = m;
    }
  });

  const rows = [], det = [];
  staff.forEach(st => {
    const n = st.n;
    let pres = 0, abs = 0, hrs = 0, late = 0, lateMin = 0, ot = 0, miss = 0;
    for (const pd of period) {
      const ds = pd.s, e = g[n + '|' + ds];
      const isOff = pd.w === offIdx;
      if (!e || e.in === null) {
        if (!isOff && ds < today) abs++;
        continue;
      }
      pres++;
      let h = '', lm = 0, o = 0, note = '';
      if (e.in > limit) { late++; lm = e.in - cfg.start * 60; lateMin += lm; }
      if (e.out !== null && e.out > e.in) {
        h = (e.out - e.in) / 60; hrs += h;
        if (h > cfg.hours) { o = h - cfg.hours; ot += o; }
      } else if (ds < today) { miss++; note = 'ناقص انصراف'; }
      else note = 'لسه جوه الموقع';
      det.push([ds, n, e.site, hm_(e.in), e.out === null ? '' : hm_(e.out), h === '' ? '' : r2_(h), lm, r2_(o), note, e.photo]);
    }
    const u = unv[n] || [];
    u.forEach(x => det.push([x.d, n, x.s, x.t === 'حضور' ? hm_(x.m) : '', x.t === 'انصراف' ? hm_(x.m) : '', '', 0, 0, 'بدون صورة – غير محسوب', '']));
    rows.push([n, st.c, st.r, st.s, pres, abs, r2_(hrs), late, lateMin, r2_(ot), miss, u.length]);
  });
  det.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : (a[1] < b[1] ? -1 : 1));

  const s1 = write_(ss, 'تقفيل ' + mp + label,
    ['الاسم', 'الكود', 'الوظيفة', 'الموقع الأساسي', 'أيام الحضور', 'الغياب', 'إجمالي الساعات', 'مرات التأخير', 'دقائق التأخير', 'ساعات إضافي', 'ناقص انصراف', 'بدون صورة (غير محسوب)'], rows);
  rows.forEach((r, i) => {
    if (r[5] > 0) s1.getRange(i + 2, 6).setFontColor('#b91c1c').setFontWeight('bold');
    if (r[7] > 2) s1.getRange(i + 2, 8).setFontColor('#b91c1c').setFontWeight('bold');
    if (r[10] > 0) s1.getRange(i + 2, 11).setFontColor('#b91c1c').setFontWeight('bold');
    if (r[11] > 0) s1.getRange(i + 2, 12).setFontColor('#b91c1c').setFontWeight('bold');
  });
  write_(ss, 'تفصيل ' + mp + label,
    ['التاريخ', 'الاسم', 'الموقع', 'الحضور', 'الانصراف', 'الساعات', 'دقائق التأخير', 'ساعات إضافي', 'ملاحظة', 'صورة الحضور'], det);
  ss.toast('تم تجهيز تقفيل ' + mp + label);
}

// أقرب موقع للوكيشن اللي الصفحة بعتته (من غير ما نكشف الإحداثيات)
function nearest_(d) {
  const rows = sites_(ss_());
  let best = null;
  rows.forEach(r => {
    if (!r['الموقع'] || !r['خط العرض'] || !r['خط الطول']) return;
    const m = meters(+d.lat, +d.lng, +r['خط العرض'], +r['خط الطول']);
    if (!isNaN(m) && (!best || m < best.m)) best = { m: m, s: r['الموقع'], r: +r['النطاق بالمتر'] || 300 };
  });
  if (!best) return out({ ok: false, msg: 'إحداثيات المواقع لسه مش متسجلة' });
  return out({ ok: true, site: best.s, dist: Math.round(best.m), within: best.m <= best.r });
}

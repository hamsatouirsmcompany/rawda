/**
 * Hamsa Tourism - Apps Script backend
 * Script properties المطلوبة: ADMIN_EMAIL و ADMIN_PASS
 * Deploy > New deployment > Web app (Execute as: Me, Access: Anyone)
 * بعد أي تعديل: Deploy > Manage deployments > Edit > New version
 */
const OFFERS = 'Offers', USERS = 'Users', SUBS = 'Subscriptions';
const H_OFFERS = ['id', 'section', 'title', 'price', 'details', 'image', 'link', 'updated'];
const H_USERS = ['email', 'name', 'phone', 'salt', 'hash', 'created'];
const H_SUBS = ['time', 'offerId', 'offerTitle', 'name', 'email', 'phone', 'status'];

function sh_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); }
  return s;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hash_(salt, pass) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pass));
}
function rows_(s, head) {
  return s.getDataRange().getValues().slice(1).filter(r => r[0] !== '').map(r => {
    const o = {}; head.forEach((h, i) => o[h] = r[i]); return o;
  });
}
function newSession_(role, email, name) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('tok_' + token, JSON.stringify({ role: role, email: email, name: name }), 21600);
  return { ok: true, token: token, role: role, name: name, email: email };
}

function doGet() {
  return json_({ ok: true, offers: rows_(sh_(OFFERS, H_OFFERS), H_OFFERS) });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad' }); }
  const props = PropertiesService.getScriptProperties();
  const email = String(d.email || '').trim().toLowerCase();

  if (d.action === 'register') {
    if (!/^\S+@\S+\.\S+$/.test(email) || String(d.pass || '').length < 6 || !d.name || d.consent !== true)
      return json_({ ok: false, error: 'invalid' });
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      const s = sh_(USERS, H_USERS);
      if (email === String(props.getProperty('ADMIN_EMAIL')).toLowerCase() || rows_(s, H_USERS).some(u => u.email === email))
        return json_({ ok: false, error: 'exists' });
      const salt = Utilities.getUuid();
      s.appendRow([email, d.name, d.phone || '', salt, hash_(salt, d.pass), new Date()]);
      return json_(newSession_('user', email, d.name));
    } finally { lock.releaseLock(); }
  }

  if (d.action === 'login') {
    Utilities.sleep(800);
    const cache = CacheService.getScriptCache(), failKey = 'fail_' + email;
    const fails = Number(cache.get(failKey) || 0);
    if (fails >= 5) return json_({ ok: false, error: 'locked' }); // 5 محاولات خاطئة = قفل 15 دقيقة
    let session = null;
    if (email === String(props.getProperty('ADMIN_EMAIL')).toLowerCase()) {
      if (d.pass === props.getProperty('ADMIN_PASS')) session = newSession_('admin', email, 'المدير');
    } else {
      const u = rows_(sh_(USERS, H_USERS), H_USERS).find(x => x.email === email);
      if (u && u.hash === hash_(u.salt, d.pass)) session = newSession_('user', email, u.name);
    }
    if (session) { cache.remove(failKey); return json_(session); }
    cache.put(failKey, String(fails + 1), 900);
    return json_({ ok: false, error: 'credentials' });
  }

  const raw = d.token && CacheService.getScriptCache().get('tok_' + d.token);
  if (!raw) return json_({ ok: false, error: 'auth' });
  const me = JSON.parse(raw);

  if (d.action === 'subscribe') {
    const o = rows_(sh_(OFFERS, H_OFFERS), H_OFFERS).find(x => String(x.id) === String(d.offerId));
    if (!o) return json_({ ok: false, error: 'missing' });
    const u = rows_(sh_(USERS, H_USERS), H_USERS).find(x => x.email === me.email) || {};
    sh_(SUBS, H_SUBS).appendRow([new Date(), o.id, o.title, me.name, me.email, u.phone || '', 'جديد']);
    return json_({ ok: true });
  }

  if (me.role !== 'admin') return json_({ ok: false, error: 'forbidden' });

  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const s = sh_(OFFERS, H_OFFERS);
    const data = s.getDataRange().getValues();
    const find = id => data.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
    if (d.action === 'save') {
      const o = d.offer || {}, id = o.id || Utilities.getUuid();
      const row = [id, o.section, o.title, o.price, o.details, o.image, o.link, new Date()];
      const idx = o.id ? find(o.id) : -1;
      if (idx > 0) s.getRange(idx + 1, 1, 1, row.length).setValues([row]); else s.appendRow(row);
      return json_({ ok: true });
    }
    if (d.action === 'delete') { const i = find(d.id); if (i > 0) s.deleteRow(i + 1); return json_({ ok: true }); }
    if (d.action === 'subs') {
      const all = sh_(SUBS, H_SUBS).getDataRange().getValues();
      const subs = all.slice(1).map((r, i) => { const o = { row: i + 2 }; H_SUBS.forEach((h, j) => o[h] = r[j]); return o; })
        .filter(o => o.time !== '').reverse();
      return json_({ ok: true, subs: subs });
    }
    if (d.action === 'setStatus') {
      const ok = ['جديد', 'تم التواصل', 'تم الحجز', 'ملغي'];
      const ss = sh_(SUBS, H_SUBS), row = Number(d.row);
      if (ok.indexOf(d.status) < 0 || row < 2 || row > ss.getLastRow()) return json_({ ok: false, error: 'invalid' });
      if (String(ss.getRange(row, 5).getValue()).toLowerCase() !== String(d.email).toLowerCase()) return json_({ ok: false, error: 'mismatch' });
      ss.getRange(row, 7).setValue(d.status);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown' });
  } finally { lock.releaseLock(); }
}

/** شغّل هذه الدالة مرة واحدة من محرر Apps Script لتحميل العروض الأولية (لا تعمل لو فيه عروض بالفعل) */
function seedOffers() {
  const s = sh_(OFFERS, H_OFFERS);
  if (s.getLastRow() > 1) { Logger.log('فيه عروض بالفعل، لم يتم التحميل'); return; }
  const data = [
 {
  "id": "tayser-1",
  "section": "tayser",
  "title": "عمرة التيسير 1,000 جنيه شهريًا",
  "price": "1,000 جنيه / شهريًا",
  "details": "المدة: 24 شهر • الإجمالي: 24,000 جنيه\nاشتراك شهري بدون مقدم وبدون فوائد وبدون مصاريف إدارية.\nقرعة علنية كل 3 شهور لتحديد أدوار السفر، والسفر بسعر الموسم كاش.\nكل المشتركين يدخلون السحب على عمرة مجانية.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-2",
  "section": "tayser",
  "title": "عمرة التيسير 1,250 جنيه شهريًا",
  "price": "1,250 جنيه / شهريًا",
  "details": "المدة: 24 شهر • الإجمالي: 30,000 جنيه\nاشتراك شهري بدون مقدم وبدون فوائد وبدون مصاريف إدارية.\nقرعة علنية كل 3 شهور لتحديد أدوار السفر، والسفر بسعر الموسم كاش.\nكل المشتركين يدخلون السحب على عمرة مجانية.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-3",
  "section": "tayser",
  "title": "عمرة التيسير 1,500 جنيه شهريًا",
  "price": "1,500 جنيه / شهريًا",
  "details": "المدة: 24 شهر • الإجمالي: 36,000 جنيه\nاشتراك شهري بدون مقدم وبدون فوائد وبدون مصاريف إدارية.\nقرعة علنية كل 3 شهور لتحديد أدوار السفر، والسفر بسعر الموسم كاش.\nكل المشتركين يدخلون السحب على عمرة مجانية.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-4",
  "section": "tayser",
  "title": "عمرة التيسير 2,000 جنيه شهريًا",
  "price": "2,000 جنيه / شهريًا",
  "details": "المدة: 12 شهر • الإجمالي: 24,000 جنيه\nاشتراك شهري بدون مقدم وبدون فوائد وبدون مصاريف إدارية.\nقرعة علنية كل 3 شهور لتحديد أدوار السفر، والسفر بسعر الموسم كاش.\nكل المشتركين يدخلون السحب على عمرة مجانية.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-5",
  "section": "tayser",
  "title": "عمرة التيسير 2,500 جنيه شهريًا",
  "price": "2,500 جنيه / شهريًا",
  "details": "المدة: 12 شهر • الإجمالي: 30,000 جنيه\nاشتراك شهري بدون مقدم وبدون فوائد وبدون مصاريف إدارية.\nقرعة علنية كل 3 شهور لتحديد أدوار السفر، والسفر بسعر الموسم كاش.\nكل المشتركين يدخلون السحب على عمرة مجانية.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-6",
  "section": "tayser",
  "title": "عمرة التيسير 3,000 جنيه شهريًا",
  "price": "3,000 جنيه / شهريًا",
  "details": "المدة: 12 شهر • الإجمالي: 36,000 جنيه\nاشتراك شهري بدون مقدم وبدون فوائد وبدون مصاريف إدارية.\nقرعة علنية كل 3 شهور لتحديد أدوار السفر، والسفر بسعر الموسم كاش.\nكل المشتركين يدخلون السحب على عمرة مجانية.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-7",
  "section": "tayser",
  "title": "عمرة التيسير الذهبية",
  "price": "4,000 جنيه / شهريًا",
  "details": "المدة: 12 شهر • الإجمالي: 48,000 جنيه\nاشتراك شهري بدون مقدم وبدون فوائد وبدون مصاريف إدارية.\nقرعة علنية كل 3 شهور لتحديد أدوار السفر، والسفر بسعر الموسم كاش.\nكل المشتركين يدخلون السحب على عمرة مجانية.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-8",
  "section": "tayser",
  "title": "عمرة التيسير البلاتينية",
  "price": "5,000 جنيه / شهريًا",
  "details": "المدة: 12 شهر • الإجمالي: 60,000 جنيه\nاشتراك شهري بدون مقدم وبدون فوائد وبدون مصاريف إدارية.\nقرعة علنية كل 3 شهور لتحديد أدوار السفر، والسفر بسعر الموسم كاش.\nكل المشتركين يدخلون السحب على عمرة مجانية.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-9",
  "section": "tayser",
  "title": "تيسير شعبان ورمضان",
  "price": "3,500 جنيه / شهريًا",
  "details": "المدة: 12 شهر • الإجمالي: 42,000 جنيه\nاشتراك شهري بدون مقدم وبدون فوائد وبدون مصاريف إدارية.\nقرعة علنية كل 3 شهور لتحديد أدوار السفر، والسفر بسعر الموسم كاش.\nكل المشتركين يدخلون السحب على عمرة مجانية.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-10",
  "section": "tayser",
  "title": "تثبيت سعر رمضان",
  "price": "45,000 جنيه",
  "details": "سعر ثابت بدون تذاكر الطيران.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-11",
  "section": "tayser",
  "title": "تيسير رمضان — 16 شهر",
  "price": "3,000 جنيه / شهريًا",
  "details": "المدة: 16 شهر، يبدأ من شهر 6.\nحجز مرن: يمكن الحجز بعد بداية الحجز بدفع الشهور السابقة بأثر رجعي.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "tayser-12",
  "section": "tayser",
  "title": "تيسير رمضان — 20 شهر",
  "price": "2,500 جنيه / شهريًا",
  "details": "المدة: 20 شهر، يبدأ من شهر 5.\nحجز مرن: يمكن الحجز بعد بداية الحجز بدفع الشهور السابقة بأثر رجعي.",
  "image": "",
  "link": "https://hamsatouirsmcompany.github.io/TAYSER/"
 },
 {
  "id": "umrah-1",
  "section": "umrah",
  "title": "العمرة بالقسط",
  "price": "ادفع 50% أو 60%",
  "details": "اختر البرنامج وادفع 50% أو 60% وقسّط الباقي من 3 إلى 12 شهر بدون تعقيد.",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/hamsa-torism/"
 },
 {
  "id": "umrah-2",
  "section": "umrah",
  "title": "عمرة كاش — ريع بخش (اقتصادي بالمواصلات)",
  "price": "من 40,500 جنيه (رباعي)",
  "details": "المدينة: نسك المدينة / درة الإيمان / طيبة هيلز — 3 ليالي\nمكة: قصر العليان / النخبة 1 / أبراج القصواء — 11 ليلة\nرباعي 40,500 • ثلاثي 43,500 • ثنائي 48,900 • سينجل 64,500 • طفل 31,500 • رضيع 17,500",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/hamsa-torism/"
 },
 {
  "id": "umrah-3",
  "section": "umrah",
  "title": "عمرة كاش — اقتصادي مشي (بير بليلة)",
  "price": "من 42,000 جنيه (رباعي)",
  "details": "المدينة: نسك المدينة / درة الإيمان / طيبة هيلز — 3 ليالي\nمكة: إعمار أفاق — 11 ليلة\nرباعي 42,000 • ثلاثي 46,500 • ثنائي 52,900 • سينجل 72,900 • طفل 31,500 • رضيع 17,500",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/hamsa-torism/"
 },
 {
  "id": "umrah-4",
  "section": "umrah",
  "title": "عمرة كاش — اقتصادي مشي (أجياد السد)",
  "price": "من 42,500 جنيه (رباعي)",
  "details": "المدينة: نسك المدينة / درة الإيمان / طيبة هيلز — 3 ليالي\nمكة: واحة الضيافة / العليان أجياد — 11 ليلة\nرباعي 42,500 • ثلاثي 46,900 • ثنائي 53,500 • سينجل 73,900 • طفل 31,500 • رضيع 17,500",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/hamsa-torism/"
 },
 {
  "id": "umrah-5",
  "section": "umrah",
  "title": "عمرة كاش 5 نجوم (أ)",
  "price": "من 66,900 جنيه (رباعي)",
  "details": "3 ليالي المدينة + 4 ليالي مكة بالإفطار\nالطيران: سعودي — القاهرة / المدينة / جدة / القاهرة\nالمدينة: فندق الحرم — مكة: الصفوة البرج الثالث (أول مطل على الحرم)\nرباعي 66,900 • ثلاثي 70,900 • ثنائي 79,900",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/hamsa-torism/"
 },
 {
  "id": "umrah-6",
  "section": "umrah",
  "title": "عمرة كاش 5 نجوم (ب)",
  "price": "من 55,500 جنيه (رباعي)",
  "details": "3 ليالي المدينة + 4 ليالي مكة بالإفطار\nالطيران: سعودي — القاهرة / المدينة / جدة / القاهرة\nالمدينة: درة الإيمان — مكة: الشهداء (أول مطل على الحرم)\nرباعي 55,500 • ثلاثي 58,900 • ثنائي 64,500",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/hamsa-torism/"
 },
 {
  "id": "umrah-7",
  "section": "umrah",
  "title": "عمرة كاش — اقتصادي مميز",
  "price": "من 47,000 جنيه (رباعي)",
  "details": "4 ليالي المدينة + 5 ليالي مكة\nالطيران: سعودي — القاهرة / المدينة / جدة / القاهرة\nالمدينة: كونكورد دار الخير — مكة: الماسة جراند (400 متر من الحرم)\nرباعي 47,000 • ثلاثي 50,900 • ثنائي 57,900",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/hamsa-torism/"
 },
 {
  "id": "hajj-1",
  "section": "hajj",
  "title": "الحج السياحي",
  "price": "395,000 – 650,000 جنيه",
  "details": "عدة مستويات وبرامج بإقامة مختلفة وقرب متفاوت من الحرم.\nخدمات مميزة للحجاج مع إشراف ومتابعة.",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/alhaju/"
 },
 {
  "id": "hajj-2",
  "section": "hajj",
  "title": "الحج الميسر",
  "price": "",
  "details": "برنامج حج ميسر بخدمات متكاملة وأسعار وخيارات متعددة. استفسر عن التفاصيل والأسعار.",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/alhaju/"
 },
 {
  "id": "hajj-3",
  "section": "hajj",
  "title": "الحج الميسر بالتقسيط",
  "price": "تقسيط حتى 120,000 جنيه",
  "details": "قسّط جزءًا من سعر أي برنامج حج على 24 شهر بواقع 5,000 جنيه شهريًا، وادفع باقي سعر البرنامج مقدمًا.",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/alhaju/"
 },
 {
  "id": "hajj-4",
  "section": "hajj",
  "title": "حج القرعة",
  "price": "",
  "details": "سجّل اهتمامك، واختر البرنامج المناسب وسنتواصل معك لاستكمال الإجراءات.",
  "image": "",
  "link": "https://amrkhaledmohamedeng-create.github.io/alhaju/"
 }
];
  const now = new Date();
  s.getRange(2, 1, data.length, H_OFFERS.length).setValues(data.map(o => [o.id, o.section, o.title, o.price, o.details, o.image, o.link, now]));
  Logger.log('تم تحميل ' + data.length + ' عرض');
}

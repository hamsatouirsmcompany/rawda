/**
 * نظام إدارة عملاء وحجوزات الروضة الشريفة
 * Backend: Google Apps Script + Google Sheets (كقاعدة بيانات)
 *
 * طريقة النشر:
 * 1) افتح script.google.com > مشروع جديد، الصق هذا الكود بدل Code.gs
 * 2) من القائمة الجانبية أضف ملف جديد باسم appsscript.json إذا احتجت (اختياري)
 * 3) شغّل الدالة setupSheets() مرة واحدة يدويًا من المحرر لإنشاء الشيتات
 *    (سيطلب منك صلاحيات - اقبلها). سينشئ لك مستخدم أدمن افتراضي:
 *    البريد: admin@system.local  |  كلمة المرور: Admin@123  (غيّرها فورًا)
 * 4) Deploy > New deployment > Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5) انسخ الرابط (Web app URL) وضعه في ملف index.html في المتغير API_URL
 */

const SHEET_NAMES = {
  USERS: 'Users',
  CUSTOMERS: 'Customers',
  APPOINTMENTS: 'Rawdah_Appointments',
  HISTORY: 'Booking_History',
  LOGS: 'Activity_Logs',
  NOTIFS: 'Notifications'
};

const HEADERS = {
  Users: ['id', 'name', 'email', 'passwordHash', 'role', 'status', 'assignedScope', 'createdAt', 'updatedAt'],
  Customers: ['id', 'fullName', 'passportNumber', 'nationality', 'phone', 'countryCode', 'personsCount',
    'arrivalDate', 'departureDate', 'arrivalCity', 'departureCity', 'madinahFrom', 'madinahTo',
    'preferredFrom', 'preferredTo', 'notes', 'assignedEmployee', 'dataStatus', 'bookingStatus',
    'priority', 'archived', 'createdAt', 'updatedAt'],
  Rawdah_Appointments: ['id', 'appointmentDate', 'appointmentTime', 'status', 'customerId',
    'bookingReference', 'source', 'bookedBy', 'bookedAt', 'notes', 'createdAt', 'updatedAt'],
  Booking_History: ['id', 'customerId', 'appointmentId', 'oldStatus', 'newStatus', 'oldAppointment',
    'newAppointment', 'reason', 'changedBy', 'createdAt'],
  Activity_Logs: ['id', 'userId', 'userName', 'action', 'entityType', 'entityId', 'description', 'createdAt'],
  Notifications: ['id', 'userId', 'title', 'message', 'type', 'isRead', 'createdAt']
};

/* ============================ ENTRY POINTS ============================ */

function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let body = {};
  try {
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.payload) {
      body = JSON.parse(e.parameter.payload);
    }
  } catch (err) {
    return respond({ ok: false, error: 'JSON غير صالح' });
  }

  const action = body.action || (e && e.parameter && e.parameter.action);
  let session = null;
  if (action !== 'login' && action !== 'setup') {
    session = validateToken(body.token);
    if (!session) return respond({ ok: false, error: 'الجلسة غير صالحة، الرجاء تسجيل الدخول مجددًا' });
  }

  try {
    switch (action) {
      case 'login': return respond(actionLogin(body));
      case 'logout': return respond(actionLogout(body, session));

      case 'listUsers': return respond(guardAdmin(session, () => actionListUsers()));
      case 'addUser': return respond(guardAdmin(session, () => actionAddUser(body, session)));
      case 'updateUser': return respond(guardAdmin(session, () => actionUpdateUser(body, session)));
      case 'toggleUserStatus': return respond(guardAdmin(session, () => actionToggleUserStatus(body, session)));

      case 'listCustomers': return respond(actionListCustomers(body, session));
      case 'addCustomer': return respond(actionAddCustomer(body, session));
      case 'updateCustomer': return respond(actionUpdateCustomer(body, session));
      case 'archiveCustomer': return respond(actionArchiveCustomer(body, session));
      case 'getCustomer': return respond(actionGetCustomer(body));
      case 'changeBookingStatus': return respond(actionChangeBookingStatus(body, session));

      case 'listAppointments': return respond(actionListAppointments(body));
      case 'addAppointmentsBulk': return respond(actionAddAppointmentsBulk(body, session));
      case 'updateAppointment': return respond(actionUpdateAppointment(body, session));
      case 'deleteAppointment': return respond(actionDeleteAppointment(body, session));
      case 'searchAvailable': return respond(actionSearchAvailable(body));
      case 'bookAppointment': return respond(actionBookAppointment(body, session));

      case 'getTimeline': return respond(actionGetTimeline(body));
      case 'listActivityLogs': return respond(guardAdmin(session, () => actionListActivityLogs(body)));
      case 'listNotifications': return respond(actionListNotifications(session));
      case 'markNotificationRead': return respond(actionMarkNotificationRead(body));

      case 'dashboardStats': return respond(actionDashboardStats(session));
      case 'dailyFollowUp': return respond(actionDailyFollowUp(session));
      case 'reportsSummary': return respond(actionReportsSummary(body));

      case 'setup': return respond(setupSheets());

      default: return respond({ ok: false, error: 'إجراء غير معروف: ' + action });
    }
  } catch (err) {
    return respond({ ok: false, error: 'خطأ في الخادم: ' + err.message });
  }
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ SETUP ============================ */

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
    }
  });
  // remove default "Sheet1" if empty and unused
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  // seed default admin if Users empty
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (usersSheet.getLastRow() < 2) {
    const id = Utilities.getUuid();
    usersSheet.appendRow([id, 'المدير العام', 'admin@system.local', hashPassword('Admin@123'),
      'admin', 'active', '', nowStr(), nowStr()]);
  }
  return { ok: true, message: 'تم إنشاء الشيتات بنجاح. البريد: admin@system.local — كلمة المرور: Admin@123' };
}

/* ============================ HELPERS ============================ */

function nowStr() { return new Date().toISOString(); }

function hashPassword(pw) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw + '::rawdah_salt::');
  return digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('الشيت غير موجود: ' + name + ' — نفّذ setupSheets() أولًا');
  return sh;
}

function sheetToObjects(sheetName) {
  const sh = getSheet(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map((row, idx) => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    obj._row = idx + 2; // actual sheet row number
    return obj;
  });
}

function appendObject(sheetName, obj) {
  const sh = getSheet(sheetName);
  const headers = HEADERS[sheetName];
  const row = headers.map(h => (obj[h] !== undefined ? obj[h] : ''));
  sh.appendRow(row);
  return obj;
}

function updateObjectById(sheetName, id, patch) {
  const sh = getSheet(sheetName);
  const headers = HEADERS[sheetName];
  const idCol = headers.indexOf('id') + 1;
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol - 1] === id) {
      headers.forEach((h, i) => {
        if (patch[h] !== undefined) sh.getRange(r + 1, i + 1).setValue(patch[h]);
      });
      const updated = {};
      headers.forEach((h, i) => updated[h] = patch[h] !== undefined ? patch[h] : data[r][i]);
      return updated;
    }
  }
  return null;
}

function findById(sheetName, id) {
  const rows = sheetToObjects(sheetName);
  return rows.find(r => r.id === id) || null;
}

function logActivity(session, action, entityType, entityId, description) {
  appendObject(SHEET_NAMES.LOGS, {
    id: Utilities.getUuid(),
    userId: session ? session.userId : 'system',
    userName: session ? session.userName : 'النظام',
    action, entityType, entityId, description,
    createdAt: nowStr()
  });
}

function addNotification(userId, title, message, type) {
  appendObject(SHEET_NAMES.NOTIFS, {
    id: Utilities.getUuid(), userId, title, message, type,
    isRead: false, createdAt: nowStr()
  });
}

/* ============================ SESSIONS (CacheService, 6h) ============================ */

function validateToken(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('session_' + token);
  if (!raw) return null;
  return JSON.parse(raw);
}

function guardAdmin(session, fn) {
  if (!session || session.role !== 'admin') return { ok: false, error: 'هذا الإجراء متاح للمدير فقط' };
  return fn();
}

/* ============================ AUTH ============================ */

function actionLogin(body) {
  const users = sheetToObjects(SHEET_NAMES.USERS);
  const user = users.find(u => u.email === body.email);
  if (!user) return { ok: false, error: 'بيانات الدخول غير صحيحة' };
  if (user.status !== 'active') return { ok: false, error: 'هذا الحساب معطّل' };
  if (user.passwordHash !== hashPassword(body.password || '')) return { ok: false, error: 'بيانات الدخول غير صحيحة' };

  const token = Utilities.getUuid();
  const session = { token, userId: user.id, userName: user.name, role: user.role };
  CacheService.getScriptCache().put('session_' + token, JSON.stringify(session), 6 * 60 * 60);
  logActivity(session, 'Login', 'User', user.id, 'تسجيل دخول: ' + user.name);
  return { ok: true, token, user: { id: user.id, name: user.name, email: user.email, role: user.role } };
}

function actionLogout(body, session) {
  if (session) {
    logActivity(session, 'Logout', 'User', session.userId, 'تسجيل خروج: ' + session.userName);
    CacheService.getScriptCache().remove('session_' + body.token);
  }
  return { ok: true };
}

/* ============================ USERS ============================ */

function actionListUsers() {
  const users = sheetToObjects(SHEET_NAMES.USERS).map(u => {
    delete u.passwordHash; delete u._row; return u;
  });
  return { ok: true, users };
}

function actionAddUser(body, session) {
  const id = Utilities.getUuid();
  const user = {
    id, name: body.name, email: body.email, passwordHash: hashPassword(body.password || 'Change@123'),
    role: body.role || 'staff', status: 'active', assignedScope: body.assignedScope || '',
    createdAt: nowStr(), updatedAt: nowStr()
  };
  appendObject(SHEET_NAMES.USERS, user);
  logActivity(session, 'إضافة موظف', 'User', id, 'تمت إضافة الموظف: ' + body.name);
  delete user.passwordHash;
  return { ok: true, user };
}

function actionUpdateUser(body, session) {
  const patch = { name: body.name, email: body.email, role: body.role, assignedScope: body.assignedScope, updatedAt: nowStr() };
  if (body.password) patch.passwordHash = hashPassword(body.password);
  const updated = updateObjectById(SHEET_NAMES.USERS, body.id, patch);
  logActivity(session, 'تعديل موظف', 'User', body.id, 'تم تعديل بيانات الموظف');
  return { ok: !!updated };
}

function actionToggleUserStatus(body, session) {
  const user = findById(SHEET_NAMES.USERS, body.id);
  if (!user) return { ok: false, error: 'الموظف غير موجود' };
  const newStatus = user.status === 'active' ? 'disabled' : 'active';
  updateObjectById(SHEET_NAMES.USERS, body.id, { status: newStatus, updatedAt: nowStr() });
  logActivity(session, 'تغيير حالة موظف', 'User', body.id, 'الحالة الجديدة: ' + newStatus);
  return { ok: true, status: newStatus };
}

/* ============================ CUSTOMERS ============================ */

function actionListCustomers(body, session) {
  let customers = sheetToObjects(SHEET_NAMES.CUSTOMERS);
  customers = customers.filter(c => String(c.archived) !== 'true');

  const f = body.filters || {};
  if (f.search) {
    const q = f.search.toString().trim();
    customers = customers.filter(c =>
      (c.fullName || '').toString().includes(q) ||
      (c.passportNumber || '').toString().includes(q) ||
      (c.phone || '').toString().includes(q)
    );
  }
  if (f.bookingStatus) customers = customers.filter(c => c.bookingStatus === f.bookingStatus);
  if (f.assignedEmployee) customers = customers.filter(c => c.assignedEmployee === f.assignedEmployee);
  if (f.dataStatus) customers = customers.filter(c => c.dataStatus === f.dataStatus);
  if (f.priority) customers = customers.filter(c => c.priority === f.priority);

  customers = customers.map(c => { c.priority = computePriority(c); return c; });
  customers.sort((a, b) => new Date(a.departureDate || '9999') - new Date(b.departureDate || '9999'));
  customers.forEach(c => delete c._row);
  return { ok: true, customers };
}

function computePriority(c) {
  if (c.priority && c.priorityManual === true) return c.priority;
  if (!c.departureDate) return '🟢 منخفض';
  const days = Math.ceil((new Date(c.departureDate) - new Date()) / (1000 * 60 * 60 * 24));
  if (days <= 2) return '🔴 عاجل';
  if (days <= 5) return '🟠 مهم';
  if (days <= 10) return '🟡 عادي';
  return '🟢 منخفض';
}

function actionAddCustomer(body, session) {
  const id = Utilities.getUuid();
  const c = Object.assign({}, body.customer, {
    id, dataStatus: body.customer.dataStatus || 'INCOMPLETE',
    bookingStatus: 'NEW', archived: false,
    assignedEmployee: body.customer.assignedEmployee || session.userName,
    createdAt: nowStr(), updatedAt: nowStr()
  });
  appendObject(SHEET_NAMES.CUSTOMERS, c);
  logActivity(session, 'إضافة عميل', 'Customer', id, 'تمت إضافة العميل: ' + c.fullName);
  addTimelineEntry(id, 'تم إنشاء العميل', session.userName);
  return { ok: true, customer: c };
}

function actionUpdateCustomer(body, session) {
  const patch = Object.assign({}, body.customer, { updatedAt: nowStr() });
  const updated = updateObjectById(SHEET_NAMES.CUSTOMERS, body.id, patch);
  logActivity(session, 'تعديل عميل', 'Customer', body.id, 'تم تعديل بيانات العميل');
  addTimelineEntry(body.id, 'تم تعديل بيانات العميل', session.userName);
  return { ok: !!updated };
}

function actionArchiveCustomer(body, session) {
  updateObjectById(SHEET_NAMES.CUSTOMERS, body.id, { archived: true, updatedAt: nowStr() });
  logActivity(session, 'أرشفة عميل', 'Customer', body.id, 'تمت أرشفة/حذف العميل');
  return { ok: true };
}

function actionGetCustomer(body) {
  const c = findById(SHEET_NAMES.CUSTOMERS, body.id);
  if (!c) return { ok: false, error: 'العميل غير موجود' };
  delete c._row;
  return { ok: true, customer: c };
}

const BOOKING_STATUSES = ['NEW', 'READY_TO_BOOK', 'SEARCHING', 'WAITING_AVAILABILITY', 'AVAILABLE',
  'BOOKING_IN_PROGRESS', 'CONFIRMED', 'TICKETED', 'MODIFICATION_REQUIRED', 'RESCHEDULED', 'CANCELLED',
  'NO_SHOW', 'REBOOK_REQUIRED', 'COMPLETED'];

function actionChangeBookingStatus(body, session) {
  const c = findById(SHEET_NAMES.CUSTOMERS, body.id);
  if (!c) return { ok: false, error: 'العميل غير موجود' };
  const oldStatus = c.bookingStatus;
  updateObjectById(SHEET_NAMES.CUSTOMERS, body.id, { bookingStatus: body.newStatus, updatedAt: nowStr() });
  appendObject(SHEET_NAMES.HISTORY, {
    id: Utilities.getUuid(), customerId: body.id, appointmentId: body.appointmentId || '',
    oldStatus, newStatus: body.newStatus, oldAppointment: '', newAppointment: '',
    reason: body.reason || '', changedBy: session.userName, createdAt: nowStr()
  });
  logActivity(session, 'تغيير حالة الحجز', 'Customer', body.id, oldStatus + ' → ' + body.newStatus);
  addTimelineEntry(body.id, 'تم تغيير الحالة إلى ' + body.newStatus, session.userName);
  return { ok: true };
}

/* ============================ APPOINTMENTS ============================ */

function actionListAppointments(body) {
  let appts = sheetToObjects(SHEET_NAMES.APPOINTMENTS);
  const f = body.filters || {};
  if (f.status) appts = appts.filter(a => a.status === f.status);
  if (f.dateFrom) appts = appts.filter(a => a.appointmentDate >= f.dateFrom);
  if (f.dateTo) appts = appts.filter(a => a.appointmentDate <= f.dateTo);
  appts.sort((a, b) => (a.appointmentDate + a.appointmentTime).localeCompare(b.appointmentDate + b.appointmentTime));
  appts.forEach(a => delete a._row);
  return { ok: true, appointments: appts };
}

function actionAddAppointmentsBulk(body, session) {
  const list = body.appointments || [];
  const created = [];
  list.forEach(item => {
    const id = Utilities.getUuid();
    const appt = {
      id, appointmentDate: item.date, appointmentTime: item.time, status: 'AVAILABLE',
      customerId: '', bookingReference: '', source: item.source || 'يدوي', bookedBy: '', bookedAt: '',
      notes: item.notes || '', createdAt: nowStr(), updatedAt: nowStr()
    };
    appendObject(SHEET_NAMES.APPOINTMENTS, appt);
    created.push(appt);
  });
  logActivity(session, 'إضافة مواعيد', 'Appointment', '-', 'تمت إضافة ' + created.length + ' موعد');
  const matches = matchNewAppointmentsToWaitingCustomers(created);
  return { ok: true, created: created.length, matches };
}

function actionUpdateAppointment(body, session) {
  const updated = updateObjectById(SHEET_NAMES.APPOINTMENTS, body.id, Object.assign({}, body.patch, { updatedAt: nowStr() }));
  logActivity(session, 'تعديل موعد', 'Appointment', body.id, 'تم تعديل الموعد');
  return { ok: !!updated };
}

function actionDeleteAppointment(body, session) {
  const sh = getSheet(SHEET_NAMES.APPOINTMENTS);
  const rows = sheetToObjects(SHEET_NAMES.APPOINTMENTS);
  const target = rows.find(r => r.id === body.id);
  if (!target) return { ok: false, error: 'الموعد غير موجود' };
  sh.deleteRow(target._row);
  logActivity(session, 'حذف موعد', 'Appointment', body.id, 'تم حذف الموعد');
  return { ok: true };
}

function actionSearchAvailable(body) {
  const c = findById(SHEET_NAMES.CUSTOMERS, body.customerId);
  if (!c) return { ok: false, error: 'العميل غير موجود' };
  const from = c.preferredFrom || c.arrivalDate;
  const to = c.preferredTo || c.departureDate;
  let appts = sheetToObjects(SHEET_NAMES.APPOINTMENTS).filter(a => a.status === 'AVAILABLE');
  appts = appts.filter(a => (!from || a.appointmentDate >= from) && (!to || a.appointmentDate <= to));
  appts.sort((a, b) => (a.appointmentDate + a.appointmentTime).localeCompare(b.appointmentDate + b.appointmentTime));
  appts.forEach(a => delete a._row);
  return { ok: true, appointments: appts };
}

function actionBookAppointment(body, session) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const appt = findById(SHEET_NAMES.APPOINTMENTS, body.appointmentId);
    if (!appt) return { ok: false, error: 'الموعد غير موجود' };
    if (appt.status !== 'AVAILABLE') {
      return { ok: false, error: '⚠️ هذا الموعد تم حجزه بالفعل.' };
    }
    const c = findById(SHEET_NAMES.CUSTOMERS, body.customerId);
    if (!c) return { ok: false, error: 'العميل غير موجود' };

    updateObjectById(SHEET_NAMES.APPOINTMENTS, body.appointmentId, {
      status: 'BOOKED', customerId: body.customerId, bookedBy: session.userName,
      bookedAt: nowStr(), bookingReference: body.bookingReference || '', updatedAt: nowStr()
    });
    updateObjectById(SHEET_NAMES.CUSTOMERS, body.customerId, { bookingStatus: 'CONFIRMED', updatedAt: nowStr() });
    appendObject(SHEET_NAMES.HISTORY, {
      id: Utilities.getUuid(), customerId: body.customerId, appointmentId: body.appointmentId,
      oldStatus: c.bookingStatus, newStatus: 'CONFIRMED', oldAppointment: '',
      newAppointment: appt.appointmentDate + ' ' + appt.appointmentTime,
      reason: 'حجز موعد', changedBy: session.userName, createdAt: nowStr()
    });
    logActivity(session, 'حجز موعد', 'Customer', body.customerId,
      'تم حجز موعد ' + appt.appointmentDate + ' ' + appt.appointmentTime);
    addTimelineEntry(body.customerId, 'تم تأكيد الحجز: ' + appt.appointmentDate + ' ' + appt.appointmentTime, session.userName);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function matchNewAppointmentsToWaitingCustomers(newAppts) {
  const waiting = sheetToObjects(SHEET_NAMES.CUSTOMERS)
    .filter(c => c.bookingStatus === 'WAITING_AVAILABILITY' || c.bookingStatus === 'READY_TO_BOOK');
  const matchedIds = new Set();
  newAppts.forEach(a => {
    waiting.forEach(c => {
      const from = c.preferredFrom || c.arrivalDate, to = c.preferredTo || c.departureDate;
      if ((!from || a.appointmentDate >= from) && (!to || a.appointmentDate <= to)) {
        matchedIds.add(c.id);
      }
    });
  });
  matchedIds.forEach(id => {
    const c = findById(SHEET_NAMES.CUSTOMERS, id);
    updateObjectById(SHEET_NAMES.CUSTOMERS, id, { bookingStatus: 'AVAILABLE', updatedAt: nowStr() });
    addTimelineEntry(id, 'تم العثور على موعد مناسب', 'النظام');
  });
  if (matchedIds.size > 0) {
    addNotification('all', 'مواعيد مناسبة', 'يوجد موعد مناسب لعدد ' + matchedIds.size + ' من العملاء.', 'match');
  }
  return matchedIds.size;
}

/* ============================ TIMELINE ============================ */

function addTimelineEntry(customerId, text, byWhom) {
  appendObject(SHEET_NAMES.HISTORY, {
    id: Utilities.getUuid(), customerId, appointmentId: '', oldStatus: '', newStatus: '',
    oldAppointment: '', newAppointment: '', reason: text, changedBy: byWhom, createdAt: nowStr()
  });
}

function actionGetTimeline(body) {
  let history = sheetToObjects(SHEET_NAMES.HISTORY).filter(h => h.customerId === body.customerId);
  history.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  history.forEach(h => delete h._row);
  return { ok: true, history };
}

/* ============================ LOGS / NOTIFICATIONS ============================ */

function actionListActivityLogs(body) {
  let logs = sheetToObjects(SHEET_NAMES.LOGS);
  logs.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  logs = logs.slice(0, 300);
  logs.forEach(l => delete l._row);
  return { ok: true, logs };
}

function actionListNotifications(session) {
  let notifs = sheetToObjects(SHEET_NAMES.NOTIFS)
    .filter(n => n.userId === 'all' || n.userId === session.userId);
  notifs.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  notifs = notifs.slice(0, 50);
  notifs.forEach(n => delete n._row);
  return { ok: true, notifications: notifs };
}

function actionMarkNotificationRead(body) {
  updateObjectById(SHEET_NAMES.NOTIFS, body.id, { isRead: true });
  return { ok: true };
}

/* ============================ DASHBOARD / REPORTS ============================ */

function actionDashboardStats(session) {
  const customers = sheetToObjects(SHEET_NAMES.CUSTOMERS).filter(c => String(c.archived) !== 'true');
  const counts = {};
  BOOKING_STATUSES.forEach(s => counts[s] = 0);
  customers.forEach(c => { counts[c.bookingStatus] = (counts[c.bookingStatus] || 0) + 1; });

  const alerts = [];
  const availableNoBooking = customers.filter(c => c.bookingStatus === 'AVAILABLE').length;
  if (availableNoBooking > 0) alerts.push('يوجد ' + availableNoBooking + ' عملاء لديهم موعد مناسب ولم يتم الحجز لهم.');
  const rebook = counts['REBOOK_REQUIRED'] || 0;
  if (rebook > 0) alerts.push('يوجد ' + rebook + ' عملاء يحتاجون إعادة حجز.');
  const soon = customers.filter(c => {
    if (!c.departureDate) return false;
    const hrs = (new Date(c.departureDate) - new Date()) / (1000 * 60 * 60);
    return hrs > 0 && hrs <= 48 && !['CONFIRMED', 'TICKETED', 'COMPLETED'].includes(c.bookingStatus);
  }).length;
  if (soon > 0) alerts.push('يوجد ' + soon + ' عملاء ستنتهي فترة السفر الخاصة بهم خلال 48 ساعة.');
  const modReq = counts['MODIFICATION_REQUIRED'] || 0;
  if (modReq > 0) alerts.push('يوجد ' + modReq + ' حجوزات تحتاج إلى مراجعة.');

  return { ok: true, total: customers.length, counts, alerts };
}

function actionDailyFollowUp(session) {
  const customers = sheetToObjects(SHEET_NAMES.CUSTOMERS).filter(c => String(c.archived) !== 'true');
  const result = {
    searching: customers.filter(c => c.bookingStatus === 'SEARCHING' || c.bookingStatus === 'WAITING_AVAILABILITY'),
    available: customers.filter(c => c.bookingStatus === 'AVAILABLE'),
    upcoming: customers.filter(c => {
      if (!c.departureDate) return false;
      const days = (new Date(c.departureDate) - new Date()) / (1000 * 60 * 60 * 24);
      return days >= 0 && days <= 3;
    }),
    rebook: customers.filter(c => c.bookingStatus === 'REBOOK_REQUIRED'),
    updateNeeded: customers.filter(c => c.dataStatus === 'UPDATE_REQUIRED' || c.dataStatus === 'INVALID_DATA')
  };
  Object.keys(result).forEach(k => result[k].forEach(c => delete c._row));
  return { ok: true, followUp: result };
}

function actionReportsSummary(body) {
  const customers = sheetToObjects(SHEET_NAMES.CUSTOMERS);
  const history = sheetToObjects(SHEET_NAMES.HISTORY);
  const range = body.range || {};
  const inRange = (d) => {
    if (!d) return false;
    const t = new Date(d).getTime();
    if (range.from && t < new Date(range.from).getTime()) return false;
    if (range.to && t > new Date(range.to).getTime()) return false;
    return true;
  };
  const filteredCustomers = range.from || range.to ? customers.filter(c => inRange(c.createdAt)) : customers;

  const byEmployee = {};
  filteredCustomers.forEach(c => {
    const emp = c.assignedEmployee || 'غير محدد';
    byEmployee[emp] = (byEmployee[emp] || 0) + 1;
  });

  const cancellations = history.filter(h => h.newStatus === 'CANCELLED' && (!range.from && !range.to || inRange(h.createdAt))).length;
  const rebooks = history.filter(h => h.reason === 'حجز موعد' && (!range.from && !range.to || inRange(h.createdAt))).length;
  const confirmed = filteredCustomers.filter(c => ['CONFIRMED', 'TICKETED', 'COMPLETED'].includes(c.bookingStatus)).length;
  const cancelledCustomers = filteredCustomers.filter(c => c.bookingStatus === 'CANCELLED').length;
  const notBooked = filteredCustomers.filter(c => ['NEW', 'READY_TO_BOOK', 'SEARCHING', 'WAITING_AVAILABILITY'].includes(c.bookingStatus)).length;
  const waiting = filteredCustomers.filter(c => c.bookingStatus === 'WAITING_AVAILABILITY').length;
  const completed = filteredCustomers.filter(c => c.bookingStatus === 'COMPLETED').length;

  return {
    ok: true,
    report: {
      totalCustomers: filteredCustomers.length,
      totalBookings: confirmed,
      confirmed, cancelled: cancelledCustomers, notBooked, waiting, completed,
      cancellations, rebooks,
      completionRate: filteredCustomers.length ? Math.round((confirmed / filteredCustomers.length) * 100) : 0,
      byEmployee
    }
  };
}

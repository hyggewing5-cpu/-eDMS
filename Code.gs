/**
 * ============================================================
 *  eDMS — Electronic Document Management System
 *  Google Sheets + Apps Script (Bound script) + Google Drive
 *  Version 1.0.0
 * ------------------------------------------------------------
 *  โครงสร้างเป็นโมดูล:
 *   1) CONFIG / SCHEMA     — โครงสร้างชีต
 *   2) DB helpers          — อ่าน/เขียนชีต (ไม่ผูกกับ eDMS โดยตรง นำไปใช้ซ้ำได้)
 *   3) AUTH module         — ล็อกอิน / session
 *   4) ACCESS module       — สิทธิ์ตามแผนก + สิทธิ์รายบุคคล
 *   5) DOCS / FILES module — เอกสาร ไฟล์ ลิงก์ ดาวน์โหลด ประวัติ
 *   6) ADMIN module        — ผู้ใช้ แผนก หมวดหมู่ ตั้งค่า
 *   7) ROUTER (api)        — จุดเรียกเดียวจาก Index.html
 *  ทุกฟังก์ชันเรียกผ่าน api(method, token, payload) เท่านั้น
 * ============================================================
 */

/* ---------- 1) CONFIG / SCHEMA ---------- */
const TZ = 'Asia/Bangkok';
const SESSION_SEC = 21600; // 6 ชั่วโมง (สูงสุดของ CacheService)
const SCHEMA = {
  Config:       ['Key', 'Value'],
  Departments:  ['DeptID', 'NameTH', 'NameEN', 'Active'],
  Categories:   ['CatID', 'NameTH', 'NameEN', 'Active', 'Direction', 'Prefix'],   // = ประเภทเอกสาร (Direction: in|internal|out|other)
  Purposes:     ['PurposeID', 'NameTH', 'NameEN', 'Active'],
  Users:        ['UserID', 'Username', 'PasswordHash', 'Salt', 'FullName', 'DeptID', 'Role', 'CanUpload', 'CanDownload', 'Active', 'Lang', 'CreatedAt', 'LastLogin'],
  Documents:    ['DocID', 'DocNo', 'Title', 'Description', 'CatID', 'DeptIDs', 'Tags', 'CreatedBy', 'CreatedAt', 'UpdatedAt',
                 'RegNo', 'DocDate', 'FromName', 'ToName', 'RefTo', 'Enclosures', 'PurposeID', 'Handler', 'Location',
                 'RetainYears', 'ExpireDate', 'Follow', 'DueDays', 'DueMode', 'DueDate', 'Remark', 'Status', 'DoneAt'],
  Files:        ['FileID', 'DocID', 'Type', 'Name', 'MimeType', 'Size', 'DriveFileId', 'Url', 'AccessMode', 'AddedBy', 'AddedAt'],
  Templates:    ['TplID', 'Name', 'CatID', 'Data', 'CreatedBy', 'CreatedAt'],
  DownloadLogs: ['LogID', 'FileID', 'DocID', 'FileName', 'DocTitle', 'UserID', 'UserName', 'DeptID', 'Action', 'At']
};

/* ---------- Web app entry ---------- */
function doGet() {
  ensureSetup_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(getSetting_('AppName') || 'eDMS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Setup (รันครั้งแรกด้วยมือ 1 ครั้ง) ---------- */
const SCHEMA_VER = '2';
const CAT_SEED = [
  ['C001', 'หนังสือรับเข้า', 'Incoming letter', 'in', ''],
  ['C002', 'หนังสือส่งภายใน', 'Internal letter', 'internal', 'กห0619.8/'],
  ['C003', 'หนังสือส่งภายนอก', 'Outgoing letter', 'out', ''],
  ['C004', 'คำสั่ง / ประกาศ', 'Orders / Notices', 'other', ''],
  ['C005', 'รายงาน', 'Reports', 'other', ''],
  ['C006', 'อื่นๆ', 'Other', 'other', '']
];
const PURPOSE_SEED = [
  ['P001', 'เพื่อดำเนินการ', 'For action'], ['P002', 'เพื่อทราบ', 'For information'],
  ['P003', 'เพื่อพิจารณา', 'For consideration'], ['P004', 'เพื่ออนุมัติ', 'For approval'],
  ['P005', 'เพื่อโปรดลงนาม', 'For signature'], ['P006', 'เพื่อเก็บเข้าแฟ้ม', 'For filing']
];
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(SCHEMA).forEach(function (name) {
    const cols = SCHEMA[name];
    const s = ss.getSheetByName(name) || ss.insertSheet(name);
    if (s.getMaxColumns() < cols.length) s.insertColumnsAfter(s.getMaxColumns(), cols.length - s.getMaxColumns());
    s.getRange(1, 1, s.getMaxRows(), cols.length).setNumberFormat('@'); // เก็บเป็นข้อความทั้งหมด
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, cols.length).setValues([cols])
        .setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
      s.setFrozenRows(1);
    } else {
      // migration: เพิ่มคอลัมน์ใหม่ต่อท้ายหัวตารางเดิม (ไม่กระทบข้อมูลเก่า)
      const have = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String);
      const missing = cols.filter(function (c) { return have.indexOf(c) < 0; });
      if (missing.length) {
        s.getRange(1, have.length + 1, 1, missing.length).setValues([missing])
          .setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
      }
    }
  });
  const dflt = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (dflt && ss.getSheets().length > 1 && dflt.getLastRow() === 0) ss.deleteSheet(dflt);

  const folderId = getSetting_('RootFolderId') || DriveApp.createFolder('eDMS_Files').getId();
  [['AppName', 'eDMS'], ['DefaultMode', 'download'], ['MaxFileMB', '20'], ['DefaultLang', 'th'],
   ['OrgName', 'รพ.กองบิน บน.5'], ['RootFolderId', folderId]].forEach(function (r) {
    if (!getSetting_(r[0])) setSetting_(r[0], r[1]);
  });
  if (!rows_('Departments').length) {
    [['D001', 'ฝ่ายบริหาร', 'Administration'], ['D002', 'ฝ่ายสารสนเทศ', 'Information Technology'],
     ['D003', 'ฝ่ายการเงิน', 'Finance']].forEach(function (r) {
      insert_('Departments', { DeptID: r[0], NameTH: r[1], NameEN: r[2], Active: 'TRUE' });
    });
  }
  const cats = rows_('Categories');
  if (!cats.length) {
    CAT_SEED.forEach(function (r) {
      insert_('Categories', { CatID: r[0], NameTH: r[1], NameEN: r[2], Active: 'TRUE', Direction: r[3], Prefix: r[4] });
    });
  } else if (!cats.some(function (c) { return ['in', 'internal', 'out'].indexOf(c.Direction) >= 0; })) {
    // อัปเกรดจากเวอร์ชันเดิม: เพิ่มประเภทหนังสือรับ/ส่ง
    cats.forEach(function (c) { if (!c.Direction) update_('Categories', c.CatID, { Direction: 'other' }); });
    CAT_SEED.slice(0, 3).forEach(function (r) {
      insert_('Categories', { CatID: uid_('C'), NameTH: r[1], NameEN: r[2], Active: 'TRUE', Direction: r[3], Prefix: r[4] });
    });
  }
  if (!rows_('Purposes').length) {
    PURPOSE_SEED.forEach(function (r) { insert_('Purposes', { PurposeID: r[0], NameTH: r[1], NameEN: r[2], Active: 'TRUE' }); });
  }
  if (!rows_('Users').length) {
    const salt = Utilities.getUuid().slice(0, 8);
    insert_('Users', {
      UserID: 'U001', Username: 'admin', Salt: salt, PasswordHash: hash_(salt, 'Admin@1234'),
      FullName: 'ผู้ดูแลระบบ', DeptID: 'D001', Role: 'admin', CanUpload: 'TRUE', CanDownload: 'TRUE',
      Active: 'TRUE', Lang: 'th', CreatedAt: now_(), LastLogin: ''
    });
  }
  PropertiesService.getScriptProperties().setProperty('SETUP_DONE', SCHEMA_VER);
  Logger.log('eDMS setup/migrate เสร็จ — ล็อกอินครั้งแรก: admin / Admin@1234 (กรุณาเปลี่ยนรหัสผ่านทันที)');
}
function ensureSetup_() {
  if (PropertiesService.getScriptProperties().getProperty('SETUP_DONE') !== SCHEMA_VER) setup();
}

/* ---------- 2) DB helpers ---------- */
function sh_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function now_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function uid_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase(); }
function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function str_(v) { return v === undefined || v === null ? '' : String(v); }

function rows_(name) {
  const s = sh_(name);
  if (!s || s.getLastRow() < 2) return [];
  const v = s.getDataRange().getValues();
  const h = v.shift();
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const o = {};
    h.forEach(function (k, i) {
      let x = r[i];
      if (x instanceof Date) x = Utilities.formatDate(x, TZ, "yyyy-MM-dd'T'HH:mm:ss");
      o[k] = x === undefined ? '' : x;
    });
    return o;
  });
}
let LOCK_DEPTH_ = 0;
function withLock_(fn) {
  if (LOCK_DEPTH_ > 0) return fn(); // ซ้อนล็อกได้ในการทำงานเดียวกัน
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  LOCK_DEPTH_++;
  try { return fn(); } finally { LOCK_DEPTH_--; lock.releaseLock(); }
}
function insert_(name, obj) {
  withLock_(function () {
    sh_(name).appendRow(SCHEMA[name].map(function (c) { return str_(obj[c]); }));
  });
}
function rowIndex_(s, id) {
  const last = s.getLastRow();
  if (last < 2) return -1;
  const ids = s.getRange(1, 1, last, 1).getValues().map(function (r) { return String(r[0]); });
  return ids.indexOf(String(id)) + 1; // 0 = ไม่พบ (แถวหัวตาราง = 1 ไม่นับ)
}
function update_(name, id, patch) {
  withLock_(function () {
    const s = sh_(name), cols = SCHEMA[name];
    const i = rowIndex_(s, id);
    if (i < 2) throw new Error('NOT_FOUND');
    const rng = s.getRange(i, 1, 1, cols.length);
    const cur = rng.getValues()[0];
    cols.forEach(function (c, k) { if (patch[c] !== undefined) cur[k] = str_(patch[c]); });
    rng.setValues([cur]);
  });
}
function remove_(name, id) {
  withLock_(function () {
    const s = sh_(name), i = rowIndex_(s, id);
    if (i >= 2) s.deleteRow(i);
  });
}
function getSetting_(key) {
  const r = rows_('Config').filter(function (x) { return x.Key === key; })[0];
  return r ? String(r.Value) : '';
}
function setSetting_(key, value) {
  if (rowIndex_(sh_('Config'), key) >= 2) update_('Config', key, { Value: value });
  else insert_('Config', { Key: key, Value: value });
}

/* ---------- 3) AUTH module ---------- */
function hash_(salt, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + String(password));
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function pubUser_(u) {
  const admin = u.Role === 'admin';
  return {
    UserID: u.UserID, Username: u.Username, FullName: u.FullName, DeptID: u.DeptID, Role: u.Role,
    CanUpload: admin || bool_(u.CanUpload), CanDownload: admin || bool_(u.CanDownload),
    Active: bool_(u.Active), Lang: u.Lang || 'th', CreatedAt: u.CreatedAt, LastLogin: u.LastLogin
  };
}
function login_(_u, p) {
  const un = String(p.username || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const fk = 'fail_' + un;
  const fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('LOCKED');
  const u = rows_('Users').filter(function (x) { return String(x.Username).toLowerCase() === un; })[0];
  if (!u || !bool_(u.Active) || hash_(u.Salt, p.password || '') !== u.PasswordHash) {
    cache.put(fk, String(fails + 1), 300);
    throw new Error('BAD_LOGIN');
  }
  cache.remove(fk);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('s_' + token, u.UserID, SESSION_SEC);
  update_('Users', u.UserID, { LastLogin: now_() });
  return { token: token, user: pubUser_(u) };
}
function logout_(_u, _p, token) { CacheService.getScriptCache().remove('s_' + token); return true; }
function auth_(token) {
  if (!token) throw new Error('SESSION_EXPIRED');
  const cache = CacheService.getScriptCache();
  const id = cache.get('s_' + token);
  if (!id) throw new Error('SESSION_EXPIRED');
  const u = rows_('Users').filter(function (x) { return x.UserID === id; })[0];
  if (!u || !bool_(u.Active)) throw new Error('SESSION_EXPIRED');
  cache.put('s_' + token, id, SESSION_SEC);
  return u;
}
function changePassword_(user, p) {
  if (hash_(user.Salt, p.oldPassword || '') !== user.PasswordHash) throw new Error('BAD_OLD_PASSWORD');
  if (String(p.newPassword || '').length < 6) throw new Error('WEAK_PASSWORD');
  const salt = Utilities.getUuid().slice(0, 8);
  update_('Users', user.UserID, { Salt: salt, PasswordHash: hash_(salt, p.newPassword) });
  return true;
}
function setLang_(user, p) {
  update_('Users', user.UserID, { Lang: p.lang === 'en' ? 'en' : 'th' });
  return true;
}

/* ---------- 4) ACCESS module ---------- */
function isAdmin_(u) { return u.Role === 'admin'; }
function canUpload_(u) { return isAdmin_(u) || bool_(u.CanUpload); }
function canDownload_(u) { return isAdmin_(u) || bool_(u.CanDownload); }
function deptList_(doc) {
  return String(doc.DeptIDs || '').split(',').map(function (x) { return x.trim(); }).filter(String);
}
function deptMatch_(user, doc) {
  const d = deptList_(doc);
  return d.indexOf('ALL') >= 0 || d.indexOf(user.DeptID) >= 0;
}
/** เห็นเอกสาร: แอดมิน | (มีสิทธิ์สืบค้น + แผนกตรงกัน) | เจ้าของเอกสาร */
function canSee_(user, doc) {
  if (isAdmin_(user)) return true;
  if (doc.CreatedBy === user.UserID) return true;
  return canDownload_(user) && deptMatch_(user, doc);
}
function canEdit_(user, doc) {
  return isAdmin_(user) || (canUpload_(user) && doc.CreatedBy === user.UserID);
}

/* ---------- 5) DOCS / FILES module ---------- */
const VIEWABLE_ = /^(image\/(png|jpe?g|gif|webp|svg\+xml)|application\/pdf|text\/plain)$/;
function viewable_(mime, name) {
  return VIEWABLE_.test(String(mime || '')) || /\.(pdf|png|jpe?g|gif|webp|txt)$/i.test(String(name || ''));
}
function userNames_() {
  const m = {};
  rows_('Users').forEach(function (u) { m[u.UserID] = u.FullName; });
  return m;
}
function groupFiles_() {
  const g = {};
  rows_('Files').forEach(function (f) { (g[f.DocID] = g[f.DocID] || []).push(f); });
  return g;
}
function decorateDoc_(user, d, files, names, done) {
  const defMode = getSetting_('DefaultMode') || 'download';
  const follow = bool_(d.Follow), status = d.Status || (follow ? 'open' : ''), due = d10_(d.DueDate);
  return {
    DocID: d.DocID, DocNo: d.DocNo, Title: d.Title, Description: d.Description, CatID: d.CatID,
    DeptIDs: deptList_(d), Tags: d.Tags, CreatedBy: d.CreatedBy, CreatedByName: names[d.CreatedBy] || '-',
    CreatedAt: d.CreatedAt, UpdatedAt: d.UpdatedAt, canEdit: canEdit_(user, d),
    canGet: isAdmin_(user) || (canDownload_(user) && deptMatch_(user, d)),
    RegNo: d.RegNo || '', DocDate: d10_(d.DocDate) || d10_(d.CreatedAt), FromName: d.FromName || '', ToName: d.ToName || '',
    RefTo: d.RefTo || '', Enclosures: d.Enclosures || '', PurposeID: d.PurposeID || '', Handler: d.Handler || '',
    Location: d.Location || '', RetainYears: Number(d.RetainYears) || 0, ExpireDate: d10_(d.ExpireDate),
    Follow: follow, DueDays: Number(d.DueDays) || 0, DueMode: d.DueMode === 'workday' ? 'workday' : 'all', DueDate: due,
    Remark: d.Remark || '', Status: status, DoneAt: d10_(d.DoneAt),
    overdue: follow && status !== 'done' && !!due && due < todayIso_(),
    files: (files || []).map(function (f) {
      const mode = f.AccessMode === 'view' || f.AccessMode === 'download' ? f.AccessMode : defMode;
      return {
        FileID: f.FileID, Type: f.Type, Name: f.Name, MimeType: f.MimeType, Size: Number(f.Size) || 0,
        AccessMode: f.AccessMode || 'default',
        mode: f.Type === 'link' ? 'link' : (mode === 'view' && viewable_(f.MimeType, f.Name) ? 'view' : 'download'),
        downloaded: !!done[f.FileID]
      };
    })
  };
}
function docsList_(user, p) {
  const files = groupFiles_(), names = userNames_();
  const done = {};
  rows_('DownloadLogs').forEach(function (l) { if (l.UserID === user.UserID) done[l.FileID] = 1; });
  const q = String(p.q || '').toLowerCase().trim();
  let docs = rows_('Documents').filter(function (d) { return canSee_(user, d); });
  if (p.catId) docs = docs.filter(function (d) { return d.CatID === p.catId; });
  if (p.deptId) docs = docs.filter(function (d) { const l = deptList_(d); return l.indexOf(p.deptId) >= 0 || l.indexOf('ALL') >= 0; });
  if (p.follow) {
    const today = todayIso_();
    docs = docs.filter(function (d) {
      if (!bool_(d.Follow)) return false;
      const st = d.Status || 'open', due = d10_(d.DueDate);
      if (p.follow === 'done') return st === 'done';
      if (p.follow === 'open') return st !== 'done';
      if (p.follow === 'overdue') return st !== 'done' && !!due && due < today;
      return true;
    });
  }
  if (q) {
    docs = docs.filter(function (d) {
      const hay = [d.RegNo, d.DocNo, d.Title, d.Description, d.Tags, d.FromName, d.ToName, d.RefTo, d.Handler, d.Remark]
        .concat((files[d.DocID] || []).map(function (f) { return f.Name; })).join(' ').toLowerCase();
      return hay.indexOf(q) >= 0;
    });
  }
  if (p.follow) docs.sort(function (a, b) { return (d10_(a.DueDate) || '9999').localeCompare(d10_(b.DueDate) || '9999'); });
  else docs.sort(function (a, b) { return String(b.CreatedAt).localeCompare(String(a.CreatedAt)); });
  const size = Math.min(Number(p.size) || 10, 50), page = Math.max(Number(p.page) || 1, 1);
  const slice = docs.slice((page - 1) * size, page * size);
  return {
    total: docs.length, page: page, size: size,
    items: slice.map(function (d) { return decorateDoc_(user, d, files[d.DocID], names, done); })
  };
}
function todayIso_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function isoOk_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }
function d10_(v) { return String(v || '').slice(0, 10); }
function addDays_(iso, n, workdayOnly) {
  const p = iso.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  let left = n;
  while (left > 0) {
    d.setUTCDate(d.getUTCDate() + 1);
    const w = d.getUTCDay();
    if (!workdayOnly || (w !== 0 && w !== 6)) left--;
  }
  return d.toISOString().slice(0, 10);
}
function addYears_(iso, y) {
  const p = iso.split('-').map(Number);
  return new Date(Date.UTC(p[0] + y, p[1] - 1, p[2])).toISOString().slice(0, 10);
}
/** เลขทะเบียน: ลำดับ/ปี พ.ศ. แยกชุด รับ (in) / ส่ง (internal+out) / อื่นๆ */
function regGroup_(dir) { return dir === 'in' ? 'in' : (dir === 'internal' || dir === 'out' ? 'out' : 'other'); }
function nextRegNo_(catId) {
  const dirOf = {};
  rows_('Categories').forEach(function (c) { dirOf[c.CatID] = c.Direction || 'other'; });
  const grp = regGroup_(dirOf[catId]);
  const be = Number(Utilities.formatDate(new Date(), TZ, 'yyyy')) + 543;
  let max = 0;
  rows_('Documents').forEach(function (d) {
    if (regGroup_(dirOf[d.CatID]) !== grp) return;
    const m = String(d.RegNo).match(/^(\d+)\/(\d{4})$/);
    if (m && Number(m[2]) === be) max = Math.max(max, Number(m[1]));
  });
  return (max + 1) + '/' + be;
}
function docsSave_(user, p) {
  if (!canUpload_(user)) throw new Error('FORBIDDEN');
  const cat = rows_('Categories').filter(function (c) { return c.CatID === p.catId; })[0];
  if (!cat) throw new Error('TYPE_REQUIRED');
  const docNo = String(p.docNo || '').trim();
  if (!docNo || docNo === String(cat.Prefix || '').trim()) throw new Error('DOCNO_REQUIRED');
  const docDate = String(p.docDate || '').slice(0, 10);
  if (!isoOk_(docDate)) throw new Error('DATE_REQUIRED');
  const fromName = String(p.fromName || '').trim(), toName = String(p.toName || '').trim();
  if (!fromName) throw new Error('FROM_REQUIRED');
  if (!toName) throw new Error('TO_REQUIRED');
  const title = String(p.title || '').trim();
  if (!title) throw new Error('TITLE_REQUIRED');
  const deptIds = (p.deptIds || []).map(String).filter(String);
  if (!deptIds.length) throw new Error('DEPT_REQUIRED');
  const depts = rows_('Departments').map(function (d) { return d.DeptID; });
  deptIds.forEach(function (d) { if (d !== 'ALL' && depts.indexOf(d) < 0) throw new Error('NO_DEPT'); });
  const purposeId = rows_('Purposes').some(function (x) { return x.PurposeID === p.purposeId; }) ? p.purposeId : '';
  const retain = Math.max(0, Math.min(200, Math.floor(Number(p.retainYears) || 0)));
  let expire = String(p.expireDate || '').slice(0, 10);
  if (!isoOk_(expire)) expire = retain ? addYears_(docDate, retain) : '';
  const dueDays = Math.max(0, Math.min(3650, Math.floor(Number(p.dueDays) || 0)));
  const dueMode = p.dueMode === 'workday' ? 'workday' : 'all';
  let due = String(p.dueDate || '').slice(0, 10);
  if (!isoOk_(due)) due = dueDays ? addDays_(docDate, dueDays, dueMode === 'workday') : '';
  const follow = !!p.follow;
  const fields = {
    DocNo: docNo, Title: title, Description: String(p.description || '').trim(), CatID: cat.CatID,
    DeptIDs: deptIds.indexOf('ALL') >= 0 ? 'ALL' : deptIds.join(','), Tags: String(p.tags || '').trim(), UpdatedAt: now_(),
    DocDate: docDate, FromName: fromName, ToName: toName, RefTo: String(p.refTo || '').trim(),
    Enclosures: String(p.enclosures || '').trim(), PurposeID: purposeId, Handler: String(p.handler || '').trim(),
    Location: String(p.location || '').trim(), RetainYears: retain || '', ExpireDate: expire,
    Follow: follow ? 'TRUE' : 'FALSE', DueDays: dueDays || '', DueMode: dueMode, DueDate: due,
    Remark: String(p.remark || '').trim()
  };
  return withLock_(function () {
    const dup = rows_('Documents').some(function (x) {
      return x.DocID !== p.docId && x.CatID === cat.CatID && String(x.DocNo).trim().toLowerCase() === docNo.toLowerCase();
    });
    if (dup) throw new Error('DUP_DOCNO');
    let docId = p.docId, regNo;
    if (docId) {
      const d = rows_('Documents').filter(function (x) { return x.DocID === docId; })[0];
      if (!d) throw new Error('NOT_FOUND');
      if (!canEdit_(user, d)) throw new Error('FORBIDDEN');
      regNo = d.RegNo || nextRegNo_(cat.CatID);
      fields.RegNo = regNo;
      fields.Status = follow ? (d.Status === 'done' ? 'done' : 'open') : '';
      if (!follow) fields.DoneAt = '';
      update_('Documents', docId, fields);
    } else {
      docId = uid_('DOC');
      regNo = nextRegNo_(cat.CatID);
      fields.DocID = docId; fields.RegNo = regNo; fields.CreatedBy = user.UserID; fields.CreatedAt = now_();
      fields.Status = follow ? 'open' : ''; fields.DoneAt = '';
      insert_('Documents', fields);
    }
    (p.removeFileIds || []).forEach(function (fid) { removeFile_(fid, docId); });
    Object.keys(p.modes || {}).forEach(function (fid) {
      const m = p.modes[fid];
      if (['default', 'download', 'view'].indexOf(m) >= 0) update_('Files', fid, { AccessMode: m });
    });
    (p.links || []).forEach(function (l) {
      const url = String(l.url || '').trim();
      if (!/^https?:\/\//i.test(url)) throw new Error('BAD_URL');
      insert_('Files', {
        FileID: uid_('F'), DocID: docId, Type: 'link', Name: String(l.name || '').trim() || url,
        MimeType: 'text/uri-list', Size: 0, DriveFileId: '', Url: url, AccessMode: 'default',
        AddedBy: user.UserID, AddedAt: now_()
      });
    });
    return { docId: docId, regNo: regNo };
  });
}
function removeFile_(fileId, docId) {
  const f = rows_('Files').filter(function (x) { return x.FileID === fileId && x.DocID === docId; })[0];
  if (!f) return;
  if (f.DriveFileId) { try { DriveApp.getFileById(f.DriveFileId).setTrashed(true); } catch (e) {} }
  remove_('Files', fileId);
}
function docsDelete_(user, p) {
  const d = rows_('Documents').filter(function (x) { return x.DocID === p.docId; })[0];
  if (!d) throw new Error('NOT_FOUND');
  if (!canEdit_(user, d)) throw new Error('FORBIDDEN');
  rows_('Files').filter(function (f) { return f.DocID === d.DocID; }).forEach(function (f) { removeFile_(f.FileID, d.DocID); });
  remove_('Documents', d.DocID);
  return true;
}
function filesUpload_(user, p) {
  if (!canUpload_(user)) throw new Error('FORBIDDEN');
  const d = rows_('Documents').filter(function (x) { return x.DocID === p.docId; })[0];
  if (!d) throw new Error('NOT_FOUND');
  if (!canEdit_(user, d)) throw new Error('FORBIDDEN');
  const maxMB = Number(getSetting_('MaxFileMB')) || 20;
  const approx = String(p.base64 || '').length * 0.75;
  if (!p.base64 || approx > maxMB * 1024 * 1024) throw new Error('FILE_TOO_LARGE');
  const name = String(p.name || 'file').replace(/[\\\/:*?"<>|]/g, '_');
  const blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.mime || 'application/octet-stream', name);
  const folder = DriveApp.getFolderById(getSetting_('RootFolderId'));
  const fid = uid_('F');
  const gf = folder.createFile(blob).setName(fid + '_' + name);
  const mode = ['default', 'download', 'view'].indexOf(p.mode) >= 0 ? p.mode : 'default';
  insert_('Files', {
    FileID: fid, DocID: d.DocID, Type: 'file', Name: name, MimeType: p.mime || '', Size: blob.getBytes().length,
    DriveFileId: gf.getId(), Url: '', AccessMode: mode, AddedBy: user.UserID, AddedAt: now_()
  });
  update_('Documents', d.DocID, { UpdatedAt: now_() });
  return { fileId: fid };
}
/** ดึงไฟล์/ลิงก์ + บันทึกประวัติ (ตรวจสิทธิ์ที่ฝั่งเซิร์ฟเวอร์ทุกครั้ง) */
function filesFetch_(user, p) {
  const f = rows_('Files').filter(function (x) { return x.FileID === p.fileId; })[0];
  if (!f) throw new Error('NOT_FOUND');
  const d = rows_('Documents').filter(function (x) { return x.DocID === f.DocID; })[0];
  if (!d) throw new Error('NOT_FOUND');
  if (!(isAdmin_(user) || (canDownload_(user) && deptMatch_(user, d)))) throw new Error('FORBIDDEN');
  let action = 'download', out;
  if (f.Type === 'link') {
    action = 'link';
    out = { type: 'link', url: f.Url, name: f.Name };
  } else {
    const defMode = getSetting_('DefaultMode') || 'download';
    const want = p.mode === 'view' || p.mode === 'download' ? p.mode : (f.AccessMode === 'default' ? defMode : f.AccessMode);
    action = want === 'view' && viewable_(f.MimeType, f.Name) ? 'view' : 'download';
    const blob = DriveApp.getFileById(f.DriveFileId).getBlob();
    out = { type: 'file', name: f.Name, mime: f.MimeType || blob.getContentType(), mode: action, base64: Utilities.base64Encode(blob.getBytes()) };
  }
  insert_('DownloadLogs', {
    LogID: uid_('L'), FileID: f.FileID, DocID: d.DocID, FileName: f.Name, DocTitle: d.Title,
    UserID: user.UserID, UserName: user.FullName, DeptID: user.DeptID, Action: action, At: now_()
  });
  return out;
}
function logsList_(user, p) {
  let logs = rows_('DownloadLogs');
  if (p.mine) logs = logs.filter(function (l) { return l.UserID === user.UserID; });
  else if (!isAdmin_(user)) throw new Error('FORBIDDEN');
  const q = String(p.q || '').toLowerCase().trim();
  if (q) logs = logs.filter(function (l) { return [l.FileName, l.DocTitle, l.UserName].join(' ').toLowerCase().indexOf(q) >= 0; });
  if (p.action) logs = logs.filter(function (l) { return l.Action === p.action; });
  if (p.deptId) logs = logs.filter(function (l) { return l.DeptID === p.deptId; });
  logs.sort(function (a, b) { return String(b.At).localeCompare(String(a.At)); });
  const size = Math.min(Number(p.size) || 20, 2000), page = Math.max(Number(p.page) || 1, 1);
  return { total: logs.length, page: page, size: size, items: logs.slice((page - 1) * size, page * size) };
}

/* ---------- 6) ADMIN module ---------- */
function usersList_() {
  return rows_('Users').map(pubUser_);
}
function usersSave_(admin, p) {
  const username = String(p.username || '').trim().toLowerCase();
  const fullName = String(p.fullName || '').trim();
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('BAD_USERNAME');
  if (!fullName) throw new Error('NAME_REQUIRED');
  const dept = rows_('Departments').filter(function (d) { return d.DeptID === p.deptId; })[0];
  if (!dept) throw new Error('DEPT_REQUIRED');
  const all = rows_('Users');
  if (all.some(function (u) { return String(u.Username).toLowerCase() === username && u.UserID !== p.userId; })) throw new Error('USERNAME_TAKEN');
  const role = p.role === 'admin' ? 'admin' : 'member';
  const fields = {
    Username: username, FullName: fullName, DeptID: p.deptId, Role: role,
    CanUpload: p.canUpload ? 'TRUE' : 'FALSE', CanDownload: p.canDownload ? 'TRUE' : 'FALSE',
    Active: p.active === false ? 'FALSE' : 'TRUE', Lang: p.lang === 'en' ? 'en' : 'th'
  };
  if (p.userId) {
    if (p.userId === admin.UserID && (role !== 'admin' || fields.Active === 'FALSE')) throw new Error('CANNOT_SELF_DEMOTE');
    update_('Users', p.userId, fields);
    return { userId: p.userId };
  }
  if (String(p.password || '').length < 6) throw new Error('WEAK_PASSWORD');
  const salt = Utilities.getUuid().slice(0, 8);
  fields.UserID = uid_('U'); fields.Salt = salt; fields.PasswordHash = hash_(salt, p.password);
  fields.CreatedAt = now_(); fields.LastLogin = '';
  insert_('Users', fields);
  return { userId: fields.UserID };
}
function usersResetPassword_(_a, p) {
  if (String(p.password || '').length < 6) throw new Error('WEAK_PASSWORD');
  const salt = Utilities.getUuid().slice(0, 8);
  update_('Users', p.userId, { Salt: salt, PasswordHash: hash_(salt, p.password) });
  return true;
}
function usersDelete_(admin, p) {
  if (p.userId === admin.UserID) throw new Error('CANNOT_SELF_DEMOTE');
  update_('Users', p.userId, { Active: 'FALSE' }); // ปิดการใช้งาน (คงประวัติไว้)
  return true;
}
function simpleSave_(sheet, idField, prefix) {
  return function (_a, p) {
    const th = String(p.nameTH || '').trim(), en = String(p.nameEN || '').trim();
    if (!th && !en) throw new Error('NAME_REQUIRED');
    const f = { NameTH: th || en, NameEN: en || th, Active: p.active === false ? 'FALSE' : 'TRUE' };
    if (p.id) { update_(sheet, p.id, f); return { id: p.id }; }
    f[idField] = uid_(prefix); insert_(sheet, f);
    return { id: f[idField] };
  };
}
function deptDelete_(_a, p) {
  if (rows_('Users').some(function (u) { return u.DeptID === p.id; })) throw new Error('IN_USE');
  remove_('Departments', p.id); return true;
}
function catDelete_(_a, p) {
  if (rows_('Documents').some(function (d) { return d.CatID === p.id; })) throw new Error('IN_USE');
  remove_('Categories', p.id); return true;
}
function settingsSave_(_a, p) {
  const name = String(p.AppName || '').trim();
  if (name) setSetting_('AppName', name);
  const org = String(p.OrgName || '').trim();
  if (org) setSetting_('OrgName', org);
  if (['download', 'view'].indexOf(p.DefaultMode) >= 0) setSetting_('DefaultMode', p.DefaultMode);
  const mb = Number(p.MaxFileMB);
  if (mb >= 1 && mb <= 50) setSetting_('MaxFileMB', String(Math.floor(mb)));
  if (['th', 'en'].indexOf(p.DefaultLang) >= 0) setSetting_('DefaultLang', p.DefaultLang);
  return publicSettings_();
}
function publicSettings_() {
  return {
    AppName: getSetting_('AppName') || 'eDMS', OrgName: getSetting_('OrgName') || '', DefaultMode: getSetting_('DefaultMode') || 'download',
    MaxFileMB: Number(getSetting_('MaxFileMB')) || 20, DefaultLang: getSetting_('DefaultLang') || 'th'
  };
}

/* ---------- Bootstrap + Dashboard ---------- */
function bootstrap_(user) {
  const departments = rows_('Departments').map(function (d) { return { DeptID: d.DeptID, NameTH: d.NameTH, NameEN: d.NameEN, Active: bool_(d.Active) }; });
  const categories = rows_('Categories').map(function (c) { return { CatID: c.CatID, NameTH: c.NameTH, NameEN: c.NameEN, Active: bool_(c.Active), Direction: c.Direction || 'other', Prefix: c.Prefix || '' }; });
  const purposes = rows_('Purposes').map(function (c) { return { PurposeID: c.PurposeID, NameTH: c.NameTH, NameEN: c.NameEN, Active: bool_(c.Active) }; });
  const people = canUpload_(user) ? rows_('Users').filter(function (u) { return bool_(u.Active); })
    .map(function (u) { return { UserID: u.UserID, FullName: u.FullName, DeptID: u.DeptID }; }) : [];
  const docs = rows_('Documents').filter(function (d) { return canSee_(user, d); });
  const files = groupFiles_(), names = userNames_();
  const logs = rows_('DownloadLogs');
  const mine = logs.filter(function (l) { return l.UserID === user.UserID; });
  const done = {}; mine.forEach(function (l) { done[l.FileID] = 1; });
  docs.sort(function (a, b) { return String(b.CreatedAt).localeCompare(String(a.CreatedAt)); });
  const month = now_().slice(0, 7), today = todayIso_();
  const fol = docs.filter(function (d) { return bool_(d.Follow) && (d.Status || 'open') !== 'done'; });
  const stats = {
    docs: docs.length,
    files: docs.reduce(function (n, d) { return n + (files[d.DocID] || []).length; }, 0),
    myDownloads: mine.length,
    myMonth: mine.filter(function (l) { return String(l.At).indexOf(month) === 0; }).length,
    followOpen: fol.length,
    followOverdue: fol.filter(function (d) { return d10_(d.DueDate) && d10_(d.DueDate) < today; }).length,
    followList: fol.slice().sort(function (a, b) { return (d10_(a.DueDate) || '9999').localeCompare(d10_(b.DueDate) || '9999'); }).slice(0, 5)
      .map(function (d) { return { DocID: d.DocID, Title: d.Title, DocNo: d.DocNo, DueDate: d10_(d.DueDate), overdue: !!d10_(d.DueDate) && d10_(d.DueDate) < today }; }),
    recent: docs.slice(0, 5).map(function (d) { return decorateDoc_(user, d, files[d.DocID], names, done); }),
    myRecent: mine.sort(function (a, b) { return String(b.At).localeCompare(String(a.At)); }).slice(0, 6)
  };
  if (isAdmin_(user)) {
    const users = rows_('Users');
    stats.users = users.filter(function (u) { return bool_(u.Active); }).length;
    stats.totalDownloads = logs.length;
    const cnt = {};
    logs.forEach(function (l) { cnt[l.FileName] = (cnt[l.FileName] || 0) + 1; });
    stats.top = Object.keys(cnt).map(function (k) { return { name: k, count: cnt[k] }; })
      .sort(function (a, b) { return b.count - a.count; }).slice(0, 5);
  }
  return { user: pubUser_(user), departments: departments, categories: categories, purposes: purposes, people: people,
           settings: publicSettings_(), stats: stats };
}

/* ---------- Follow-up / Templates / Pickers / Types ---------- */
function docsDone_(user, p) {
  const d = rows_('Documents').filter(function (x) { return x.DocID === p.docId; })[0];
  if (!d) throw new Error('NOT_FOUND');
  if (!canSee_(user, d)) throw new Error('FORBIDDEN');
  if (!bool_(d.Follow)) throw new Error('NOT_FOLLOW');
  const done = !!p.done;
  update_('Documents', d.DocID, { Status: done ? 'done' : 'open', DoneAt: done ? todayIso_() : '' });
  return true;
}
function tplList_(user) {
  return rows_('Templates').filter(function (t) { return t.CreatedBy === user.UserID; })
    .map(function (t) { return { TplID: t.TplID, Name: t.Name, CatID: t.CatID, Data: t.Data }; });
}
function tplSave_(user, p) {
  if (!canUpload_(user)) throw new Error('FORBIDDEN');
  const name = String(p.name || '').trim();
  if (!name) throw new Error('NAME_REQUIRED');
  const json = JSON.stringify(p.data || {});
  if (json.length > 40000) throw new Error('FILE_TOO_LARGE');
  const id = uid_('T');
  insert_('Templates', { TplID: id, Name: name, CatID: p.catId || '', Data: json, CreatedBy: user.UserID, CreatedAt: now_() });
  return { tplId: id };
}
function tplDelete_(user, p) {
  const t = rows_('Templates').filter(function (x) { return x.TplID === p.tplId; })[0];
  if (!t || t.CreatedBy !== user.UserID) throw new Error('NOT_FOUND');
  remove_('Templates', t.TplID);
  return true;
}
/** รายชื่อหน่วยงาน/ผู้รับ-ผู้ส่ง ที่เคยใช้ในเอกสารที่ผู้ใช้มองเห็น (สำหรับตัวเลือก "หน่วยงาน") */
function orgsList_(user) {
  const seen = {};
  rows_('Documents').filter(function (d) { return canSee_(user, d); }).forEach(function (d) {
    [d.FromName, d.ToName].forEach(function (v) {
      String(v || '').split(',').forEach(function (x) { x = x.trim(); if (x) seen[x] = 1; });
    });
  });
  return Object.keys(seen).sort().slice(0, 300);
}
function catSave_(_a, p) {
  const th = String(p.nameTH || '').trim(), en = String(p.nameEN || '').trim();
  if (!th && !en) throw new Error('NAME_REQUIRED');
  const f = {
    NameTH: th || en, NameEN: en || th, Active: p.active === false ? 'FALSE' : 'TRUE',
    Direction: ['in', 'internal', 'out', 'other'].indexOf(p.direction) >= 0 ? p.direction : 'other',
    Prefix: String(p.prefix || '').trim()
  };
  if (p.id) { update_('Categories', p.id, f); return { id: p.id }; }
  f.CatID = uid_('C'); insert_('Categories', f);
  return { id: f.CatID };
}
function purposeDelete_(_a, p) {
  if (rows_('Documents').some(function (d) { return d.PurposeID === p.id; })) throw new Error('IN_USE');
  remove_('Purposes', p.id); return true;
}

/* ---------- 7) ROUTER ---------- */
function ROUTES_() {
  return {
    'login':          { public: true, fn: login_ },
    'logout':         { fn: logout_ },
    'bootstrap':      { fn: bootstrap_ },
    'me.password':    { fn: changePassword_ },
    'me.lang':        { fn: setLang_ },
    'docs.list':      { fn: function (u, p) { if (!canDownload_(u) && !canUpload_(u)) throw new Error('FORBIDDEN'); return docsList_(u, p); } },
    'docs.save':      { fn: docsSave_ },
    'docs.delete':    { fn: docsDelete_ },
    'docs.done':      { fn: docsDone_ },
    'files.upload':   { fn: filesUpload_ },
    'files.fetch':    { fn: filesFetch_ },
    'tpl.list':       { fn: tplList_ },
    'tpl.save':       { fn: tplSave_ },
    'tpl.delete':     { fn: tplDelete_ },
    'orgs.list':      { fn: orgsList_ },
    'logs.list':      { fn: logsList_ },
    'users.list':     { admin: true, fn: usersList_ },
    'users.save':     { admin: true, fn: usersSave_ },
    'users.reset':    { admin: true, fn: usersResetPassword_ },
    'users.delete':   { admin: true, fn: usersDelete_ },
    'depts.save':     { admin: true, fn: simpleSave_('Departments', 'DeptID', 'D') },
    'depts.delete':   { admin: true, fn: deptDelete_ },
    'cats.save':      { admin: true, fn: catSave_ },
    'cats.delete':    { admin: true, fn: catDelete_ },
    'purposes.save':  { admin: true, fn: simpleSave_('Purposes', 'PurposeID', 'P') },
    'purposes.delete':{ admin: true, fn: purposeDelete_ },
    'settings.save':  { admin: true, fn: settingsSave_ }
  };
}
function api(method, token, payload) {
  try {
    const route = ROUTES_()[method];
    if (!route) throw new Error('NOT_FOUND');
    const user = route.public ? null : auth_(token);
    if (route.admin && user.Role !== 'admin') throw new Error('FORBIDDEN');
    return JSON.stringify({ ok: true, data: route.fn(user, payload || {}, token) });
  } catch (err) {
    return JSON.stringify({ ok: false, error: String((err && err.message) || err) });
  }
}

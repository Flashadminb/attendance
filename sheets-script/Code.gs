/**
 * Attendance API สำหรับ Google Sheets
 *
 * รหัสลับ (TOKEN) อยู่ในไฟล์ token.gs ซึ่งเก็บเฉพาะในเครื่อง ไม่ขึ้น GitHub
 *
 * ติดตั้งด้วย clasp (ในโฟลเดอร์ sheets-script):
 *   npx @google/clasp@2.4.2 push -f
 *   npx @google/clasp@2.4.2 deploy -i <deploymentId> -d "update"   (อัปเดต deployment เดิม URL ไม่เปลี่ยน)
 *
 * ติดตั้งเองแบบไม่ใช้ clasp:
 * 1. เปิด Google Sheet → ส่วนขยาย (Extensions) → Apps Script
 * 2. วางโค้ดไฟล์นี้ และสร้างไฟล์ token.gs ที่มี  const TOKEN = 'รหัสลับของคุณ';
 * 3. Deploy → New deployment → Web app · Execute as: Me · Who has access: Anyone
 * 4. ในหน้าเว็บ Attendance → แท็บ ตั้งค่า → Google Sheets ใส่ URL (/exec) และรหัสลับ
 *
 * โครงสร้างชีตที่สร้างให้อัตโนมัติ
 *   rec_YYYY-MM  ข้อมูลสแกนรายวัน แยกตามเดือน (1 แถว = 1 คน 1 วัน)
 *   _summary     ตัวเลขสรุปของบริษัท (ขาด/สาย/ออกก่อน) ต่อคนต่อเดือน
 *   _store       การตั้งค่า กะพนักงาน ใบเตือน ประวัติแก้ย้อนหลัง (JSON)
 *   _meta        เวลาแก้ไขล่าสุดของแต่ละชีต (ใช้ให้เว็บโหลดเฉพาะส่วนที่เปลี่ยน)
 * อย่าแก้ไขหรือเปลี่ยนชื่อชีตเหล่านี้ด้วยมือ
 */

// TOKEN ประกาศไว้ใน token.gs

/** กดรันฟังก์ชันนี้ 1 ครั้งในหน้า Apps Script เพื่ออนุญาตสิทธิ์ให้สคริปต์เข้าถึงชีต */
function authorize() {
  sheet('_meta', ['name', 'u']);
  Logger.log('อนุญาตสิทธิ์แล้ว ชีต: ' + ss().getName());
}

const REC_HEAD = ['key', 'date', 'id', 'name', 'pos', 'branch', 'empStatus', 'raw', 'src', 'u', 'prev'];
const SUM_HEAD = ['key', 'id', 'month', 'absent', 'lateN', 'lateMin', 'earlyN', 'earlyMin'];
const CHUNK = 45000; // 1 เซลล์เก็บได้ 50,000 ตัวอักษร

function doGet() {
  return out({ ok: true, msg: 'Attendance API พร้อมใช้งาน' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (typeof TOKEN === 'undefined' || !TOKEN) return out({ ok: false, error: 'ยังไม่ได้ตั้ง TOKEN ใน token.gs' });
    if (req.token !== TOKEN) return out({ ok: false, error: 'รหัสลับไม่ถูกต้อง' });
    const fns = { sync: syncAll, meta: meta, getMonths: getMonths, upsert: upsert, getStore: getStore, setStore: setStore };
    const fn = fns[req.action];
    if (!fn) return out({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + req.action });
    const res = fn(req);
    res.ok = true;
    return out(res);
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** หา/สร้างชีต และตั้งให้ทุกเซลล์เป็นข้อความ (กัน Sheets แปลงวันที่/รหัสพนักงานเอง) */
function sheet(name, head) {
  let sh = ss().getSheetByName(name);
  if (!sh) {
    sh = ss().insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRows(sh, width) {
  const n = sh.getLastRow();
  return n > 1 ? sh.getRange(2, 1, n - 1, width).getDisplayValues() : [];
}

/** เวลาแก้ไขล่าสุดของทุกเดือน และของข้อมูลตั้งค่า */
function meta() {
  const v = readRows(sheet('_meta', ['name', 'u']), 2);
  const months = {}, store = {};
  v.forEach(function (r) {
    const n = String(r[0]);
    if (n.indexOf('rec_') === 0) months[n.slice(4)] = { u: Number(r[1]) };
    else if (n.indexOf('store_') === 0) store[n.slice(6)] = Number(r[1]);
  });
  return { months: months, store: store };
}

function bump(name) {
  const sh = sheet('_meta', ['name', 'u']);
  const names = readRows(sh, 1).map(function (r) { return String(r[0]); });
  const u = Date.now();
  const i = names.indexOf(name);
  if (i < 0) sh.getRange(names.length + 2, 1, 1, 2).setValues([[name, String(u)]]);
  else sh.getRange(i + 2, 2).setValue(String(u));
  return u;
}

function getMonths(req) {
  const rows = {};
  (req.months || []).forEach(function (m) {
    const sh = ss().getSheetByName('rec_' + m);
    rows[m] = sh ? readRows(sh, REC_HEAD.length) : [];
  });
  const sumSh = ss().getSheetByName('_summary');
  const sums = sumSh ? readRows(sumSh, SUM_HEAD.length).filter(function (r) { return (req.months || []).indexOf(String(r[2])) >= 0; }) : [];
  return { rows: rows, sums: sums };
}

/** เพิ่มหรือแก้แถวตาม key (คอลัมน์แรก) */
function upsertSheet(sh, rows, width) {
  if (!rows.length) return;
  const data = readRows(sh, width);
  const idx = {};
  data.forEach(function (r, i) { idx[r[0]] = i; });
  rows.forEach(function (r) {
    const row = [];
    for (let j = 0; j < width; j++) row.push(r[j] == null ? '' : String(r[j]));
    const i = idx[row[0]];
    if (i === undefined) { idx[row[0]] = data.length; data.push(row); }
    else data[i] = row;
  });
  if (sh.getMaxRows() < data.length + 1) sh.insertRowsAfter(sh.getMaxRows(), data.length + 1 - sh.getMaxRows());
  sh.getRange(2, 1, data.length, width).setNumberFormat('@').setValues(data);
}

function upsert(req) {
  const byMonth = {};
  (req.rows || []).forEach(function (r) {
    const m = String(r[1]).slice(0, 7);
    if (!byMonth[m]) byMonth[m] = [];
    byMonth[m].push(r);
  });
  const months = {};
  Object.keys(byMonth).forEach(function (m) {
    upsertSheet(sheet('rec_' + m, REC_HEAD), byMonth[m], REC_HEAD.length);
    months[m] = { u: bump('rec_' + m) };
  });
  if (req.sums && req.sums.length) {
    upsertSheet(sheet('_summary', SUM_HEAD), req.sums, SUM_HEAD.length);
    // ให้เครื่องอื่นโหลดเดือนที่ตัวเลขสรุปเปลี่ยนด้วย
    const ms = {};
    req.sums.forEach(function (s) { ms[String(s[2])] = 1; });
    Object.keys(ms).forEach(function (m) { if (!months[m]) months[m] = { u: bump('rec_' + m) }; });
  }
  return { months: months };
}

function getStore(req) {
  const sh = sheet('_store', ['key']);
  const n = sh.getLastRow(), values = {};
  if (n < 2) return { values: values };
  const v = sh.getRange(2, 1, n - 1, sh.getLastColumn()).getDisplayValues();
  v.forEach(function (r) {
    if ((req.keys || []).indexOf(String(r[0])) >= 0) values[r[0]] = r.slice(1).map(function (c) { return String(c).slice(1); }).join('');
  });
  return { values: values };
}

function writeStore(key, value) {
  const sh = sheet('_store', ['key']);
  value = String(value || '');
  const parts = [];
  for (let i = 0; i < value.length; i += CHUNK) parts.push('~' + value.slice(i, i + CHUNK)); // ~ กันไม่ให้ขึ้นต้นด้วย = แล้วกลายเป็นสูตร
  const keys = readRows(sh, 1).map(function (r) { return String(r[0]); });
  const i = keys.indexOf(key);
  const row = i < 0 ? keys.length + 2 : i + 2;
  const need = parts.length + 1;
  if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  sh.getRange(row, 1, 1, sh.getMaxColumns()).clearContent();
  sh.getRange(row, 1, 1, need).setNumberFormat('@').setValues([[key].concat(parts)]);
}

function setStore(req) {
  writeStore(req.key, req.value);
  return { u: bump('store_' + req.key) };
}

/**
 * ซิงค์ทุกอย่างในคำขอเดียว: รับข้อมูลที่เว็บส่งมา แล้วตอบกลับเฉพาะสิ่งที่เว็บยังไม่มี
 * req.rows / req.sums / req.stores  = ข้อมูลที่ต้องบันทึก
 * req.have = { months:{เดือน:เวอร์ชัน}, store:{key:เวอร์ชัน} } ที่เว็บมีอยู่แล้ว
 * req.want = เดือนที่เว็บต้องการโหลด · req.recent = เครื่องใหม่: โหลดกี่เดือนล่าสุดก่อน
 */
function syncAll(req) {
  const metaSh = sheet('_meta', ['name', 'u']);
  const M = {};
  readRows(metaSh, 2).forEach(function (r) { M[String(r[0])] = Number(r[1]); });
  let touched = false;
  const now = Date.now();
  function bumpM(n) { const u = Math.max(now, (M[n] || 0) + 1); M[n] = u; touched = true; return u; }

  // 1) บันทึกข้อมูลสแกนแยกตามเดือน
  const by = {};
  (req.rows || []).forEach(function (r) { const m = String(r[1]).slice(0, 7); (by[m] = by[m] || []).push(r); });
  const wrote = {};
  Object.keys(by).forEach(function (m) {
    const prev = M['rec_' + m] || 0;
    upsertSheet(sheet('rec_' + m, REC_HEAD), by[m], REC_HEAD.length);
    wrote[m] = { prev: prev, u: bumpM('rec_' + m) };
  });
  if (req.sums && req.sums.length) {
    upsertSheet(sheet('_summary', SUM_HEAD), req.sums, SUM_HEAD.length);
    req.sums.forEach(function (s) { const m = String(s[2]); if (!wrote[m]) wrote[m] = { prev: M['rec_' + m] || 0, u: bumpM('rec_' + m) }; });
  }
  // 2) บันทึกการตั้งค่า / กะ / ใบเตือน
  const storeU = {};
  Object.keys(req.stores || {}).forEach(function (k) { writeStore(k, req.stores[k]); storeU[k] = bumpM('store_' + k); });
  if (touched) {
    const outRows = Object.keys(M).map(function (n) { return [n, String(M[n])]; });
    metaSh.getRange(2, 1, outRows.length, 2).setNumberFormat('@').setValues(outRows);
  }

  // 3) เลือกเฉพาะเดือนที่เว็บต้องโหลด
  const months = {}, store = {};
  Object.keys(M).forEach(function (n) {
    if (n.indexOf('rec_') === 0) months[n.slice(4)] = M[n];
    else if (n.indexOf('store_') === 0) store[n.slice(6)] = M[n];
  });
  const have = req.have || {}, hm = have.months || {}, hs = have.store || {};
  const send = {};
  (req.want || []).forEach(function (m) { if (months[m]) send[m] = 1; });
  Object.keys(hm).forEach(function (m) {
    if (!months[m]) return;
    const w = wrote[m];
    if (w && Number(hm[m]) === w.prev) return;     // เว็บเพิ่งส่งมาเอง และไม่มีใครแก้ระหว่างนั้น
    if (Number(hm[m]) < months[m]) send[m] = 1;    // มีเครื่องอื่นแก้
  });
  Object.keys(wrote).forEach(function (m) { if (!(m in hm) && wrote[m].prev > 0) send[m] = 1; });
  if (req.recent && !Object.keys(hm).length) Object.keys(months).sort().reverse().slice(0, req.recent).forEach(function (m) { send[m] = 1; });

  const rows = {};
  Object.keys(send).forEach(function (m) { const sh = ss().getSheetByName('rec_' + m); rows[m] = sh ? readRows(sh, REC_HEAD.length) : []; });
  const sumSh = ss().getSheetByName('_summary');
  const sums = Object.keys(send).length && sumSh ? readRows(sumSh, SUM_HEAD.length).filter(function (r) { return send[String(r[2])]; }) : [];
  const needStore = Object.keys(store).filter(function (k) { return !storeU[k] && (Number(hs[k]) || 0) < store[k]; });
  const storeVals = needStore.length ? getStore({ keys: needStore }).values : {};
  return { months: months, store: store, wrote: wrote, storeU: storeU, rows: rows, sums: sums, storeVals: storeVals };
}

/**********************************************************************
 * ระบบรายงานสถานการณ์ยาเสพติด จังหวัดหนองบัวลำภู (รายสัปดาห์)
 * Backend : Google Apps Script + Google Sheets
 * รูปแบบ  : JSONP (กัน CORS) | LockService | Upsert ไม่ทบยอด
 * --------------------------------------------------------------------
 * วิธีติดตั้ง
 *  1) สร้าง Google Sheet ใหม่ คัดลอก ID จาก URL มาใส่ SHEET_ID ด้านล่าง
 *  2) Extensions > Apps Script วางโค้ดนี้ทั้งหมด
 *  3) รันฟังก์ชัน setup() หนึ่งครั้ง (สร้างชีต + หัวตาราง)
 *  4) Deploy > New deployment > Web app
 *        Execute as : Me
 *        Who has access : Anyone
 *     คัดลอก URL .../exec ไปวางใน input.html และ index.html (ตัวแปร API)
 *  5) เปลี่ยน PIN ผ่านเมนู Script Properties (คีย์ ENTRY_PIN) หรือใช้ค่าเริ่มต้น
 **********************************************************************/

const SHEET_ID  = '17bdbyXnql6_I0np1yn3sCTXJtGRHYdE0XnoiYEx4NTg';   // <<< ใส่ ID ชีตของคุณ
const SHEET_NAME = 'WeeklyData';
// รหัส PIN ไม่เก็บในซอร์สโค้ด — ตั้งค่าใน Apps Script: Project Settings > Script properties
//   ENTRY_PIN = <PIN เจ้าหน้าที่>  (จำเป็น)
//   SUPER_PIN = <PIN ผู้ดูแล>      (ถ้าต้องการ)

// ลำดับคอลัมน์ในชีต (ห้ามสลับลำดับ)
const HEADERS = [
  'id','year','month','week','district',
  'cases','suspects','dealers','methPills','iceGrams',
  'intoTreatment','completedTreatment','xray','psychiatric',
  'notes','updatedAt','updatedBy'
];
// ฟิลด์ตัวเลขที่ต้องแปลงเป็น Number
const NUM_FIELDS = ['cases','suspects','dealers','methPills','iceGrams',
                    'intoTreatment','completedTreatment','xray','psychiatric'];

/* ---------- ติดตั้งครั้งแรก ---------- */
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // ตั้ง PIN ที่ Project Settings > Script properties (ENTRY_PIN) — ไม่กำหนดในโค้ด
  return 'setup done';
}

/* ---------- Router (JSONP) ---------- */
function doGet(e) {
  const p = e.parameter || {};
  const cb = p.callback || 'callback';
  let out;
  try {
    const action = p.action || 'init';
    if (action === 'init' || action === 'getData') out = { ok: true, rows: readAll() };
    else if (action === 'getOne')  out = { ok: true, row: readOne(p.year, p.month, p.week, p.district) };
    else if (action === 'save')    out = saveRow(p);
    else if (action === 'ping')    out = { ok: true, t: new Date().toISOString() };
    else if (action === 'ai')      out = llmSummary(p);
    else out = { ok: false, error: 'unknown action' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* ---------- อ่านทั้งหมด ---------- */
function readAll() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const o = {};
    head.forEach((h, j) => { o[h] = values[i][j]; });
    if (!o.id) continue;
    NUM_FIELDS.forEach(f => { o[f] = Number(o[f]) || 0; });
    o.year = Number(o.year); o.month = Number(o.month); o.week = Number(o.week);
    rows.push(o);
  }
  return rows;
}

/* ---------- อ่านแถวเดียว (สำหรับ prefill) ---------- */
function readOne(year, month, week, district) {
  const id = makeId(year, month, week, district);
  const rows = readAll();
  return rows.find(r => r.id === id) || null;
}

/* ---------- บันทึก (upsert ไม่ทบยอด) ---------- */
function saveRow(p) {
  if (!checkPin(p.pin)) return { ok: false, error: 'PIN ไม่ถูกต้อง' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const id = makeId(p.year, p.month, p.week, p.district);
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().flat();
    const idx = ids.indexOf(id);                  // หา id เดิม

    const rec = {
      id: id, year: Number(p.year), month: Number(p.month),
      week: Number(p.week), district: p.district,
      cases: num(p.cases), suspects: num(p.suspects), dealers: num(p.dealers),
      methPills: num(p.methPills), iceGrams: num(p.iceGrams),
      intoTreatment: num(p.intoTreatment), completedTreatment: num(p.completedTreatment),
      xray: num(p.xray), psychiatric: num(p.psychiatric),
      notes: p.notes || '', updatedAt: new Date().toISOString(), updatedBy: p.user || '-'
    };
    const rowArr = HEADERS.map(h => rec[h]);

    if (idx >= 0) {                               // มีอยู่แล้ว -> เขียนทับแถวเดิม
      sh.getRange(idx + 2, 1, 1, HEADERS.length).setValues([rowArr]);
      return { ok: true, mode: 'update', id: id };
    } else {                                      // ยังไม่มี -> เพิ่มแถวใหม่
      sh.appendRow(rowArr);
      return { ok: true, mode: 'insert', id: id };
    }
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ตรวจสิทธิ์ PIN (อ่านจาก Script Properties เท่านั้น ไม่มีรหัสในซอร์ส) ---------- */
function checkPin(pin) {
  pin = String(pin || '');
  const props = PropertiesService.getScriptProperties();
  const entryPin = props.getProperty('ENTRY_PIN');
  const superPin = props.getProperty('SUPER_PIN');
  if (entryPin && pin === String(entryPin)) return true;
  if (superPin && pin === String(superPin)) return true;
  return false;
}

/* ---------- สรุปด้วย AI ขั้นสูง (LLM ผ่าน Anthropic API) ----------
 * เปิดใช้: ตั้งค่า Script Properties คีย์ ANTHROPIC_KEY = <API key ของคุณ>
 * ถ้าไม่ตั้งค่า จะแจ้งว่ายังไม่พร้อม (แดชบอร์ดจะใช้สรุปแบบ rule-based แทน) */
function llmSummary(p) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!key) return { ok: false, error: 'ยังไม่ได้ตั้งค่า ANTHROPIC_KEY ใน Script Properties' };
  let stats = {};
  try { stats = JSON.parse(p.data || '{}'); } catch (e) {}
  const prompt = 'คุณเป็นนักวิเคราะห์ข้อมูลของศูนย์อำนวยการป้องกันและปราบปรามยาเสพติด จังหวัดหนองบัวลำภู '
    + 'จงเขียนบทสรุปผู้บริหารภาษาไทยกระชับ 4-6 บรรทัด จากข้อมูลสถานการณ์ต่อไปนี้ '
    + 'เน้นแนวโน้ม จุดที่ต้องเร่งรัด และข้อเสนอเชิงนโยบายสั้น ๆ (ห้ามแต่งตัวเลขเกินจากข้อมูล):\n'
    + JSON.stringify(stats, null, 2);
  try {
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001', max_tokens: 600,
        messages: [{ role: 'user', content: prompt }]
      })
    });
    const data = JSON.parse(res.getContentText());
    const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('\n');
    if (!text) return { ok: false, error: 'AI ไม่ตอบกลับ', raw: data };
    return { ok: true, summary: text };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/* ---------- helpers ---------- */
function makeId(y, m, w, d) { return [y, m, 'W' + w, d].join('|'); }
function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }

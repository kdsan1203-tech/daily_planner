/**
 * GAUSS 소결 기록 — Apps Script 웹앱
 *
 * 「소결 데이터」 시트의 기존 구조를 그대로 사용합니다.
 *   Jobs          : 슬라이서가 자동 등록하는 출력 작업 (읽기 + 소결완료 표시만)
 *   SinterBatches : 소결 1회 = 1행 (기존 18개 열 + 뒤에 추가 열)
 *   Parts         : 소결에 들어간 부품별 치수·무게·수축률·판정
 *   Photos        : 사진 목록 (이 프로그램이 새로 만듦)
 *   ActivityLog   : 누가 언제 무엇을 했는지
 *   Lookups       : 소재·판정·소결로 목록 (기존 시트에서 읽음)
 *
 * SHEET_ID 를 비워두면 이 스크립트가 붙어 있는 시트에 같은 구조를 만들어 씁니다.
 */

const SHEET_ID = '1XXW7sJW40M_DUqXltIbhVG39lDVO3asyC3wqLaPe7Rw';   // 소결 데이터
const LOT_SHEET_ID = '17uOLhQ3eiAFCP4VEKZXTwGm2t_bXOinfv4robQ1SBU0'; // 가우스 잉크 Lot No 관리 시스템
const PHOTO_FOLDER = 'GAUSS 소결 사진';

const SB_BASE = ['batch_id','date','furnace','atmosphere_gas','gas_flow_lpm','vacuum',
  'initial_temp_c','initial_hold_min','ramp_up_c_min','peak_temp_c','peak_hold_min',
  'ramp_down_c_min','cool_temp_c','cool_hold_min','furnace_position','log_file_url','operator','notes'];
const SB_EXTRA = ['status','material','lot_no','steps_json','total_min',
  'cylinder_l','cyl_bar_start','cyl_bar_end','purge_count','gas_margin',
  'check_baffle','check_door','check_program','result_summary',
  'created_by','created_at','updated_at','deleted'];

const PARTS_BASE = ['part_id','job_id','batch_id','green_x_mm','green_y_mm','green_z_mm','green_wt_g',
  'sint_x_mm','sint_y_mm','sint_z_mm','sint_wt_g','shrink_x_pct','shrink_y_pct','shrink_z_pct',
  'wt_loss_pct','density_g_cm3','grade','photo_url','notes'];
const PARTS_EXTRA = ['part_name','updated_at','deleted'];

const PHOTOS_HEADERS = ['batch_id','part_id','kind','file_id','url','file_name','uploaded_by','uploaded_at','deleted'];
const LOG_HEADERS = ['at','user','action','batch_id','detail'];

const DEFAULT_LOOKUPS = {
  material: ['316L SUS','Aluminum','Copper','Iron','Nickel','Titanium','Tungsten'],
  grade: ['정상','크랙','휨','붕괴'],
  furnace: ['소결로1']
};

/* ── 진입점 ─────────────────────────────────────────── */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GAUSS 소결 기록')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 최초 1회: 편집기에서 실행해 권한 승인 + 시트 준비 */
function setup() {
  const ss = ss_();
  table_(ss, 'SinterBatches', SB_BASE, SB_EXTRA);
  table_(ss, 'Parts', PARTS_BASE, PARTS_EXTRA);
  table_(ss, 'Photos', PHOTOS_HEADERS, []);
  table_(ss, 'ActivityLog', LOG_HEADERS, []);
  folder_();
  Logger.log('준비 완료: ' + ss.getName());
}

/* ── 시트 유틸 ──────────────────────────────────────── */

function ss_() {
  if (SHEET_ID) {
    try { return SpreadsheetApp.openById(SHEET_ID); }
    catch (e) {
      throw new Error('「소결 데이터」 시트를 열 수 없습니다. 시트 소유자에게 편집 권한을 받거나, ' +
        'Code.gs 맨 위 SHEET_ID 를 비워 이 스크립트의 시트를 쓰세요.');
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** 시트를 열고(없으면 만들고) 빠진 열은 맨 뒤에 추가한 뒤 {sh, h, col} 반환 */
function table_(ss, name, base, extra) {
  let sh = ss.getSheetByName(name);
  const want = base.concat(extra);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, want.length).setValues([want]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  let h = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = want.filter(k => h.indexOf(k) < 0);
  if (missing.length) {
    sh.getRange(1, h.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    h = h.concat(missing);
  }
  const col = {};
  h.forEach((k, i) => { if (k) col[k] = i; });
  return { sh: sh, h: h, col: col };
}

function tz_() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }

function cell_(v) {
  if (v instanceof Date) {
    const hasTime = v.getHours() || v.getMinutes() || v.getSeconds();
    return Utilities.formatDate(v, tz_(), hasTime ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  return v;
}

function rows_(t) {
  const n = t.sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = t.sh.getRange(2, 1, n, t.h.length).getValues();
  return vals.map((r, i) => {
    const o = { _row: i + 2 };
    t.h.forEach((k, j) => { if (k) o[k] = cell_(r[j]); });
    return o;
  });
}

function isDeleted_(o) { return o.deleted === true || String(o.deleted).toUpperCase() === 'TRUE'; }

/** key 열 값으로 행을 찾아 obj 의 필드만 덮어쓰기. 없으면 새 행 추가 */
function upsert_(t, key, obj) {
  const all = rows_(t);
  const hit = all.find(r => String(r[key]) === String(obj[key]));
  const row = t.h.map(k => (k in obj) ? obj[k] : (hit ? (hit[k] === undefined ? '' : hit[k]) : ''));
  if (hit) t.sh.getRange(hit._row, 1, 1, row.length).setValues([row]);
  else t.sh.appendRow(row);
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = parseFloat(v);
  return isNaN(n) ? '' : n;
}

function pct_(before, after) {
  const b = parseFloat(before), a = parseFloat(after);
  if (!b || isNaN(a)) return '';
  return Math.round((b - a) / b * 1000) / 10;
}

function me_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

function log_(ss, action, batchId, detail, who) {
  try {
    table_(ss, 'ActivityLog', LOG_HEADERS, []).sh
      .appendRow([new Date(), who || me_(), action, batchId || '', detail || '']);
  } catch (e) {}
}

function folder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

/* ── 조회 ──────────────────────────────────────────── */

function getBootstrap() {
  const ss = ss_();
  return {
    user: me_(),
    sheetName: ss.getName(),
    sheetUrl: ss.getUrl(),
    lookups: lookups_(ss),
    jobs: jobs_(ss),
    lots: lots_(),
    batches: listBatches_(ss)
  };
}

function listBatches() { return listBatches_(ss_()); }

function lookups_(ss) {
  const out = JSON.parse(JSON.stringify(DEFAULT_LOOKUPS));
  const sh = ss.getSheetByName('Lookups');
  if (!sh || sh.getLastRow() < 2) return out;
  const vals = sh.getDataRange().getValues();
  const head = vals[0].map(String);
  ['material', 'grade', 'furnace'].forEach(k => {
    const i = head.indexOf(k);
    if (i < 0) return;
    const list = vals.slice(1).map(r => String(r[i]).trim()).filter(Boolean);
    if (list.length) out[k] = list;
  });
  return out;
}

function jobs_(ss) {
  const sh = ss.getSheetByName('Jobs');
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const h = vals[0].map(String);
  const c = k => h.indexOf(k);
  const seen = {};
  const out = [];
  for (let i = vals.length - 1; i >= 1 && out.length < 120; i--) {
    const r = vals[i];
    const id = String(r[c('job_id')] || '');
    if (!id || seen[id] || r[c('status')] === '제외') continue;
    seen[id] = true;
    out.push({
      job_id: id,
      at: String(cell_(r[c('registered_at')]) || '').slice(0, 10),
      name: String(r[c('output_name')] || '').replace(/\.gcode$/i, ''),
      material: String(r[c('material')] || ''),
      printer: String(r[c('printer_profile')] || ''),
      status: String(r[c('status')] || ''),
      bbox: ['bbox_x_mm', 'bbox_y_mm', 'bbox_z_mm'].map(k => {
        const v = parseFloat(r[c(k)]);
        return isNaN(v) ? '' : Math.round(v * 10) / 10;
      })
    });
  }
  return out;
}

function lots_() {
  if (!LOT_SHEET_ID) return [];
  try {
    const sh = SpreadsheetApp.openById(LOT_SHEET_ID).getSheetByName('Lots');
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
      .map(r => String(r[0])).filter(Boolean).reverse();
  } catch (e) { return []; }
}

function listBatches_(ss) {
  const sb = rows_(table_(ss, 'SinterBatches', SB_BASE, SB_EXTRA)).filter(b => b.batch_id && !isDeleted_(b));
  const parts = rows_(table_(ss, 'Parts', PARTS_BASE, PARTS_EXTRA)).filter(p => !isDeleted_(p));
  const photos = rows_(table_(ss, 'Photos', PHOTOS_HEADERS, [])).filter(p => !isDeleted_(p));

  return sb.map(b => {
    const ps = parts.filter(p => p.batch_id === b.batch_id);
    const grades = {};
    ps.forEach(p => { if (p.grade) grades[p.grade] = (grades[p.grade] || 0) + 1; });
    const sx = ps.reduce((a, p) => a.concat([parseFloat(p.shrink_x_pct), parseFloat(p.shrink_y_pct)]), []).filter(v => !isNaN(v));
    const sz = ps.map(p => parseFloat(p.shrink_z_pct)).filter(v => !isNaN(v));
    const avg = a => a.length ? Math.round(a.reduce((s, v) => s + v, 0) / a.length * 10) / 10 : '';
    const ph = photos.filter(p => p.batch_id === b.batch_id);
    return {
      batch_id: b.batch_id, date: b.date, status: b.status || '진행중',
      material: b.material, lot_no: b.lot_no, furnace: b.furnace,
      gas: b.atmosphere_gas, flow: b.gas_flow_lpm,
      peak: b.peak_temp_c, peak_hold: b.peak_hold_min, total_min: b.total_min,
      operator: b.operator, result: b.result_summary,
      partCount: ps.length, grades: grades,
      shrinkXY: avg(sx), shrinkZ: avg(sz),
      photoCount: ph.length,
      thumbs: ph.slice(-4).map(p => p.file_id)
    };
  }).sort((a, b) => String(b.date + b.batch_id).localeCompare(String(a.date + a.batch_id)));
}

function getBatch(batchId) {
  const ss = ss_();
  const b = rows_(table_(ss, 'SinterBatches', SB_BASE, SB_EXTRA)).find(r => r.batch_id === batchId && !isDeleted_(r));
  if (!b) throw new Error(batchId + ' 기록을 찾을 수 없습니다.');
  let steps = [];
  try { steps = b.steps_json ? JSON.parse(b.steps_json) : []; } catch (e) {}
  if (!steps.length && b.peak_temp_c !== '') steps = stepsFromFlat_(b);
  delete b._row;
  b.steps = steps;
  const parts = rows_(table_(ss, 'Parts', PARTS_BASE, PARTS_EXTRA))
    .filter(p => p.batch_id === batchId && !isDeleted_(p))
    .map(p => { delete p._row; return p; });
  const photos = rows_(table_(ss, 'Photos', PHOTOS_HEADERS, []))
    .filter(p => p.batch_id === batchId && !isDeleted_(p))
    .map(p => ({ file_id: p.file_id, url: p.url, kind: p.kind, part_id: p.part_id, name: p.file_name, at: p.uploaded_at }));
  return { batch: b, parts: parts, photos: photos };
}

/** 예전 방식(평면 열)으로만 입력된 행을 스텝 표로 복원 */
function stepsFromFlat_(b) {
  const s = [];
  if (b.initial_temp_c !== '') s.push({ label: '탈지', target: b.initial_temp_c, ramp: '', hold: b.initial_hold_min });
  s.push({ label: '소결', target: b.peak_temp_c, ramp: '', hold: b.peak_hold_min });
  if (b.cool_temp_c !== '') s.push({ label: '냉각', target: b.cool_temp_c, ramp: '', hold: b.cool_hold_min });
  return s;
}

/* ── 저장 ──────────────────────────────────────────── */

function nextBatchId_(t, dateStr) {
  const d = dateStr ? new Date(dateStr + 'T12:00:00') : new Date();
  const prefix = 'SB-' + Utilities.formatDate(d, tz_(), 'yyMMdd') + '-';
  let max = 0;
  rows_(t).forEach(r => {
    const v = String(r.batch_id || '');
    if (v.indexOf(prefix) === 0) max = Math.max(max, parseInt(v.slice(prefix.length), 10) || 0);
  });
  return prefix + ('0' + (max + 1)).slice(-2);
}

/** 스텝 표 → 기존 평면 열(initial/ramp_up/peak/ramp_down/cool) 계산 */
function flatFromSteps_(steps) {
  const out = { initial_temp_c: '', initial_hold_min: '', ramp_up_c_min: '', peak_temp_c: '', peak_hold_min: '',
    ramp_down_c_min: '', cool_temp_c: '', cool_hold_min: '', total_min: '' };
  const s = (steps || []).filter(x => x && x.target !== '' && x.target !== undefined && !isNaN(parseFloat(x.target)));
  if (!s.length) return out;
  let total = 0, prev = 25, pk = 0;
  const rates = s.map((x, i) => {
    const t = parseFloat(x.target), ramp = parseFloat(x.ramp) || 0, hold = parseFloat(x.hold) || 0;
    total += ramp + hold;
    const r = ramp > 0 ? (t - prev) / ramp : '';
    prev = t;
    if (t > parseFloat(s[pk].target)) pk = i;
    return r;
  });
  const r1 = v => v === '' ? '' : Math.round(Math.abs(v) * 10) / 10;
  if (pk > 0) { out.initial_temp_c = num_(s[0].target); out.initial_hold_min = num_(s[0].hold); }
  out.peak_temp_c = num_(s[pk].target);
  out.peak_hold_min = num_(s[pk].hold);
  out.ramp_up_c_min = r1(rates[pk]);
  if (pk < s.length - 1) {
    out.ramp_down_c_min = r1(rates[pk + 1]);
    const last = s[s.length - 1];
    out.cool_temp_c = num_(last.target);
    out.cool_hold_min = num_(last.hold);
  }
  out.total_min = total;
  return out;
}

/**
 * payload = { batch: {...필드}, parts: [{part_id?, job_id, part_name, green_*, sint_*, density_g_cm3, grade, notes}],
 *             removedPartIds: [] }
 */
function saveBatch(payload) {
  const ss = ss_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sbT = table_(ss, 'SinterBatches', SB_BASE, SB_EXTRA);
    const ptT = table_(ss, 'Parts', PARTS_BASE, PARTS_EXTRA);
    const b = payload.batch || {};
    const isNew = !b.batch_id;
    const id = isNew ? nextBatchId_(sbT, b.date) : b.batch_id;
    const now = new Date();
    const steps = (b.steps || []).filter(x => x && (x.target !== '' || x.hold !== '' || x.label));
    const flat = flatFromSteps_(steps);
    const who = b.operator || me_();

    const row = {
      batch_id: id,
      date: b.date || Utilities.formatDate(now, tz_(), 'yyyy-MM-dd'),
      furnace: b.furnace || '',
      atmosphere_gas: b.atmosphere_gas || '',
      gas_flow_lpm: num_(b.gas_flow_lpm),
      vacuum: b.vacuum || '',
      furnace_position: b.furnace_position || '',
      operator: b.operator || '',
      notes: b.notes || '',
      status: b.status || '진행중',
      material: b.material || '',
      lot_no: b.lot_no || '',
      steps_json: JSON.stringify(steps),
      cylinder_l: num_(b.cylinder_l),
      cyl_bar_start: num_(b.cyl_bar_start),
      cyl_bar_end: num_(b.cyl_bar_end),
      purge_count: num_(b.purge_count),
      gas_margin: num_(b.gas_margin),
      check_baffle: !!b.check_baffle,
      check_door: !!b.check_door,
      check_program: !!b.check_program,
      result_summary: b.result_summary || '',
      updated_at: now,
      deleted: false
    };
    Object.keys(flat).forEach(k => row[k] = flat[k]);
    if (isNew) { row.created_by = who; row.created_at = now; }
    upsert_(sbT, 'batch_id', row);

    const existing = rows_(ptT).filter(p => p.batch_id === id);
    let seq = existing.reduce((m, p) => Math.max(m, parseInt(String(p.part_id).split('-P')[1], 10) || 0), 0);
    const doneJobs = [];
    (payload.parts || []).forEach(p => {
      const pid = p.part_id || (id + '-P' + ('0' + (++seq)).slice(-2));
      upsert_(ptT, 'part_id', {
        part_id: pid, job_id: p.job_id || '', batch_id: id, part_name: p.part_name || '',
        green_x_mm: num_(p.green_x_mm), green_y_mm: num_(p.green_y_mm), green_z_mm: num_(p.green_z_mm), green_wt_g: num_(p.green_wt_g),
        sint_x_mm: num_(p.sint_x_mm), sint_y_mm: num_(p.sint_y_mm), sint_z_mm: num_(p.sint_z_mm), sint_wt_g: num_(p.sint_wt_g),
        shrink_x_pct: pct_(p.green_x_mm, p.sint_x_mm), shrink_y_pct: pct_(p.green_y_mm, p.sint_y_mm),
        shrink_z_pct: pct_(p.green_z_mm, p.sint_z_mm), wt_loss_pct: pct_(p.green_wt_g, p.sint_wt_g),
        density_g_cm3: num_(p.density_g_cm3), grade: p.grade || '', notes: p.notes || '',
        updated_at: now, deleted: false
      });
      if (p.job_id) doneJobs.push(p.job_id);
    });
    (payload.removedPartIds || []).forEach(pid => {
      if (existing.some(p => p.part_id === pid)) upsert_(ptT, 'part_id', { part_id: pid, deleted: true, updated_at: now });
    });

    if (row.status === '완료' && doneJobs.length) markJobsSintered_(ss, doneJobs);

    log_(ss, isNew ? '등록' : '수정', id,
      [row.material, row.peak_temp_c ? row.peak_temp_c + '℃' : '', (payload.parts || []).length + '개 부품', row.status].filter(Boolean).join(' · '), who);
    return getBatch(id);
  } finally {
    lock.releaseLock();
  }
}

/** Lookups 에 정의된 '소결완료' 상태로 출력 작업을 표시 */
function markJobsSintered_(ss, jobIds) {
  const sh = ss.getSheetByName('Jobs');
  if (!sh || sh.getLastRow() < 2) return;
  const vals = sh.getDataRange().getValues();
  const h = vals[0].map(String);
  const ci = h.indexOf('job_id'), si = h.indexOf('status');
  if (ci < 0 || si < 0) return;
  for (let i = 1; i < vals.length; i++) {
    const st = String(vals[i][si]);
    if (jobIds.indexOf(String(vals[i][ci])) >= 0 && (st === '등록됨' || st === '출력함')) {
      sh.getRange(i + 1, si + 1).setValue('소결완료');
    }
  }
}

function deleteBatch(batchId, who) {
  const ss = ss_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const now = new Date();
    upsert_(table_(ss, 'SinterBatches', SB_BASE, SB_EXTRA), 'batch_id', { batch_id: batchId, deleted: true, updated_at: now });
    const ptT = table_(ss, 'Parts', PARTS_BASE, PARTS_EXTRA);
    rows_(ptT).filter(p => p.batch_id === batchId)
      .forEach(p => upsert_(ptT, 'part_id', { part_id: p.part_id, deleted: true, updated_at: now }));
    log_(ss, '삭제', batchId, '', who);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ── 사진 ──────────────────────────────────────────── */

/** 사진(또는 로그 파일) 1개 업로드. data 는 base64 (data: 접두어 없이) */
function uploadFile(batchId, partId, kind, name, mime, data, who) {
  const ss = ss_();
  const blob = Utilities.newBlob(Utilities.base64Decode(data), mime || 'application/octet-stream',
    batchId + (partId ? '_' + partId.split('-').pop() : '') + '_' + (kind || '기타') + '_' + (name || 'file'));
  const file = folder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  const url = file.getUrl();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    table_(ss, 'Photos', PHOTOS_HEADERS, []).sh.appendRow(
      [batchId, partId || '', kind || '기타', file.getId(), url, name || '', who || me_(), new Date(), false]);
    if (partId) syncPartPhotos_(ss, partId);
    if (kind === '로그') {
      upsert_(table_(ss, 'SinterBatches', SB_BASE, SB_EXTRA), 'batch_id', { batch_id: batchId, log_file_url: url });
    }
  } finally {
    lock.releaseLock();
  }
  log_(ss, '사진', batchId, (partId ? partId + ' · ' : '') + (kind || '') + ' · ' + (name || ''), who);
  return { file_id: file.getId(), url: url, kind: kind, part_id: partId || '', name: name };
}

function deletePhoto(fileId, who) {
  const ss = ss_();
  const t = table_(ss, 'Photos', PHOTOS_HEADERS, []);
  const hit = rows_(t).find(r => r.file_id === fileId);
  if (!hit) return { ok: false };
  t.sh.getRange(hit._row, t.col.deleted + 1).setValue(true);
  if (hit.part_id) syncPartPhotos_(ss, hit.part_id);
  log_(ss, '사진삭제', hit.batch_id, hit.file_name, who);
  return { ok: true };
}

/** Parts.photo_url 에 해당 부품의 사진 링크를 줄바꿈으로 모아 둠 (시트에서 바로 열어볼 수 있게) */
function syncPartPhotos_(ss, partId) {
  const urls = rows_(table_(ss, 'Photos', PHOTOS_HEADERS, []))
    .filter(r => r.part_id === partId && !isDeleted_(r)).map(r => r.url);
  upsert_(table_(ss, 'Parts', PARTS_BASE, PARTS_EXTRA), 'part_id', { part_id: partId, photo_url: urls.join('\n') });
}

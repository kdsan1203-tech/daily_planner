/**
 * GAUSS 소결조건 기록 시스템
 * 구글시트 + Apps Script 웹앱
 *
 * 「가우스 잉크 Lot No 관리 시스템」과 같은 구조로 설계됨:
 *   Runs / Steps / Images / ActivityLog / Config 시트, 소프트 삭제, 작성자 자동 기록
 *   소결런 번호 형식: SR-YYMMDD-NN   (잉크 Lot 형식 GI-CU-260828-01 과 대응)
 *
 * 설치 방법은 같은 폴더의 README.md 참고.
 */

const SS_NAME_SUFFIX = 'GAUSS 소결조건 기록';
const RUNS_SHEET   = 'Runs';
const STEPS_SHEET  = 'Steps';
const IMAGES_SHEET = 'Images';
const LOG_SHEET    = 'ActivityLog';
const CONFIG_SHEET = 'Config';

const RUNS_HEADER = [
  'RunNo','일자','LotNo','소재코드','소재명','샘플수량','세터소재',
  '가스종류','봄베용적(L)','1차압시작(bar)','1차압종료(bar)','퍼징횟수',
  '백필유량(L/min)','상시유량(L/min)',
  '최고온도(℃)','총사이클시간(분)','가스필요량(L)','가스가용량(L)','여유율',
  '차단벽삽입','도어대각체결','프로그램저장확인',
  '결과색상','변형여부','크랙여부','수축률(%)','소결성공여부',
  '특이사항','다음런개선점',
  '작성자','등록시각','수정시각','삭제됨','삭제자','삭제시각'
];

const STEPS_HEADER = ['RunNo','StepNo','구간명','목표온도(℃)','승온시간(분)','승온속도(℃/min)','유지시간(분)','비고'];
const IMAGES_HEADER = ['RunNo','FileId','파일명','구분','업로더','시각','메모'];
const LOG_HEADER = ['시각','직원','이메일','동작','대상RunNo','상세'];
const CONFIG_HEADER = ['키','값'];

const DEFAULT_CONFIG = [
  ['관리자이메일','kdsan1203@gmail.com'],
  ['기본봄베용적(L)','47'],
  ['최소여유율','1.5'],
  ['야간무인최소여유율','2.0'],
  ['비고','여유율 = 가스가용량 ÷ 가스필요량. SOP §6 기준.']
];

/* 소재 목록 - 잉크 Lot No 시스템의 SDS 시트와 같은 코드 체계 */
const MATERIALS = [
  {code:'CU',  name:'Copper (구리)',              gas:'Ar-5%H₂',  minPurge:2, safeTemp:150},
  {code:'SUS', name:'316L Stainless Steel',       gas:'Ar-5%H₂',  minPurge:2, safeTemp:150},
  {code:'TI',  name:'Titanium (티타늄)',           gas:'고순도 Ar', minPurge:3, safeTemp:100},
  {code:'FE',  name:'Iron (철)',                   gas:'Ar-5%H₂',  minPurge:2, safeTemp:150},
  {code:'NI',  name:'Nickel (니켈)',               gas:'Ar-5%H₂',  minPurge:2, safeTemp:150},
  {code:'AL',  name:'Aluminum (알루미늄)',         gas:'고순도 Ar', minPurge:3, safeTemp:150},
  {code:'W',   name:'Tungsten (텅스텐)',           gas:'Ar-5%H₂',  minPurge:3, safeTemp:150},
  {code:'WC',  name:'Tungsten Carbide',           gas:'Ar-5%H₂',  minPurge:3, safeTemp:150},
  {code:'PH',  name:'17-4 PH (SUS)',              gas:'Ar-5%H₂',  minPurge:2, safeTemp:150},
  {code:'NB',  name:'Niobium (니오븀)',            gas:'고순도 Ar', minPurge:3, safeTemp:100}
];

/* ── 웹앱 진입점 ─────────────────────────────────────── */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('소결조건 기록 · GAUSS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ── 시트 준비 ───────────────────────────────────────── */

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#EEEDE9');
  }
  return sh;
}

function ensureSheets_() {
  sheet_(RUNS_SHEET, RUNS_HEADER);
  sheet_(STEPS_SHEET, STEPS_HEADER);
  sheet_(IMAGES_SHEET, IMAGES_HEADER);
  sheet_(LOG_SHEET, LOG_HEADER);
  const cfg = sheet_(CONFIG_SHEET, CONFIG_HEADER);
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, DEFAULT_CONFIG.length, 2).setValues(DEFAULT_CONFIG);
  }
}

/** 최초 1회 실행용 - 시트 5개와 사진 폴더를 만들어 둡니다. */
function setup() {
  ensureSheets_();
  photoFolder_();
  SpreadsheetApp.getActiveSpreadsheet().toast('시트 준비 완료', SS_NAME_SUFFIX, 5);
}

function config_(key) {
  const sh = sheet_(CONFIG_SHEET, CONFIG_HEADER);
  const last = sh.getLastRow();
  if (last < 2) return '';
  const rows = sh.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) if (rows[i][0] === key) return rows[i][1];
  return '';
}

/* ── 사진 폴더 (잉크 Lot 시스템과 같은 명명 규칙) ───────── */

function photoFolder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folderName = 'GAUSS 소결 런 사진 (' + ss.getName() + ')';
  const it = DriveApp.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(folderName);
}

/* ── 사용자 / 로그 ───────────────────────────────────── */

function userEmail_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

function userName_() {
  const email = userEmail_();
  if (!email) return '알 수 없음';
  const admin = String(config_('관리자이메일') || '');
  if (admin.split(',').map(function (s) { return s.trim(); }).indexOf(email) >= 0) return '김대산';
  return email.split('@')[0];
}

function log_(action, runNo, detail) {
  try {
    sheet_(LOG_SHEET, LOG_HEADER).appendRow([
      new Date(), userName_(), userEmail_(), action, runNo || '', detail || ''
    ]);
  } catch (e) { /* 로그 실패가 저장을 막지 않도록 */ }
}

/* ── 런 번호 자동 생성: SR-YYMMDD-NN ─────────────────── */

function nextRunNo_(dateStr) {
  const tz = Session.getScriptTimeZone();
  const d = dateStr ? new Date(dateStr + 'T00:00:00') : new Date();
  const ymd = Utilities.formatDate(d, tz, 'yyMMdd');
  const prefix = 'SR-' + ymd + '-';
  const sh = sheet_(RUNS_SHEET, RUNS_HEADER);
  const last = sh.getLastRow();
  let max = 0;
  if (last >= 2) {
    const col = sh.getRange(2, 1, last - 1, 1).getValues();
    col.forEach(function (r) {
      const v = String(r[0] || '');
      if (v.indexOf(prefix) === 0) {
        const n = parseInt(v.slice(prefix.length), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    });
  }
  return prefix + ('0' + (max + 1)).slice(-2);
}

/* ── 초기 데이터 (화면 로딩 시 1회) ──────────────────── */

function getBootstrap() {
  ensureSheets_();
  return {
    materials: MATERIALS,
    user: userName_(),
    defaultCylinder: Number(config_('기본봄베용적(L)')) || 47,
    minMargin: Number(config_('최소여유율')) || 1.5,
    nightMargin: Number(config_('야간무인최소여유율')) || 2.0,
    suggestedRunNo: nextRunNo_(''),
    knownLots: knownLotNos_()
  };
}

/**
 * 이미 기록된 LotNo 목록 - 입력 자동완성용.
 * 잉크 Lot No 관리 시스템 시트 ID를 Config에 '잉크시트ID'로 넣어두면
 * 그쪽 Lots 시트에서 직접 읽어옵니다. 없으면 이 시트의 과거 입력값만 사용.
 */
function knownLotNos_() {
  const set = {};
  try {
    const inkId = String(config_('잉크시트ID') || '').trim();
    if (inkId) {
      const lots = SpreadsheetApp.openById(inkId).getSheetByName('Lots');
      if (lots && lots.getLastRow() >= 2) {
        lots.getRange(2, 1, lots.getLastRow() - 1, 1).getValues()
          .forEach(function (r) { if (r[0]) set[r[0]] = true; });
      }
    }
  } catch (e) { /* 권한 없거나 시트 없으면 조용히 넘어감 */ }
  try {
    const sh = sheet_(RUNS_SHEET, RUNS_HEADER);
    if (sh.getLastRow() >= 2) {
      sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues()
        .forEach(function (r) { if (r[0]) set[r[0]] = true; });
    }
  } catch (e) {}
  return Object.keys(set).sort().reverse();
}

/* ── 저장 ────────────────────────────────────────────── */

/**
 * payload = {
 *   runNo, date, lotNo, materialCode, materialName, qty, setter,
 *   gasType, cylinderVol, pStart, pEnd, purgeCount, backfillFlow, steadyFlow,
 *   maxTemp, totalMinutes, gasNeed, gasAvail, margin,
 *   baffle, doorTight, programSaved,
 *   resultColor, deformed, cracked, shrinkage, success,
 *   notes, improve,
 *   steps: [{stepNo, label, targetTemp, rampMin, rampRate, holdMin, note}],
 *   photos: [{name, data, kind, memo}]      // kind: '장입' | '결과' | '기타'
 * }
 */
function submitRun(payload) {
  ensureSheets_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const runNo = (payload.runNo && payload.runNo.trim())
      ? payload.runNo.trim()
      : nextRunNo_(payload.date);

    const now = new Date();
    sheet_(RUNS_SHEET, RUNS_HEADER).appendRow([
      runNo,
      payload.date || '',
      payload.lotNo || '',
      payload.materialCode || '',
      payload.materialName || '',
      payload.qty || '',
      payload.setter || '',
      payload.gasType || '',
      payload.cylinderVol || '',
      payload.pStart || '',
      payload.pEnd || '',
      payload.purgeCount || '',
      payload.backfillFlow || '',
      payload.steadyFlow || '',
      payload.maxTemp || '',
      payload.totalMinutes || '',
      payload.gasNeed || '',
      payload.gasAvail || '',
      payload.margin || '',
      payload.baffle ? '확인' : '',
      payload.doorTight ? '확인' : '',
      payload.programSaved ? '확인' : '',
      payload.resultColor || '',
      payload.deformed || '',
      payload.cracked || '',
      payload.shrinkage || '',
      payload.success || '',
      payload.notes || '',
      payload.improve || '',
      userName_(),
      now, '', '', '', ''
    ]);

    const stepsSheet = sheet_(STEPS_SHEET, STEPS_HEADER);
    (payload.steps || []).forEach(function (s, i) {
      stepsSheet.appendRow([
        runNo, s.stepNo || (i + 1), s.label || '',
        s.targetTemp || '', s.rampMin || '', s.rampRate || '', s.holdMin || '', s.note || ''
      ]);
    });

    const saved = savePhotos_(runNo, payload.photos);

    log_('런등록', runNo,
      (payload.materialName || '') + ' · ' + (payload.lotNo || 'Lot 미지정') +
      ' · 최고온 ' + (payload.maxTemp || '-') + '℃ · 사진 ' + saved.length + '장');

    return { ok: true, runNo: runNo, photoCount: saved.length, nextRunNo: nextRunNo_(payload.date) };
  } finally {
    lock.releaseLock();
  }
}

function savePhotos_(runNo, photos) {
  if (!photos || !photos.length) return [];
  const folder = photoFolder_();
  const sh = sheet_(IMAGES_SHEET, IMAGES_HEADER);
  const out = [];
  photos.forEach(function (p) {
    if (!p || !p.data) return;
    const m = /^data:([^;]+);base64,(.+)$/.exec(p.data);
    if (!m) return;
    const fname = runNo + '_' + (p.kind || '기타') + '_' + (p.name || 'photo.jpg');
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], fname));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    sh.appendRow([runNo, file.getId(), p.name || fname, p.kind || '기타', userName_(), new Date(), p.memo || '']);
    out.push(file.getId());
  });
  return out;
}

/** 기존 런에 사진만 추가 */
function addPhotos(runNo, photos) {
  ensureSheets_();
  const saved = savePhotos_(runNo, photos);
  log_('사진등록', runNo, saved.length + '장');
  return { ok: true, count: saved.length };
}

/* ── 조회 ────────────────────────────────────────────── */

function getRecentRuns(limit) {
  ensureSheets_();
  const sh = sheet_(RUNS_SHEET, RUNS_HEADER);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, RUNS_HEADER.length).getValues();
  const tz = Session.getScriptTimeZone();

  const imgs = {};
  const ish = sheet_(IMAGES_SHEET, IMAGES_HEADER);
  if (ish.getLastRow() >= 2) {
    ish.getRange(2, 1, ish.getLastRow() - 1, IMAGES_HEADER.length).getValues()
      .forEach(function (r) {
        if (!imgs[r[0]]) imgs[r[0]] = [];
        imgs[r[0]].push({ fileId: r[1], name: r[2], kind: r[3] });
      });
  }

  return values
    .filter(function (row) { return !row[RUNS_HEADER.indexOf('삭제됨')]; })
    .reverse()
    .slice(0, limit || 30)
    .map(function (row) {
      const o = {};
      RUNS_HEADER.forEach(function (h, i) {
        const v = row[i];
        o[h] = (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm') : v;
      });
      o.photos = imgs[row[0]] || [];
      return o;
    });
}

function getStepsForRun(runNo) {
  const sh = sheet_(STEPS_SHEET, STEPS_HEADER);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, STEPS_HEADER.length).getValues()
    .filter(function (r) { return r[0] === runNo; })
    .map(function (r) {
      return { stepNo: r[1], label: r[2], targetTemp: r[3], rampMin: r[4], rampRate: r[5], holdMin: r[6], note: r[7] };
    });
}

/** 과거 런을 그대로 불러와 새 런의 출발점으로 사용 */
function loadRunForCopy(runNo) {
  const sh = sheet_(RUNS_SHEET, RUNS_HEADER);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const values = sh.getRange(2, 1, last - 1, RUNS_HEADER.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === runNo) {
      const o = {};
      RUNS_HEADER.forEach(function (h, j) { o[h] = values[i][j]; });
      o.steps = getStepsForRun(runNo);
      return o;
    }
  }
  return null;
}

/** 소프트 삭제 (잉크 Lot 시스템과 동일 방식) */
function deleteRun(runNo, reason) {
  const sh = sheet_(RUNS_SHEET, RUNS_HEADER);
  const last = sh.getLastRow();
  if (last < 2) return { ok: false };
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) {
    if (col[i][0] === runNo) {
      const row = i + 2;
      sh.getRange(row, RUNS_HEADER.indexOf('삭제됨') + 1, 1, 3)
        .setValues([[true, userName_(), new Date()]]);
      log_('런삭제', runNo, reason || '');
      return { ok: true };
    }
  }
  return { ok: false };
}

/**
 * 소결조건 기록 프로그램 (GAUSS 소결로)
 * 구글시트 + Apps Script 웹앱
 *
 * 설치 방법은 저장소의 sintering-tracker/README.md 참고.
 */

const RUNS_SHEET  = 'Runs';
const STEPS_SHEET = 'Steps';
const PHOTO_FOLDER_NAME = '소결로_런사진';

const RUNS_HEADER = [
  'RunID','일자','작성자','소재','샘플수량',
  '가스종류','1차압시작(bar)','1차압종료(bar)','퍼징횟수',
  '백필유량(L/min)','상시유량(L/min)',
  '차단벽삽입','도어대각체결',
  '최고온도(℃)','총사이클시간(분)',
  '결과색상','변형여부','크랙여부','수축률(%)',
  '특이사항','장입사진','결과사진','등록시각'
];

const STEPS_HEADER = ['RunID','StepNo','목표온도(℃)','승온시간(분)','유지시간(분)','비고'];

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('소결조건 기록 - GAUSS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getPhotoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

/** base64 데이터 URI 배열을 드라이브에 저장하고 공개 링크 배열을 반환 */
function saveImages_(images, prefix) {
  if (!images || !images.length) return [];
  const folder = getPhotoFolder_();
  const urls = [];
  images.forEach(function (img, idx) {
    if (!img || !img.data) return;
    const match = /^data:(.+);base64,(.+)$/.exec(img.data);
    if (!match) return;
    const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1],
      prefix + '_' + (idx + 1) + '_' + (img.name || 'photo'));
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls.push(file.getUrl());
  });
  return urls;
}

/**
 * 클라이언트에서 전달한 런 데이터를 저장.
 * payload = {
 *   runId, date, author, material, qty,
 *   gasType, pStart, pEnd, purgeCount, backfillFlow, steadyFlow,
 *   baffle, doorTight,
 *   maxTemp, totalMinutes,
 *   resultColor, deformed, cracked, shrinkage,
 *   notes,
 *   steps: [{stepNo, targetTemp, rampMin, holdMin, note}, ...],
 *   loadPhotos: [{name, data}], resultPhotos: [{name, data}]
 * }
 */
function submitRun(payload) {
  const runsSheet = getSheet_(RUNS_SHEET, RUNS_HEADER);
  const stepsSheet = getSheet_(STEPS_SHEET, STEPS_HEADER);

  const runId = payload.runId && payload.runId.trim()
    ? payload.runId.trim()
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');

  const loadPhotoUrls = saveImages_(payload.loadPhotos, runId + '_장입');
  const resultPhotoUrls = saveImages_(payload.resultPhotos, runId + '_결과');

  runsSheet.appendRow([
    runId,
    payload.date || '',
    payload.author || '',
    payload.material || '',
    payload.qty || '',
    payload.gasType || '',
    payload.pStart || '',
    payload.pEnd || '',
    payload.purgeCount || '',
    payload.backfillFlow || '',
    payload.steadyFlow || '',
    payload.baffle ? '확인' : '',
    payload.doorTight ? '확인' : '',
    payload.maxTemp || '',
    payload.totalMinutes || '',
    payload.resultColor || '',
    payload.deformed || '',
    payload.cracked || '',
    payload.shrinkage || '',
    payload.notes || '',
    loadPhotoUrls.join('\n'),
    resultPhotoUrls.join('\n'),
    new Date()
  ]);

  (payload.steps || []).forEach(function (s, i) {
    stepsSheet.appendRow([
      runId,
      s.stepNo || (i + 1),
      s.targetTemp || '',
      s.rampMin || '',
      s.holdMin || '',
      s.note || ''
    ]);
  });

  return { ok: true, runId: runId };
}

/** 최근 런 목록 (요약) 을 반환 - 화면 하단 이력 테이블용 */
function getRecentRuns(limit) {
  const sh = getSheet_(RUNS_SHEET, RUNS_HEADER);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(limit || 20, last - 1);
  const startRow = last - n + 1;
  const values = sh.getRange(startRow, 1, n, RUNS_HEADER.length).getValues();
  return values.reverse().map(function (row) {
    const o = {};
    RUNS_HEADER.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
}

/** 특정 RunID의 스텝 목록 반환 */
function getStepsForRun(runId) {
  const sh = getSheet_(STEPS_SHEET, STEPS_HEADER);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, STEPS_HEADER.length).getValues();
  return values.filter(function (row) { return row[0] === runId; })
    .map(function (row) {
      return { stepNo: row[1], targetTemp: row[2], rampMin: row[3], holdMin: row[4], note: row[5] };
    });
}

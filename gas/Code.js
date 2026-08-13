/** 72구절 암송 시험 — 웹앱 서버 (구글시트 컨테이너 바인딩)
 *  doGet: index.html을 그대로 서빙하되 URL 파라미터(?s= ?v= ?p= ?t=)를 window.__QS로 주입
 *  submitResult: 리더가 시험을 제출하면 '성적' 시트에 한 줄씩 기록
 */
var SHEET_NAME = '성적';

/** 최초 1회 실행용: 권한 승인 트리거 (실행하면 시트 이름을 읽기만 함) */
function authorize() {
  var name = SpreadsheetApp.getActive().getName();
  Logger.log('권한 승인 완료: ' + name);
}

function doGet(e) {
  var qs = buildQs_(e && e.parameter);
  var html = HtmlService.createHtmlOutputFromFile('index').getContent();
  return HtmlService.createHtmlOutput('<script>window.__QS = ' + JSON.stringify(qs) + ';</script>' + html)
    .setTitle('72구절 암송 시험')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildQs_(p) {
  if (!p) return '';
  var keys = ['s', 'v', 'p', 't'];
  var out = [];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (p[k]) out.push(k + '=' + encodeURIComponent(String(p[k])));
  }
  return out.join('&');
}

/** 외부 POST 진입점: JSON 페이로드로 성적 제출(예비 경로) 또는 자가진단 */
function doPost(e) {
  try {
    var r = JSON.parse(e.postData.contents);
    if (r && r.__selftest) return jsonOut_(selfTest_());
    return jsonOut_(submitResult(r));
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 시트에 테스트 행을 쓰고 → 읽어 확인하고 → 지운다. 흔적 없이 왕복 검증 */
function selfTest_() {
  submitResult({ name: '__selftest', week: 99, theme: 't', ref1: 'r1', s1: 1, ref2: 'r2', s2: 2, avg: 2, pass: false, stamp: 'test', code: 'TEST-TEST' });
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  var last = sh.getLastRow();
  var row = sh.getRange(last, 1, 1, 11).getValues()[0];
  var ok = row[1] === '__selftest';
  if (ok) sh.deleteRow(last);
  return { ok: ok, rowsAfter: sh.getLastRow(), sheetName: SpreadsheetApp.getActive().getName() };
}

function submitResult(r) {
  if (!r || !r.name || !r.week) throw new Error('잘못된 제출 데이터');
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(['제출시각', '이름', '주차', '주제', '구절1', '점수1', '구절2', '점수2', '평균', '결과', '확인코드']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      String(r.stamp || ''),
      String(r.name).slice(0, 20),
      Number(r.week) || 0,
      String(r.theme || ''),
      String(r.ref1 || ''),
      Number(r.s1) || 0,
      String(r.ref2 || ''),
      Number(r.s2) || 0,
      Number(r.avg) || 0,
      r.pass ? '합격' : '불합격',
      String(r.code || '')
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

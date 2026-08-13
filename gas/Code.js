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
  var execUrl = ScriptApp.getService().getUrl(); // 샌드박스 iframe에서는 location으로 알 수 없는 실제 /exec 주소
  return HtmlService.createHtmlOutput('<script>window.__QS = ' + JSON.stringify(qs) + ';window.__URL = ' + JSON.stringify(execUrl) + ';</script>' + html)
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

/** 외부 POST 진입점: JSON 페이로드로 성적 제출(예비 경로), 자가진단, 시트 단장 */
function doPost(e) {
  try {
    var r = JSON.parse(e.postData.contents);
    if (r && r.__selftest) return jsonOut_(selfTest_());
    if (r && r.__beautify) return jsonOut_(beautify_());
    return jsonOut_(submitResult(r));
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/** '성적' 시트가 없으면 헤더와 함께 생성 */
function ensureSheet_(ss) {
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['제출시각', '이름', '주차', '주제', '구절1', '점수1', '구절2', '점수2', '평균', '결과', '확인코드']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 성적 시트 단장: 앱과 같은 딥그린 톤. 여러 번 실행해도 안전(멱등) */
function beautify_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ensureSheet_(ss);
  var maxRows = Math.max(sh.getMaxRows(), 100);
  sh.setTabColor('#1E5748');
  sh.setHiddenGridlines(true);
  /* 헤더 */
  sh.getRange(1, 1, 1, 11)
    .setBackground('#1E5748').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(1, 38);
  sh.setFrozenRows(1);
  /* 열 너비 */
  var widths = [150, 90, 64, 170, 140, 72, 140, 72, 72, 84, 110];
  for (var i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);
  /* 본문 정렬: 숫자·결과·코드는 가운데 */
  var body = sh.getRange(2, 1, maxRows - 1, 11);
  body.setVerticalAlignment('middle');
  [3, 6, 8, 9, 10, 11].forEach(function (c) { sh.getRange(2, c, maxRows - 1, 1).setHorizontalAlignment('center'); });
  /* 줄무늬 배경 (기존 밴딩 제거 후 재적용) */
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.getRange(1, 1, maxRows, 11)
    .applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, true, false)
    .setHeaderRowColor('#1E5748').setFirstRowColor('#FFFFFF').setSecondRowColor('#F2F7F3');
  /* 조건부 서식: 합격/불합격 배지, 90점 기준 점수 색 */
  var resRange = sh.getRange(2, 10, maxRows - 1, 1);
  var scoreRanges = [sh.getRange(2, 6, maxRows - 1, 1), sh.getRange(2, 8, maxRows - 1, 1), sh.getRange(2, 9, maxRows - 1, 1)];
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('합격').setBackground('#DFF0E3').setFontColor('#1E6E3C').setBold(true).setRanges([resRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('불합격').setBackground('#F6E3DE').setFontColor('#A93D2B').setBold(true).setRanges([resRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(90).setFontColor('#1E6E3C').setBold(true).setRanges(scoreRanges).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(90).setFontColor('#A93D2B').setRanges(scoreRanges).build()
  ]);
  /* 헤더 필터 */
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, maxRows, 11).createFilter();
  return { ok: true, styled: true };
}

function jsonOut_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 시트에 테스트 행을 쓰고 → 읽어 확인하고 → 지운다. 흔적 없이 왕복 + 수식주입 방어 검증
 *  이름을 '=1+1'로 제출: 방어가 되면 문자열 그대로, 뚫리면 수식으로 계산된 2가 읽힌다 */
function selfTest_() {
  submitResult({ name: '=1+1', week: 99, theme: 't', ref1: 'r1', s1: 1, ref2: 'r2', s2: 2, avg: 2, pass: false, stamp: 'test', code: 'TEST-TEST' });
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  var last = sh.getLastRow();
  var row = sh.getRange(last, 1, 1, 11).getValues()[0];
  var ok = row[1] === '=1+1' && row[10] === 'TEST-TEST';
  if (row[10] === 'TEST-TEST') sh.deleteRow(last); // 자기 행만 삭제
  return { ok: ok, formulaSafe: row[1] === '=1+1', rowsAfter: sh.getLastRow() };
}

/** 시트 수식 주입 방지: =,+,-,@ 등으로 시작하는 문자열을 텍스트로 강제 */
function deFormula_(v) {
  v = String(v);
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}

function submitResult(r) {
  if (!r || !r.name || !r.week) throw new Error('잘못된 제출 데이터');
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    var sh = ensureSheet_(SpreadsheetApp.getActive());
    sh.appendRow([
      deFormula_(String(r.stamp || '').slice(0, 30)),
      deFormula_(String(r.name).slice(0, 20)),
      Number(r.week) || 0,
      deFormula_(String(r.theme || '').slice(0, 40)),
      deFormula_(String(r.ref1 || '').slice(0, 40)),
      Number(r.s1) || 0,
      deFormula_(String(r.ref2 || '').slice(0, 40)),
      Number(r.s2) || 0,
      Number(r.avg) || 0,
      r.pass ? '합격' : '불합격',
      deFormula_(String(r.code || '').slice(0, 20))
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

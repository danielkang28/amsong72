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
    if (r && r.__dashboard) return jsonOut_(dashboard_());
    if (r && r.__dashcheck) return jsonOut_(dashCheck_());
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

/** '개요' 시트 생성(있으면 재생성): 리더별 응시 횟수·평균·합격률·최근 제출 + 셀 안 성장 스파크라인 + 주차별 전체 평균 차트.
 *  전부 '성적' 시트를 참조하는 수식이라 제출이 쌓이면 자동 갱신됨 */
function dashboard_() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss);
  var old = ss.getSheetByName('개요');
  if (old) ss.deleteSheet(old);
  var sh = ss.insertSheet('개요', 0);
  var SLOTS = 30;      // 리더 최대 인원
  var S = "'성적'!";   // 성적 시트 열: A제출시각 B이름 C주차 I평균 J결과
  sh.setTabColor('#1E5748');
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(1);

  sh.getRange('A1:G1').setValues([['이름', '응시 횟수', '평균 점수', '합격 / 응시', '합격률', '최근 제출', '성장 그래프 (주차별 점수)']])
    .setBackground('#1E5748').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(1, 38);

  sh.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + S + 'B2:B,' + S + 'B2:B<>""))),"")');
  var rows = [];
  for (var r = 2; r < 2 + SLOTS; r++) {
    var A = '$A' + r, B = 'B' + r;
    rows.push([
      '=IF(' + A + '="","",COUNTIF(' + S + '$B:$B,' + A + '))',
      '=IF(' + A + '="","",ROUND(AVERAGEIF(' + S + '$B:$B,' + A + ',' + S + '$I:$I),1))',
      '=IF(' + A + '="","",COUNTIFS(' + S + '$B:$B,' + A + ',' + S + '$J:$J,"합격")&" / "&' + B + ')',
      '=IF(' + A + '="","",ROUND(100*COUNTIFS(' + S + '$B:$B,' + A + ',' + S + '$J:$J,"합격")/' + B + ')&"%")',
      '=IF(' + A + '="","",INDEX(SORT(FILTER(' + S + '$A$2:$A,' + S + '$B$2:$B=' + A + '),1,FALSE),1))',
      '=IF(' + A + '="","",IFERROR(SPARKLINE(SORT(FILTER({' + S + '$C$2:$C,' + S + '$I$2:$I},' + S + '$B$2:$B=' + A + '),1,TRUE),{"charttype","line";"linewidth",2;"color","#1E5748";"ymin",0;"ymax",100}),""))'
    ]);
  }
  sh.getRange(2, 2, SLOTS, 6).setFormulas(rows);

  [96, 72, 84, 96, 72, 150, 260].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setRowHeights(2, SLOTS, 34);
  sh.getRange(2, 1, SLOTS, 7).setVerticalAlignment('middle');
  sh.getRange(2, 1, SLOTS, 1).setFontWeight('bold');
  sh.getRange(2, 2, SLOTS, 5).setHorizontalAlignment('center');
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.getRange(1, 1, SLOTS + 1, 7).applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, true, false)
    .setHeaderRowColor('#1E5748').setFirstRowColor('#FFFFFF').setSecondRowColor('#F2F7F3');
  var avgR = sh.getRange(2, 3, SLOTS, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(90).setFontColor('#1E6E3C').setBold(true).setRanges([avgR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(90).setFontColor('#A93D2B').setRanges([avgR]).build()
  ]);

  /* 주차별 전체 평균 (차트 데이터) */
  sh.getRange('I1:J1').setValues([['주차', '전체 평균']]).setFontWeight('bold').setFontColor('#5F6A64');
  sh.getRange('I2').setFormula('=SEQUENCE(36)');
  var wk = [];
  for (var w = 2; w <= 37; w++) {
    wk.push(['=IF(COUNTIF(' + S + '$C:$C,$I' + w + ')=0,"",ROUND(AVERAGEIF(' + S + '$C:$C,$I' + w + ',' + S + '$I:$I),1))']);
  }
  sh.getRange(2, 10, 36, 1).setFormulas(wk);
  sh.setColumnWidth(9, 50);
  sh.setColumnWidth(10, 76);
  sh.getRange(2, 9, 36, 2).setHorizontalAlignment('center').setFontColor('#5F6A64');
  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange('I1:J37'))
    .setPosition(2, 12, 0, 0)
    .setOption('title', '주차별 전체 평균 점수')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#1E5748'])
    .setOption('interpolateNulls', true)
    .setOption('pointSize', 5)
    .setOption('vAxis', { minValue: 0, maxValue: 100, title: '점수' })
    .setOption('hAxis', { title: '주차' })
    .setOption('width', 560).setOption('height', 320)
    .build());
  return { ok: true, dashboard: true };
}

/** 개요 시트 점검: 테스트 행 1건을 넣고 개요 수식이 오류 없이 계산되는지 읽은 뒤 테스트 행을 지운다 */
function dashCheck_() {
  var ss = SpreadsheetApp.getActive();
  submitResult({ name: '__dashcheck', week: 99, theme: 't', ref1: 'r1', s1: 95, ref2: 'r2', s2: 85, avg: 90, pass: true, stamp: 'test', code: 'TEST-TEST' });
  SpreadsheetApp.flush();
  var d = ss.getSheetByName('개요');
  var errors = [], row2 = null;
  if (d) {
    var vals = d.getRange(1, 1, 40, 10).getDisplayValues();
    row2 = vals[1];
    vals.forEach(function (row, i) { row.forEach(function (c, j) { if (String(c).charAt(0) === '#') errors.push({ r: i + 1, c: j + 1, v: c }); }); });
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  var last = sh.getLastRow();
  if (sh.getRange(last, 11).getValue() === 'TEST-TEST') sh.deleteRow(last); // 자기 행만 삭제
  return { ok: !!d && errors.length === 0, hasDashboard: !!d, errors: errors, row2: row2 };
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

/** 72구절 암송 시험 — 웹앱 서버 (구글시트 컨테이너 바인딩)
 *  doGet: index.html을 그대로 서빙하되 URL 파라미터(?s= ?v= ?p= ?t=)를 window.__QS로 주입
 *  submitResult: 리더가 시험을 제출하면 '성적' 시트에 한 줄씩 기록 (재응시는 같은 (이름, 주차, 1주차주일) 중 최신만 '최종')
 *  doPost 유틸: {"__selftest":true} {"__beautify":true} {"__dashboard":true} {"__dashcheck":true} {"__migrate":true,"start":"YYYY-MM-DD"}
 *  배포는 저장소 루트의 deploy.ps1 (clasp push → version → redeploy, 웹앱 URL 유지)
 */
var SHEET_NAME = '성적';
var DASH_NAME = '개요';
var TOTAL_WEEKS = 36;    // 과정(사이클) 길이. 주차 번호는 누적(37, 38, …)으로 저장
var PASS = 90;           // 합격 기준 점수(시트 서식용)
var CHART_WEEKS = 108;   // 개요 차트 데이터 행 수(3사이클)
var HEADERS = ['제출시각', '이름', '주차', '주제', '구절1', '점수1', '구절2', '점수2', '평균', '결과', '확인코드', '대표', '1주차주일', '버전'];
var COL = { stamp: 1, name: 2, week: 3, theme: 4, ref1: 5, s1: 6, ref2: 7, s2: 8, avg: 9, result: 10, code: 11, rep: 12, start: 13, ver: 14 };
var NCOL = HEADERS.length;
var FINAL = '최종';
var TEST_CODE = 'TEST-TEST';
var DUP_SCAN = 300;      // 중복 확인코드 검사 범위(최근 행 수)

/** 열 번호 → 열 문자 (1 → A) */
function colL_(c) { return String.fromCharCode(64 + c); }

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

/** 외부 POST 진입점: JSON 페이로드로 성적 제출(예비 경로), 자가진단, 시트 단장, 마이그레이션 */
function doPost(e) {
  try {
    var r = JSON.parse(e.postData.contents);
    if (r && r.__selftest) return jsonOut_(selfTest_());
    if (r && r.__beautify) return jsonOut_(beautify_());
    if (r && r.__dashboard) return jsonOut_(dashboard_());
    if (r && r.__dashcheck) return jsonOut_(dashCheck_());
    if (r && r.__migrate) return jsonOut_(migrate_(r.start));
    return jsonOut_(submitResult(r));
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/** '성적' 시트가 없으면 헤더와 함께 생성. 있으면 1행 1..14열 중 빈 헤더 칸만 채운다(마이그레이션) */
function ensureSheet_(ss) {
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    if (sh.getMaxColumns() < NCOL) sh.insertColumnsAfter(sh.getMaxColumns(), NCOL - sh.getMaxColumns());
    sh.getRange(1, 1, 1, NCOL).setValues([HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }
  if (sh.getMaxColumns() < NCOL) sh.insertColumnsAfter(sh.getMaxColumns(), NCOL - sh.getMaxColumns());
  /* getLastColumn()과 무관하게 항상 1행 1..14를 읽어 빈 칸만 보충(중간 헤더가 지워진 경우도 복구) */
  var head = sh.getRange(1, 1, 1, NCOL).getValues()[0];
  var changed = false;
  for (var i = 0; i < NCOL; i++) {
    if (head[i] === '' || head[i] === null) { head[i] = HEADERS[i]; changed = true; }
  }
  if (changed) {
    sh.getRange(1, 1, 1, NCOL).setValues([head]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 스프레드시트 시간대: 실행당 1회만 조회(행마다 서비스 호출하면 제출이 느려짐) */
var TZ_ = null;
function tz_() {
  if (!TZ_) TZ_ = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return TZ_;
}

/** 셀 값(문자열 또는 시트가 날짜로 바꾼 Date) → 'YYYY-MM-DD' 비교용 문자열 */
function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

/** (이름, 주차, 1주차주일) 그룹 키 */
function groupKey_(name, week, start) {
  return String(name) + '\u0001' + Number(week) + '\u0001' + ymd_(start);
}

/** 현재 텀 = '성적' M열에서 가장 늦은 YYYY-MM-DD ('test' 같은 값은 제외). 없으면 '' */
function currentStart_(sh) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return '';
  var vals = sh.getRange(2, COL.start, n, 1).getValues();
  var best = '';
  for (var i = 0; i < n; i++) {
    var s = ymd_(vals[i][0]);
    if (/^\d{4}-\d{2}-\d{2}$/.test(s) && s > best) best = s;
  }
  return best;
}

/** 성적 시트 단장: 앱과 같은 딥그린 톤. 여러 번 실행해도 안전(멱등) */
function beautify_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ensureSheet_(ss);
  var LN = colL_(NCOL);
  var colRange = function (c) { return sh.getRange(colL_(c) + '2:' + colL_(c)); }; // 2행부터 열 끝까지(열 전체)
  sh.setTabColor('#1E5748');
  sh.setHiddenGridlines(true);
  /* 헤더 */
  sh.getRange(1, 1, 1, NCOL)
    .setBackground('#1E5748').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(1, 38);
  sh.setFrozenRows(1);
  /* 열 너비 (HEADERS 순서) */
  var widths = [150, 90, 64, 170, 140, 72, 140, 72, 72, 84, 110, 64, 110, 100];
  for (var i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);
  /* 본문 정렬: 숫자·결과·코드·대표·주일·버전은 가운데 */
  sh.getRange('A2:' + LN).setVerticalAlignment('middle');
  [COL.week, COL.s1, COL.s2, COL.avg, COL.result, COL.code, COL.rep, COL.start, COL.ver].forEach(function (c) { colRange(c).setHorizontalAlignment('center'); });
  colRange(COL.start).setNumberFormat('yyyy-mm-dd');
  /* 줄무늬 배경 (기존 밴딩 제거 후 열 전체에 재적용) */
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.getRange('A:' + LN)
    .applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, true, false)
    .setHeaderRowColor('#1E5748').setFirstRowColor('#FFFFFF').setSecondRowColor('#F2F7F3');
  /* 조건부 서식: 합격/불합격 배지, PASS 기준 점수 색, '최종' 배지 */
  var resRange = colRange(COL.result);
  var repRange = colRange(COL.rep);
  var scoreRanges = [colRange(COL.s1), colRange(COL.s2), colRange(COL.avg)];
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('합격').setBackground('#DFF0E3').setFontColor('#1E6E3C').setBold(true).setRanges([resRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('불합격').setBackground('#F6E3DE').setFontColor('#A93D2B').setBold(true).setRanges([resRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(PASS).setFontColor('#1E6E3C').setBold(true).setRanges(scoreRanges).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(PASS).setFontColor('#A93D2B').setRanges(scoreRanges).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(FINAL).setBackground('#E6F4E9').setFontColor('#1E6E3C').setBold(true).setRanges([repRange]).build()
  ]);
  /* 헤더 필터 (A:N 전체 열) */
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange('A:' + LN).createFilter();
  return { ok: true, styled: true };
}

/** '개요' 시트(비파괴): 없으면 만들고, 있으면 A1:J60만 지우고 다시 쓴다(L열 이후 메모·사용자 서식은 유지).
 *  현재 텀(성적 M열 최신 1주차주일)의 '최종' 행만 집계. 전부 '성적' 시트를 참조하는 수식이라 제출이 쌓이면 자동 갱신됨 */
function dashboard_() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss);
  var sh = ss.getSheetByName(DASH_NAME);
  if (!sh) sh = ss.insertSheet(DASH_NAME, 0);
  var SLOTS = 40;          // 리더 최대 인원 (표 2~41행)
  var LAST = SLOTS + 1;    // 표 마지막 행
  var AREA_ROWS = 60, AREA_COLS = 10; // A1:J60 재생성 영역
  var S = "'" + SHEET_NAME + "'!";
  var full = function (k) { var L = colL_(COL[k]); return S + '$' + L + ':$' + L; };        // '성적'!$B:$B
  var body = function (k) { var L = colL_(COL[k]); return S + '$' + L + '$2:$' + L; };      // '성적'!$B$2:$B
  var rel = function (k) { var L = colL_(COL[k]); return S + L + '2:' + L; };               // '성적'!B2:B
  var Q = '"' + FINAL + '"';

  /* 1) 재생성 영역 정리: 내용·서식, 영역에 걸친 밴딩·조건부서식만 제거 */
  var area = sh.getRange(1, 1, AREA_ROWS, AREA_COLS);
  area.getBandings().forEach(function (b) { b.remove(); });
  area.clearContent();
  area.clearFormat();
  var keepRules = sh.getConditionalFormatRules().filter(function (rule) {
    return rule.getRanges().every(function (rg) {
      return rg.getRow() > AREA_ROWS || rg.getColumn() > AREA_COLS;
    });
  });

  sh.setTabColor('#1E5748');
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(1);

  /* 2) 설정 블록 (50~53행) */
  sh.getRange('A50:A53').setValues([['현재 텀 1주차 주일'], ['이번 주차'], ['이번 주 표기'], ['미제출 명단(톡방용)']])
    .setFontWeight('bold').setFontColor('#5F6A64');
  sh.getRange('B50:B53').setFormulas([
    ['=IFERROR(INDEX(SORT(UNIQUE(FILTER(' + rel('start') + ',' + rel('start') + '<>"",ISNUMBER(' + rel('start') + '))),1,FALSE),1),"")'], // 날짜(숫자)만 — 'test' 같은 문자열 start 제외
    ['=IFERROR(MAX(FILTER(' + rel('week') + ',' + rel('start') + '=B50,' + rel('rep') + '=' + Q + ')),0)'],
    ['=IF(B51>' + TOTAL_WEEKS + ',INT((B51-1)/' + TOTAL_WEEKS + ')+1&"회차 "&MOD(B51-1,' + TOTAL_WEEKS + ')+1&"주차",B51&"주차")'],
    ['=IFERROR("["&B52&" 미제출] "&TEXTJOIN(", ",TRUE,FILTER(A2:A' + LAST + ',H2:H' + LAST + '="미제출")),"미제출 없음")']
  ]);
  sh.getRange('B50').setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('left');
  sh.getRange('B51').setHorizontalAlignment('left');
  sh.getRange('A' + AREA_ROWS).setValue('※ A1:J109 영역(표·설정·차트 데이터)은 재생성 시 덮어씁니다. 차트는 L2에 놓입니다. 메모는 L열 아래쪽(L25 이하)이나 다른 시트에 적어 주세요.')
    .setFontColor('#9AA39E').setFontStyle('italic');

  /* 3) 표 (1~41행) */
  sh.getRange('A1:H1').setValues([['이름', '응시 주차', '평균 점수', '합격', '합격률', '최근 제출', '성장 그래프(주차별)', '이번 주']])
    .setBackground('#1E5748').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(1, 38);
  sh.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + rel('name') + ',' + rel('name') + '<>"",' + rel('start') + '=$B$50))),"")');
  var rows = [];
  for (var r = 2; r <= LAST; r++) {
    var A = '$A' + r, B = '$B' + r;
    var cond = full('name') + ',' + A + ',' + full('start') + ',$B$50,' + full('rep') + ',' + Q;       // 이름·현재 텀·최종
    var passCnt = 'COUNTIFS(' + cond + ',' + full('result') + ',"합격")';
    var g = function (expr) { return '=IF(' + A + '="","",IFERROR(' + expr + ',""))'; };
    rows.push([
      g('COUNTIFS(' + cond + ')'),
      g('ROUND(AVERAGEIFS(' + full('avg') + ',' + cond + '),1)'),
      g(passCnt + '&" / "&' + B),
      g('ROUND(100*' + passCnt + '/' + B + ')&"%"'),
      g('INDEX(SORT(FILTER(' + body('stamp') + ',' + body('name') + '=' + A + ',' + body('start') + '=$B$50),1,FALSE),1)'),
      g('SPARKLINE(SORT(FILTER({' + body('week') + ',' + body('avg') + '},' + body('name') + '=' + A + ',' + body('rep') + '=' + Q + ',' + body('start') + '=$B$50),1,TRUE),{"charttype","line";"linewidth",2;"color","#1E5748";"ymin",0;"ymax",100})'),
      '=IF(' + A + '="","",IFERROR(INDEX(FILTER(' + body('result') + ',' + body('name') + '=' + A + ',' + body('week') + '=$B$51,' + body('rep') + '=' + Q + ',' + body('start') + '=$B$50),1),"미제출"))'
    ]);
  }
  sh.getRange(2, 2, SLOTS, 7).setFormulas(rows);

  /* 4) 주차별 전체 평균 보조열 (차트 데이터) */
  sh.getRange('I1:J1').setValues([['주차', '전체 평균']]).setFontWeight('bold').setFontColor('#5F6A64');
  sh.getRange('I2').setFormula('=SEQUENCE(MAX(' + TOTAL_WEEKS + ',$B$51))');
  var wk = [];
  for (var w = 2; w <= 1 + CHART_WEEKS; w++) {
    wk.push(['=IF($I' + w + '="","",IFERROR(ROUND(AVERAGEIFS(' + full('avg') + ',' + full('week') + ',$I' + w + ',' + full('start') + ',$B$50,' + full('rep') + ',' + Q + '),1),""))']);
  }
  sh.getRange(2, 10, CHART_WEEKS, 1).setFormulas(wk);
  sh.getRange(2, 9, CHART_WEEKS, 2).setHorizontalAlignment('center').setFontColor('#5F6A64');

  /* 5) 서식 */
  [130, 72, 84, 84, 72, 150, 260, 84, 50, 76].forEach(function (wd, i) { sh.setColumnWidth(i + 1, wd); });
  sh.setRowHeights(2, SLOTS, 34);
  sh.getRange(2, 1, SLOTS, 8).setVerticalAlignment('middle');
  sh.getRange(2, 1, SLOTS, 1).setFontWeight('bold');
  sh.getRange(2, 2, SLOTS, 5).setHorizontalAlignment('center');
  sh.getRange(2, 8, SLOTS, 1).setHorizontalAlignment('center');
  sh.getRange(1, 1, LAST, 8).applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, true, false)
    .setHeaderRowColor('#1E5748').setFirstRowColor('#FFFFFF').setSecondRowColor('#F2F7F3');
  var avgR = sh.getRange(2, 3, SLOTS, 1);
  var nowR = sh.getRange(2, 8, SLOTS, 1);
  sh.setConditionalFormatRules(keepRules.concat([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(PASS).setFontColor('#1E6E3C').setBold(true).setRanges([avgR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(PASS).setFontColor('#A93D2B').setRanges([avgR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('합격').setFontColor('#1E6E3C').setBold(true).setRanges([nowR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('불합격').setFontColor('#D9822B').setBold(true).setRanges([nowR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('미제출').setFontColor('#9AA39E').setRanges([nowR]).build()
  ]));

  /* 6) 차트: 있으면 범위만 갱신(위치·사용자 수정 유지), 없으면 삽입 */
  var chartRange = sh.getRange(1, 9, 1 + CHART_WEEKS, 2); // I1:J(1+CHART_WEEKS)
  var charts = sh.getCharts();
  if (charts.length) {
    sh.updateChart(charts[0].modify().clearRanges().addRange(chartRange).build());
  } else {
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(chartRange)
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
  }
  return { ok: true, dashboard: true };
}

/** 개요 시트 점검: 현재 텀으로 테스트 행 1건을 넣고 개요 A1:J60 수식이 오류 없이 계산되는지 읽은 뒤 테스트 행을 지운다 */
function dashCheck_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ensureSheet_(ss);
  var d = null, errors = [], row2 = null;
  try {
    submitResult({ name: '__dashcheck', week: 99, theme: 't', ref1: 'r1', s1: 95, ref2: 'r2', s2: 85, avg: 90, pass: true, stamp: 'test', code: TEST_CODE, start: currentStart_(sh), ver: 'dashcheck' });
    SpreadsheetApp.flush();
    d = ss.getSheetByName(DASH_NAME);
    if (d) {
      var vals = d.getRange(1, 1, 60, 10).getDisplayValues();
      row2 = vals[1];
      vals.forEach(function (row, i) { row.forEach(function (c, j) { if (String(c).charAt(0) === '#') errors.push({ r: i + 1, c: j + 1, v: c }); }); });
    }
  } finally {
    deleteTestRows_(sh); // 예외가 나도 자기 행(K=TEST-TEST)만 반드시 삭제
  }
  return { ok: !!d && errors.length === 0, hasDashboard: !!d, errors: errors, row2: row2 };
}

function jsonOut_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 확인코드가 TEST-TEST인 행(자가진단·개요점검 행)만 아래에서부터 삭제. 삭제 수 반환 */
function deleteTestRows_(sh) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  var codes = sh.getRange(2, COL.code, n, 1).getValues();
  var del = 0;
  for (var i = n - 1; i >= 0; i--) {
    if (String(codes[i][0]) === TEST_CODE) { sh.deleteRow(i + 2); del++; }
  }
  return del;
}

/** 시트에 테스트 행을 쓰고 → 같은 payload를 한 번 더 보내 중복 차단을 확인 → 읽어 확인하고 → 지운다.
 *  이름을 '=1+1'로 제출: 방어가 되면 문자열 그대로, 뚫리면 수식으로 계산된 2가 읽힌다 */
function selfTest_() {
  var p = { name: '=1+1', week: 99, theme: 't', ref1: 'r1', s1: 1, ref2: 'r2', s2: 2, avg: 2, pass: false, stamp: 'test', code: TEST_CODE, start: 'test', ver: 'selftest' };
  var sh = ensureSheet_(SpreadsheetApp.getActive());
  var formulaSafe = false, shapeOk = false, dupDetected = false;
  try {
    var r1 = submitResult(p);
    var r2 = submitResult(p);
    var last = sh.getLastRow();
    var row = sh.getRange(last, 1, 1, NCOL).getValues()[0];
    var mine = row[COL.code - 1] === TEST_CODE;
    formulaSafe = mine && row[COL.name - 1] === '=1+1';
    shapeOk = mine && row[COL.rep - 1] === FINAL && row[COL.start - 1] === 'test' && row[COL.ver - 1] === 'selftest';
    dupDetected = !(r1 && r1.dup) && !!(r2 && r2.dup);
  } finally {
    deleteTestRows_(sh); // 예외가 나도 자기 행(K=TEST-TEST)만 반드시 삭제
  }
  return { ok: formulaSafe && dupDetected && shapeOk, formulaSafe: formulaSafe, dupDetected: dupDetected, rowsAfter: sh.getLastRow() };
}

/** 시트 수식 주입 방지: =,+,-,@ 등으로 시작하는 문자열을 텍스트로 강제 */
function deFormula_(v) {
  v = String(v);
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}

/** 0..100 범위의 숫자(또는 숫자 문자열)인지 */
function isScore_(v) {
  if (typeof v === 'number') return isFinite(v) && v >= 0 && v <= 100;
  if (typeof v === 'string' && v.trim() !== '') { var n = Number(v); return isFinite(n) && n >= 0 && n <= 100; }
  return false;
}

function submitResult(r) {
  var bad = function () { return new Error('잘못된 제출 데이터'); };
  if (!r || typeof r !== 'object') throw bad();
  var inName = String(r.name === null || r.name === undefined ? '' : r.name).trim().slice(0, 20);
  var week = Number(r.week);
  if (!inName || !isFinite(week) || Math.floor(week) !== week || week < 1) throw bad();
  if (!isScore_(r.s1) || !isScore_(r.s2) || !isScore_(r.avg)) throw bad();
  var code = String(r.code || '').slice(0, 20);
  var start = String(r.start || '').slice(0, 20);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('잠시 후 다시 시도해 주세요');
  try {
    var sh = ensureSheet_(SpreadsheetApp.getActive());
    var n = sh.getLastRow() - 1;
    /* B..M 한 번에 읽기: 이름 정규화 · 중복 확인코드 · '최종' 갱신에 공통 사용 */
    var data = n > 0 ? sh.getRange(2, COL.name, n, COL.start - COL.name + 1).getValues() : [];
    var at = function (row, k) { return row[COL[k] - COL.name]; };

    /* 중복: 최근 DUP_SCAN 행에 같은 확인코드가 있으면 저장 생략 */
    if (code) {
      for (var i = n - 1; i >= Math.max(0, n - DUP_SCAN); i--) {
        if (String(at(data[i], 'code')) === code) return { ok: true, dup: true };
      }
    }

    /* 이름 정규화: 공백을 모두 뺀 표기가 같은 기존 이름이 있으면 그 표기(첫 등장)를 채택 */
    var key = inName.replace(/\s+/g, '');
    var name = inName;
    for (var j = 0; j < n; j++) {
      var ex = String(at(data[j], 'name'));
      if (ex && ex.replace(/\s+/g, '') === key) { name = ex.slice(0, 20); break; }
    }

    var row = [];
    row[COL.stamp - 1] = deFormula_(String(r.stamp || '').slice(0, 30));
    row[COL.name - 1] = deFormula_(name);
    row[COL.week - 1] = week;
    row[COL.theme - 1] = deFormula_(String(r.theme || '').slice(0, 40));
    row[COL.ref1 - 1] = deFormula_(String(r.ref1 || '').slice(0, 40));
    row[COL.s1 - 1] = Number(r.s1);
    row[COL.ref2 - 1] = deFormula_(String(r.ref2 || '').slice(0, 40));
    row[COL.s2 - 1] = Number(r.s2);
    row[COL.avg - 1] = Number(r.avg);
    row[COL.result - 1] = r.pass ? '합격' : '불합격';
    row[COL.code - 1] = deFormula_(code);
    row[COL.rep - 1] = FINAL;
    row[COL.start - 1] = deFormula_(start);
    row[COL.ver - 1] = deFormula_(String(r.ver || '').slice(0, 20));
    sh.appendRow(row);

    /* '최종' 갱신: 같은 (이름, 주차, 1주차주일)의 이전 행은 대표 해제.
       이름·주차가 같은 행만 날짜 키(groupKey_/ymd_)를 계산해 행 수에 비례한 비용을 줄인다 */
    var gk = groupKey_(name, week, start);
    for (var k = 0; k < n; k++) {
      var d = data[k];
      if (at(d, 'rep') !== FINAL || String(at(d, 'name')) !== name || Number(at(d, 'week')) !== week) continue;
      if (groupKey_(at(d, 'name'), at(d, 'week'), at(d, 'start')) === gk) {
        sh.getRange(k + 2, COL.rep).setValue('');
      }
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/** 기존 행 마이그레이션(멱등): 헤더 14열 보장, 1주차주일(M)이 빈 행에 start 기록,
 *  각 (이름, 주차, M) 그룹에서 가장 아래(최신) 행만 '최종' */
function migrate_(start) {
  start = String(start || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(start)) throw new Error('start 형식(YYYY-MM-DD)을 확인하세요');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('잠시 후 다시 시도해 주세요');
  try {
    var sh = ensureSheet_(SpreadsheetApp.getActive());
    var n = sh.getLastRow() - 1;
    if (n < 1) return { ok: true, rows: 0, finals: 0 };
    var data = sh.getRange(2, COL.name, n, COL.start - COL.name + 1).getValues();
    var at = function (row, k) { return row[COL[k] - COL.name]; };
    var startCol = [], lastOf = {};
    for (var i = 0; i < n; i++) {
      var m = at(data[i], 'start');
      if (m === '' || m === null) m = start;
      startCol.push([m]);
      lastOf[groupKey_(at(data[i], 'name'), at(data[i], 'week'), m)] = i;
    }
    var repCol = [], finals = 0;
    for (var j = 0; j < n; j++) {
      var isFinal = lastOf[groupKey_(at(data[j], 'name'), at(data[j], 'week'), startCol[j][0])] === j;
      repCol.push([isFinal ? FINAL : '']);
      if (isFinal) finals++;
    }
    sh.getRange(2, COL.start, n, 1).setValues(startCol);
    sh.getRange(2, COL.rep, n, 1).setValues(repCol);
    return { ok: true, rows: n, finals: finals };
  } finally {
    lock.releaseLock();
  }
}

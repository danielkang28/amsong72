/** 72구절 암송 시험 — 웹앱 서버 (구글시트 컨테이너 바인딩)
 *  doGet: index.html을 그대로 서빙하되 URL 파라미터(?s= ?v= ?p= ?t=)를 window.__QS로 주입
 *  submitResult: 리더가 시험을 제출하면 '성적' 시트에 한 줄씩 기록
 */
var SHEET_NAME = '성적';

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

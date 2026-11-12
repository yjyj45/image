/**
 * 7~8차시 그림책 개요 앱 - 교사용 앱스 스크립트
 *
 * 하는 일
 *  1) 앱 화면을 이 스크립트 주소로 보여 줍니다.   (Index.html 파일이 있을 때)
 *  2) 학생 작업을 '모둠 이름 + 내 번호'로 자동 저장하고, 7차시 작업을 8차시에
 *     다른 기기에서도 이어서 할 수 있게 합니다.   → 시트의 '저장' 탭
 *  3) 학생이 '시트로 보내기'를 누르면 쪽별 글·장면·프롬프트를 '제출' 탭에 한 줄씩 저장합니다.
 *     같은 모둠의 같은 쪽을 다시 보내면 그 줄이 바뀝니다.
 * 받지 않는 것: 학생 이름, 학번, 계정 정보. (모둠 이름·모둠 안 번호만 받습니다.)
 *
 * ── 설정 방법 (처음 한 번, 약 5분) ──────────────────────────
 *  1. 구글 시트를 새로 만들고 '확장 프로그램 > Apps Script'를 엽니다.
 *  2. 기본 코드를 지우고 이 코드를 붙여 넣습니다. (파일 이름: Code.gs)
 *  3. 왼쪽 '파일 +' > 'HTML'을 눌러 이름을 Index 로 만들고, 앱 화면 파일(Index.html)의
 *     내용을 모두 붙여 넣습니다. 저장합니다.
 *  4. 위쪽 함수 선택칸에서 'setup'을 골라 '실행'합니다. (권한 허용 창이 뜨면 내 계정으로 허용)
 *  5. '배포 > 새 배포 > 유형: 웹 앱'
 *       - 다음 사용자 인증으로 실행: 나
 *       - 액세스 권한이 있는 사용자: 모든 사용자  (학생이 로그인 없이 쓰려면 필요)
 *  6. 나온 웹 앱 주소(…/exec)를 학생에게 줍니다. 이것만 주면 됩니다.
 *     (제출 코드, 비밀 코드는 필요 없습니다.)
 *  7. 확인: 웹 앱 주소를 열어 앱이 보이면 성공입니다.
 *     학생 화면 위쪽에 '☁️ 저장됨 ✔'이 나오는지, '제출' '저장' 탭에 줄이 생기는지 봅니다.
 *
 * ── 알아 둘 점 ────────────────────────────────────────
 *  - 코드나 화면을 바꿨다면: '배포 > 배포 관리 > 연필 아이콘 > 버전: 새 버전 > 배포'를 눌러야
 *    적용됩니다. (주소는 그대로입니다. '새 배포'를 하면 주소가 바뀝니다.)
 *  - 주소를 아는 사람은 누구나 쓸 수 있고, 모둠 이름과 번호를 똑같이 쓰면 저장된 작업을 열 수
 *    있습니다. 수업이 끝나면 배포를 끝내세요. 학생 이름이나 학번은 쓰지 않게 안내하세요.
 *  - 같은 모둠 이름과 같은 번호를 두 학생이 쓰면 서로 덮어씁니다. 번호는 모둠 안에서 겹치지
 *    않게 하세요. 모둠 이름에 반을 넣으면(예: 3반 1모둠) 반끼리 겹치지 않습니다.
 *  - 학교 계정이 외부 접근을 막아 '모든 사용자' 배포가 안 되면, 앱의 '파일로 저장'을 쓰세요.
 *  - (예전 방식) 화면을 GitHub Pages 등 다른 곳에 두고 이 스크립트를 제출 창구로만 쓸 수도
 *    있습니다. 이때는 주소 뒤에 ?k=비밀코드를 붙인 '제출 코드'를 씁니다. 비밀 코드는 setup 실행
 *    때 로그에 나옵니다. 이 방식에서는 저장·이어서 하기가 켜지지 않습니다.
 *  - 한 번에 12쪽, 칸마다 2000자까지 받습니다.
 */

var SHEET_NAME = '제출';
var HEADERS = ['키', '보낸 시각', '모둠', '보낸 번호', '쪽', '담당', '글', '그릴 장면', '행동·표정', '배경',
  '색감·분위기', '구도', '최종 프롬프트', '수정했나요', '핵심 장면', '그림책 제목', '주인공 외형', '처음 프롬프트', '그림체'];
var MAX_ROWS_PER_SEND = 12;      // 한 번에 받을 쪽 수 (표지 + 본문 여유 있게)
var MAX_TEXT = 2000;             // 칸마다 받을 글자 수

var WORK_SHEET = '저장';
var WORK_HEADERS = ['키', '저장 시각', '모둠', '번호', '내용1', '내용2', '내용3', '내용4'];
var CHUNK = 40000;               // 시트 한 칸에는 5만 자까지만 들어가서, 나눠서 저장합니다.
var MAX_CHUNKS = 4;

/** 처음 한 번 직접 실행합니다. '제출'·'저장' 시트를 만들고 비밀 코드를 정합니다. */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('SECRET');
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    props.setProperty('SECRET', secret);
  }
  props.setProperty('SSID', SpreadsheetApp.getActiveSpreadsheet().getId());
  getSheet_();
  getWorkSheet_();
  Logger.log('설정이 끝났어요. 이제 웹 앱으로 배포해요.');
  Logger.log('(예전 방식에서만 쓰는) 비밀 코드: ' + secret);
}

/** 웹 앱 주소를 열면 앱 화면이 나옵니다. 주소 뒤에 ?ping=1 을 붙이면 연결 확인용 글자만 나옵니다. */
function doGet(e) {
  if (!(e && e.parameter && e.parameter.ping)) {
    try {
      return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('그림책 개요 짜기')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) { /* Index 파일이 없으면 아래 연결 확인 글자를 보여 줍니다. */ }
  }
  return json_({ ok: true, hello: '그림책 개요 앱 제출 받기' });
}

/** (예전 방식) 다른 곳에 둔 앱 화면에서 보낸 제출을 받습니다. */
function doPost(e) {
  try {
    var secret = PropertiesService.getScriptProperties().getProperty('SECRET');
    var given = e && e.parameter ? e.parameter.k : '';
    if (!secret || !given || given !== secret) return json_({ ok: false, error: 'code' });
    var data = JSON.parse((e.postData && e.postData.contents) || '{}');
    return json_(writeRows_(data));
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  }
}

/** 앱 화면에서 '시트로 보내기'를 누르면 불립니다. */
function submitRows(data) {
  try { return writeRows_(data); } catch (err) { return { ok: false, error: 'server' }; }
}

/** 앱 화면이 2.5초마다 자동으로 불러서 학생 작업을 저장합니다. */
function saveWork(group, mine, json) {
  var lock = LockService.getScriptLock();
  try {
    var key = workKey_(group, mine);
    if (!key) return { ok: false, error: 'key' };
    json = String(json == null ? '' : json);
    if (!json || json.length > CHUNK * MAX_CHUNKS) return { ok: false, error: 'size' };
    var chunks = [];
    for (var i = 0; i < json.length; i += CHUNK) chunks.push(json.slice(i, i + CHUNK));
    while (chunks.length < MAX_CHUNKS) chunks.push('');
    var now = nowText_();
    var line = [key, now, clean_(group), clean_(mine)].concat(chunks);

    lock.waitLock(20000);
    var sheet = getWorkSheet_();
    var rowNo = findRow_(sheet, key);
    if (!rowNo) {
      rowNo = Math.max(sheet.getLastRow(), 1) + 1;
      if (rowNo > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 100);
    }
    var range = sheet.getRange(rowNo, 1, 1, WORK_HEADERS.length);
    range.setNumberFormat('@');             // 글자로만 저장 (= 나 - 로 시작해도 수식으로 읽지 않게)
    range.setValues([line]);
    return { ok: true, at: now };
  } catch (err) {
    return { ok: false, error: 'server' };
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* 잠금을 못 잡았으면 무시 */ }
  }
}

/** '이어서 하기'에서 불립니다. 저장된 작업을 돌려줍니다. */
function loadWork(group, mine) {
  try {
    var key = workKey_(group, mine);
    if (!key) return { ok: true, json: null };
    var sheet = getWorkSheet_();
    var rowNo = findRow_(sheet, key);
    if (!rowNo) return { ok: true, json: null };
    var v = sheet.getRange(rowNo, 1, 1, WORK_HEADERS.length).getValues()[0];
    var json = [v[4], v[5], v[6], v[7]].join('');
    return { ok: true, json: json || null, at: String(v[1] || '') };
  } catch (err) {
    return { ok: false, error: 'server' };
  }
}

/** 제출 내용을 '제출' 탭에 씁니다. */
function writeRows_(data) {
  if (!data || data.v !== 1 || !Array.isArray(data.rows)) return { ok: false, error: 'format' };
  if (data.rows.length === 0 || data.rows.length > MAX_ROWS_PER_SEND) return { ok: false, error: 'size' };
  var group = clean_(data.group), sender = clean_(data.sender);
  if (!group) return { ok: false, error: 'group' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getSheet_();
    var index = {};                         // 키 -> 줄 번호
    var last = sheet.getLastRow();
    if (last > 1) {
      var keys = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) index[keys[i][0]] = i + 2;
    }
    var now = new Date();
    var count = 0;
    data.rows.forEach(function (r) {
      var page = clean_(r.page);
      if (!page) return;
      var key = group + '|' + page;
      var line = [key, now, group, sender, page, clean_(r.owner), clean_(r.text), clean_(r.scene), clean_(r.action),
        clean_(r.setting), clean_(r.mood), clean_(r.camera), clean_(r.prompt), r.edited ? '예' : '아니요',
        r.star ? '★' : '', clean_(data.title), clean_(data.look), clean_(r.first), clean_(r.style)];
      if (index[key]) {
        sheet.getRange(index[key], 1, 1, line.length).setValues([line]);
      } else {
        sheet.appendRow(line);
        index[key] = sheet.getLastRow();
      }
      count++;
    });
    return { ok: true, count: count };
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* 무시 */ }
  }
}

function ss_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SSID'));
}

function getSheet_() {
  var ss = ss_();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.hideColumns(1);                   // 키 칸은 숨겨 둡니다.
  } else {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);   // 칸 이름을 최신으로 맞춥니다.
  }
  return sheet;
}

function getWorkSheet_() {
  var ss = ss_();
  var sheet = ss.getSheetByName(WORK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(WORK_SHEET);
    sheet.getRange(1, 1, sheet.getMaxRows(), WORK_HEADERS.length).setNumberFormat('@');
    sheet.appendRow(WORK_HEADERS);
    sheet.setFrozenRows(1);
    sheet.hideColumns(1);
    sheet.hideColumns(5, 4);                // 저장 내용은 사람이 읽는 글이 아니라서 숨겨 둡니다.
  }
  return sheet;
}

function findRow_(sheet, key) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var keys = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) if (keys[i][0] === key) return i + 2;
  return 0;
}

/** 모둠 이름 + 번호로 저장 칸의 열쇠를 만듭니다. 비어 있으면 '' 을 돌려줍니다. */
function workKey_(group, mine) {
  var g = String(group == null ? '' : group).trim().replace(/\s+/g, ' ').slice(0, 60).toLowerCase();
  var m = String(mine == null ? '' : mine).replace(/[^0-9]/g, '').slice(0, 3);
  return g && m ? 'w|' + g + '|' + m : '';
}

function nowText_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M월 d일 HH:mm');
}

/** 글자 수를 줄이고, 수식으로 읽히는 글(= + - @ 로 시작)은 앞에 ' 를 붙여 글자로만 저장합니다. */
function clean_(v) {
  var t = String(v == null ? '' : v).replace(/\r\n?/g, '\n').trim().slice(0, MAX_TEXT);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

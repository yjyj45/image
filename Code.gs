/**
 * 7~8차시 그림책 개요 앱 - 활동 내용 받기 (교사용 앱스 스크립트)
 *
 * 하는 일: 학생이 앱에서 '시트로 보내기'를 누르면 쪽별 글·장면·프롬프트를
 *          이 시트의 '제출' 탭에 한 줄씩 저장합니다. (선택 기능)
 *          같은 모둠의 같은 쪽을 다시 보내면 새 줄이 아니라 그 줄이 바뀝니다.
 * 받지 않는 것: 학생 이름, 학번, 계정 정보. (모둠 이름·모둠 안 번호만 받습니다.)
 *
 * ── 설정 방법 (처음 한 번, 약 5분) ──────────────────────────
 *  1. 구글 시트를 새로 만들고 '확장 프로그램 > Apps Script'를 엽니다.
 *  2. 기본 코드를 모두 지우고 이 코드를 붙여 넣은 뒤 저장합니다.
 *  3. 위쪽 함수 선택칸에서 'setup'을 골라 '실행'합니다. (권한 허용 창이 뜨면 내 계정으로 허용)
 *  4. 아래 '실행 로그'에 나온 비밀 코드를 적어 둡니다.
 *  5. '배포 > 새 배포 > 유형: 웹 앱'
 *       - 다음 사용자 인증으로 실행: 나
 *       - 액세스 권한이 있는 사용자: 모든 사용자  (학생이 로그인 없이 보내려면 필요)
 *  6. 나온 웹 앱 주소(…/exec) 뒤에 ?k=비밀코드 를 붙이면 '제출 코드'가 됩니다.
 *       예) https://script.google.com/macros/s/AKfy.../exec?k=a1b2c3d4e5
 *  7. 앱 3단계의 '선생님이 제출 코드를 알려 줬나요?'에 제출 코드를 붙여 넣고
 *     '학생용 링크 복사'를 눌러, 그 링크(…#c=…)를 학생에게 줍니다.
 *     링크로 들어오면 제출 코드가 그 기기에 저절로 저장됩니다.
 *  8. 확인: 웹 앱 주소를 브라우저로 열어 {"ok":true,...}가 보이면 연결된 것입니다.
 *
 * ── 알아 둘 점 ────────────────────────────────────────
 *  - 코드를 바꿨다면: '배포 > 배포 관리 > 연필 아이콘 > 버전: 새 버전 > 배포'를 눌러야
 *    적용됩니다. (주소와 제출 코드는 그대로입니다. '새 배포'를 하면 주소가 바뀝니다.)
 *  - 제출 코드를 아는 사람은 누구나 시트에 보낼 수 있습니다. 수업이 끝나면 배포를 끝내거나
 *    스크립트 속성 SECRET 을 바꿔 닫아 두세요.
 *  - 학교 계정이 외부 접근을 막아 '모든 사용자' 배포가 안 되면, 앱의 '파일로 저장'을 쓰세요.
 *  - 한 번에 12쪽, 칸마다 2000자까지 받습니다.
 */

var SHEET_NAME = '제출';
var HEADERS = ['키', '보낸 시각', '모둠', '보낸 번호', '쪽', '담당', '글', '그릴 장면', '행동·표정', '배경',
  '색감·분위기', '구도', '최종 프롬프트', '수정했나요', '핵심 장면', '그림책 제목', '주인공 외형', '처음 프롬프트', '그림체'];
var MAX_ROWS_PER_SEND = 12;      // 한 번에 받을 쪽 수 (표지 + 본문 여유 있게)
var MAX_TEXT = 2000;             // 칸마다 받을 글자 수

/** 처음 한 번 직접 실행합니다. 비밀 코드를 정하고 '제출' 시트를 만듭니다. */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('SECRET');
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    props.setProperty('SECRET', secret);
  }
  getSheet_();
  Logger.log('제출 코드에 붙일 비밀 코드: ' + secret);
  Logger.log('학생에게 알려 줄 제출 코드 = (웹 앱 주소) + ?k=' + secret);
}

/** 연결 확인용. 웹 앱 주소를 브라우저로 열면 OK가 보입니다. */
function doGet() {
  return json_({ ok: true, hello: '그림책 개요 앱 제출 받기' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var secret = PropertiesService.getScriptProperties().getProperty('SECRET');
    var given = e && e.parameter ? e.parameter.k : '';
    if (!secret || !given || given !== secret) return json_({ ok: false, error: 'code' });

    var data = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (!data || data.v !== 1 || !Array.isArray(data.rows)) return json_({ ok: false, error: 'format' });
    if (data.rows.length === 0 || data.rows.length > MAX_ROWS_PER_SEND) return json_({ ok: false, error: 'size' });

    lock.waitLock(20000);
    var sheet = getSheet_();
    var group = clean_(data.group), sender = clean_(data.sender);
    if (!group) return json_({ ok: false, error: 'group' });

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
    return json_({ ok: true, count: count });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* 잠금을 못 잡았으면 무시 */ }
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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

/** 글자 수를 줄이고, 수식으로 읽히는 글(= + - @ 로 시작)은 앞에 ' 를 붙여 글자로만 저장합니다. */
function clean_(v) {
  var t = String(v == null ? '' : v).replace(/\r\n?/g, '\n').trim().slice(0, MAX_TEXT);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

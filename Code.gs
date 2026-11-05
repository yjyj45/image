/**
 * 7~8차시 그림책 개요 앱 - 활동 내용 받기 (교사용 앱스 스크립트)
 *
 * 하는 일: 학생이 앱에서 '시트로 보내기'를 누르면, 그 학생의 쪽별 글·장면·프롬프트를
 *          이 스크립트가 연결된 구글 시트의 '제출' 시트에 한 줄씩 저장합니다.
 *          같은 모둠의 같은 쪽을 다시 보내면 새 줄이 생기지 않고 그 줄이 바뀝니다.
 * 받지 않는 것: 학생 이름, 학번, 계정 정보. 모둠 이름(번호)과 모둠 안 번호만 받습니다.
 *
 * 처음 한 번 하는 설정은 'Apps스크립트_설정안내.md'를 보세요.
 */

var SHEET_NAME = '제출';
var HEADERS = ['키', '보낸 시각', '모둠', '보낸 번호', '쪽', '담당', '글', '그릴 장면', '행동·표정', '배경',
  '색감·분위기', '구도', '최종 프롬프트', '직접 고쳤나요', '핵심 장면', '그림책 제목', '주인공 외형'];
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
        r.star ? '★' : '', clean_(data.title), clean_(data.look)];
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

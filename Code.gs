/**
 * 교시별_요청사항_일정표 → Google 캘린더 자동 동기화
 *
 * - 월별 시트(2026년 09월 ...)에서 "요청 선생님" 또는 "비고(요청사항)"이 채워진 행을 캘린더 일정으로 등록
 * - 누가 시트를 수정하면 즉시 반영 (설치형 onEdit 트리거) + 1시간마다 전체 동기화(안전장치)
 * - 새 요청이 등록되면 OWNER_EMAIL 로 이메일 알림
 * - 각 행의 이벤트ID는 숨김 열(I열)에 저장 → 수정 시 기존 일정 업데이트, 요청 지우면 일정 삭제
 *
 * 처리상태(H열) 자동 규칙 (5분마다 갱신):
 *   확인중 : 요청은 있는데 아직 캘린더 등록 전 (날짜/시간을 못 읽은 경우 포함)
 *   대기   : 캘린더 등록 완료, 아직 시작 전
 *   진행중 : 지금 시각이 해당 교시 시간 안에 있음
 *   완료   : 해당 교시 시간이 지남
 *   보류   : 직접 지정 (스크립트가 절대 덮어쓰지 않음)
 *
 * 월별 제출 서류(출근기록부/활동일지/급식확인서)는 아래 "월별 제출 서류 자동 생성" 섹션 참고
 *
 * 설정(성명/이메일/학교 등)은 코드가 아니라 시트의 "설정" 탭에서 바꿉니다.
 * 최초 1회: 메뉴 [디지털튜터 캘린더] > [처음 설치 안내] → [트리거 설치(최초 1회)] 실행 후 권한 승인
 */

// ───────────────────────── 설정 (시트의 "설정" 탭에서 읽음) ─────────────────────────
// 코드를 고치지 않고 "설정" 탭의 값만 바꾸면 됩니다. 탭이 없으면 아래 기본값으로 자동 생성됩니다.
var CONFIG_SHEET_NAME = '설정';
var CONFIG_DEFAULTS = [
  // [키, 기본값, 설명]
  ['OWNER_EMAIL',      'your-email@gmail.com', '알림/서류 PDF를 받을 이메일 (캘린더 소유자)'],
  ['TUTOR_NAME',       '홍길동',               '성명 (서류 작성자)'],
  ['TUTOR_TITLE',      '디지털튜터',           '직'],
  ['SCHOOL_NAME',      '○○중학교',             '학교명'],
  ['SCHOOL_YEAR',      '2026학년도',           '학년도 (서류 제목)'],
  ['DEFAULT_PLACE',    '디지털 교실',          '기본 근무장소'],
  ['TITLE_PREFIX',     '[디지털튜터]',         '캘린더 일정 제목 앞에 붙는 말'],
  ['DEFAULT_ACTIVITY', '디지털 기기 점검 및 수업 지원 준비', '요청이 없는 날의 지도 내용'],
  ['DEFAULT_CLASS',    '전체',                 '요청이 없는 날의 지도반'],
  ['MEAL_REASON',      '근무',                 '급식확인서 급식사유'],
  ['MEAL_CUTOFF',      '13:00',                '급식 포함 기준: 이 시각 이후에 퇴근하는 날만 급식']
];
var configCache_ = null;
/** 설정값 읽기: C('TUTOR_NAME') */
function C(key) {
  if (!configCache_) {
    configCache_ = {};
    CONFIG_DEFAULTS.forEach(function (d) { configCache_[d[0]] = d[1]; });
    try {
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
      if (sh && sh.getLastRow() >= 2) {
        sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
          var k = String(r[0] || '').trim();
          if (k && r[1] !== '' && r[1] !== null) configCache_[k] = (r[1] instanceof Date) ? Utilities.formatDate(r[1], TZ, 'HH:mm') : String(r[1]).trim();
        });
      }
    } catch (e) { console.error('설정 읽기 실패: ' + e); }
  }
  return configCache_[key];
}
/** "설정" 탭이 없으면 기본값으로 생성 */
function ensureConfigSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(CONFIG_SHEET_NAME)) return false;
  var sh = ss.insertSheet(CONFIG_SHEET_NAME, 0);
  var rows = [['항목', '값', '설명']].concat(CONFIG_DEFAULTS);
  sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat('@');
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#d9d9d9');
  sh.getRange(2, 2, rows.length - 1, 1).setBackground('#fff8dc');
  [160, 260, 380].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(1);
  return true;
}
var HEADER_ROW = 2;                        // 헤더가 있는 행
var FIRST_DATA_ROW = 3;                    // 데이터 시작 행
var COL = { DATE: 1, DAY: 2, WEEKDAY: 3, PERIOD: 4, TIME: 5, TEACHER: 6, NOTE: 7, STATUS: 8, EVENT_ID: 9 };
var EVENT_ID_HEADER = '이벤트ID(자동)';
var TZ = 'Asia/Seoul';
var TIME_BUDGET_MS = 4.5 * 60 * 1000;      // Apps Script 6분 제한 대비 (초과 시 1분 후 자동 이어서 실행)

var STATUS_CHECKING = '확인중';
var STATUS_WAITING = '대기';
var STATUS_ONGOING = '진행중';
var STATUS_DONE = '완료';
var STATUS_HOLD = '보류';
var AUTO_STATUSES = [STATUS_CHECKING, STATUS_WAITING, STATUS_ONGOING];
var MANUAL_STATUSES = [STATUS_HOLD];                       // 사용자가 직접 지정, 스크립트가 건드리지 않음 (보류)
var DROPDOWN_STATUSES = [STATUS_WAITING, STATUS_CHECKING, STATUS_ONGOING, STATUS_DONE, STATUS_HOLD];
var OBSOLETE_STATUSES = ['처리중'];                          // 드롭다운에서 제거할 옛 값

// ───────────────────────── 메뉴 ─────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('디지털튜터 캘린더')
    .addItem('처음 설치 안내', 'showSetupGuide')
    .addItem('전체 동기화 (지금 실행)', 'syncAll')
    .addItem('처리상태 갱신 (지금 실행)', 'updateStatuses')
    .addSeparator()
    .addItem('월별 서류 생성 (출근기록부/활동일지/급식확인서)', 'generateMonthlyDocsPrompt')
    .addItem('월별 서류 PDF 메일로 받기', 'emailMonthlyDocsPrompt')
    .addSeparator()
    .addItem('트리거 설치 (최초 1회)', 'setupTriggers')
    .addToUi();
}

/** 처음 설치 안내 (사본으로 받은 사람용) */
function showSetupGuide() {
  var created = ensureConfigSheet_();
  SpreadsheetApp.getUi().alert('디지털튜터 자동화 - 처음 설치 안내',
    '1) "' + CONFIG_SHEET_NAME + '" 탭에서 성명 / 이메일 / 학교명 등을 본인 것으로 바꾸세요.' +
    (created ? ' (방금 기본값으로 만들었습니다)' : '') + '\n' +
    '2) "근무일정" 탭에 날짜 / 요일 / 출근 / 퇴근 / 실근무시간 / 비고 / 활동 순으로 근무일정을 입력하세요.\n' +
    '3) 월별 시트(예: 2026년 09월)는 기존 양식 그대로 사용하면 됩니다. 2행이 헤더, A열이 "날짜"여야 합니다.\n' +
    '4) 메뉴 > "트리거 설치 (최초 1회)" 를 누르고 Google 권한을 승인하세요.\n' +
    '   (Google에서 확인하지 않은 앱 경고가 뜨면 "고급" > "이동" 을 누르면 됩니다)\n\n' +
    '설치가 끝나면 시트에 요청이 입력될 때마다 캘린더 등록 + 이메일 알림이 오고,\n매월 1일 08시에 전월 서류 PDF가 이메일로 옵니다.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ───────────────────────── 트리거 ─────────────────────────
function setupTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureConfigSheet_();
  configCache_ = null;
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('updateStatuses').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('monthlyDocsJob').timeBased().onMonthDay(1).atHour(8).create();   // 매월 1일 08시: 전월 서류 PDF 메일
  syncAll();
  try { ss.toast('트리거 설치 완료! 이제 시트가 수정되면 자동으로 캘린더에 등록됩니다.', '디지털튜터 캘린더', 8); } catch (e) {}
}

/** 설치형 onEdit 트리거: 수정된 행만 동기화 */
function onEditInstallable(e) {
  try {
    var sheet = e.range.getSheet();
    if (!isScheduleSheet_(sheet)) return;
    var startRow = Math.max(e.range.getRow(), FIRST_DATA_ROW);
    var endRow = e.range.getLastRow();
    if (endRow < FIRST_DATA_ROW) return;
    var editor = (e.user && e.user.getEmail && e.user.getEmail()) || '';
    syncRows_(sheet, startRow, endRow, editor);
  } catch (err) {
    console.error('onEdit 오류: ' + err);
  }
}

/** 전체 시트 동기화 (시간 트리거 / 메뉴) */
function syncAll() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  var deadline = Date.now() + TIME_BUDGET_MS;
  var aborted = false;
  try {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i = 0; i < sheets.length && !aborted; i++) {
      var sheet = sheets[i];
      if (!isScheduleSheet_(sheet)) continue;
      var last = sheet.getLastRow();
      if (last >= FIRST_DATA_ROW) aborted = !syncRows_(sheet, FIRST_DATA_ROW, last, '', deadline);
    }
  } finally {
    lock.releaseLock();
  }
  cleanupContinueTriggers_();
  if (aborted) {
    // 시간 예산 초과 → 1분 뒤 이어서 실행
    ScriptApp.newTrigger('syncAll').timeBased().after(60 * 1000).create();
    console.log('시간 예산 초과, 1분 후 이어서 동기화합니다.');
  } else {
    console.log('전체 동기화 완료');
  }
}

/** 이어서 실행용 일회성 트리거 정리 (설치형 onEdit / 매시간 트리거는 유지) */
function cleanupContinueTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAll' && t.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
}

/**
 * 처리상태만 시간 기준으로 갱신 (5분마다). 캘린더 API를 호출하지 않아 가볍다.
 * 요청 있음 + 이벤트ID 없음 → 확인중 / 시작 전 → 대기 / 진행 중 → 진행중 / 지난 뒤 → 완료
 */
function updateStatuses() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var changed = 0;
    SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sheet) {
      if (!isScheduleSheet_(sheet)) return;
      var last = sheet.getLastRow();
      if (last < FIRST_DATA_ROW) return;
      var values = sheet.getRange(FIRST_DATA_ROW, 1, last - FIRST_DATA_ROW + 1, COL.EVENT_ID).getValues();
      var currentDate = null;
      for (var i = 0; i < values.length; i++) {
        var r = values[i];
        var d = parseDate_(r[COL.DATE - 1]);
        if (d) currentDate = d;
        var teacher = String(r[COL.TEACHER - 1] || '').trim();
        var note = String(r[COL.NOTE - 1] || '').trim();
        if (!teacher && !note) continue;
        var status = String(r[COL.STATUS - 1] || '').trim();
        var eventId = String(r[COL.EVENT_ID - 1] || '').trim();
        var times = parseTimeRange_(String(r[COL.TIME - 1] || ''));
        var desired;
        if (!eventId || !currentDate || !times) {
          desired = STATUS_CHECKING;
        } else {
          var se = buildStartEnd_(currentDate, times);
          desired = computeStatus_(se.start, se.end);
        }
        if (applyStatus_(sheet, FIRST_DATA_ROW + i, status, desired)) changed++;
      }
    });
    console.log('처리상태 갱신: ' + changed + '건 변경');
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────── 핵심 로직 ─────────────────────────
function isScheduleSheet_(sheet) {
  if (sheet.getLastRow() < HEADER_ROW) return false;
  var h = String(sheet.getRange(HEADER_ROW, COL.DATE).getValue()).trim();
  return h === '날짜';
}

function ensureEventIdColumn_(sheet) {
  var cell = sheet.getRange(HEADER_ROW, COL.EVENT_ID);
  if (String(cell.getValue()).trim() !== EVENT_ID_HEADER) {
    cell.setValue(EVENT_ID_HEADER);
    sheet.hideColumns(COL.EVENT_ID);
  }
}

function syncRows_(sheet, startRow, endRow, editor, deadline) {
  ensureEventIdColumn_(sheet);
  ensureStatusOptions_(sheet, endRow);
  var cal = CalendarApp.getDefaultCalendar();
  var ss = sheet.getParent();
  var sheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();

  // 날짜 열이 병합되어 있으므로, 데이터 시작행부터 읽어 날짜를 아래로 채움
  var values = sheet.getRange(FIRST_DATA_ROW, 1, endRow - FIRST_DATA_ROW + 1, COL.EVENT_ID).getValues();
  var currentDate = null;
  var created = [];

  var completed = true;
  for (var i = 0; i < values.length; i++) {
    if (deadline && Date.now() > deadline) { completed = false; break; }
    var row = FIRST_DATA_ROW + i;
    var r = values[i];
    var d = parseDate_(r[COL.DATE - 1]);
    if (d) currentDate = d;
    if (row < startRow) continue;

    var period = String(r[COL.PERIOD - 1] || '').trim();
    var timeStr = String(r[COL.TIME - 1] || '').trim();
    var teacher = String(r[COL.TEACHER - 1] || '').trim();
    var note = String(r[COL.NOTE - 1] || '').trim();
    var status = String(r[COL.STATUS - 1] || '').trim();
    var eventId = String(r[COL.EVENT_ID - 1] || '').trim();
    var hasRequest = !!(teacher || note);

    // 요청이 지워진 경우 → 일정 삭제 + 상태를 "대기"(기본값)로 되돌림
    if (!hasRequest) {
      if (eventId) {
        try { var old = cal.getEventById(eventId); if (old) old.deleteEvent(); } catch (e) {}
        sheet.getRange(row, COL.EVENT_ID).setValue('');
        if (status !== STATUS_WAITING) sheet.getRange(row, COL.STATUS).setValue(STATUS_WAITING);
      }
      continue;
    }

    var times = parseTimeRange_(timeStr);
    if (!currentDate || !times) {
      // 날짜/시간을 알 수 없어 등록 불가 → 확인중
      applyStatus_(sheet, row, status, STATUS_CHECKING);
      continue;
    }

    var se = buildStartEnd_(currentDate, times);
    var start = se.start, end = se.end;

    var title = C('TITLE_PREFIX') + ' ' + period + (teacher ? ' ' + teacher : '') + (note ? ' - ' + note : '');
    if (status && AUTO_STATUSES.indexOf(status) < 0) title = '[' + status + '] ' + title;
    var desc = [
      '요청 선생님: ' + (teacher || '-'),
      '요청사항: ' + (note || '-'),
      '처리상태: ' + (status || '-'),
      '',
      '시트: ' + sheet.getName() + ' ' + row + '행',
      sheetUrl
    ].join('\n');

    var ev = null;
    if (eventId) { try { ev = cal.getEventById(eventId); } catch (e) { ev = null; } }

    if (ev) {
      // 변경사항 있으면 업데이트
      if (ev.getTitle() !== title) ev.setTitle(title);
      if (ev.getStartTime().getTime() !== start.getTime() || ev.getEndTime().getTime() !== end.getTime()) ev.setTime(start, end);
      if (ev.getDescription() !== desc) ev.setDescription(desc);
    } else {
      ev = cal.createEvent(title, start, end, { description: desc });
      ev.removeAllReminders();
      ev.addPopupReminder(10);
      sheet.getRange(row, COL.EVENT_ID).setValue(ev.getId());
      created.push({ when: Utilities.formatDate(start, TZ, 'yyyy-MM-dd (E) HH:mm') + '~' + Utilities.formatDate(end, TZ, 'HH:mm'), period: period, teacher: teacher, note: note, sheet: sheet.getName(), row: row });
    }
    // 등록 완료 → 시간 기준 상태 반영
    applyStatus_(sheet, row, status, computeStatus_(start, end));
  }

  if (created.length) notifyCreated_(created, editor, sheetUrl);
  return completed;
}

function notifyCreated_(created, editor, sheetUrl) {
  var lines = created.map(function (c) {
    return '• ' + c.when + ' ' + c.period + (c.teacher ? ' / ' + c.teacher + ' 선생님' : '') + (c.note ? ' / ' + c.note : '');
  });
  var subject = '[디지털튜터 요청] 새 요청 ' + created.length + '건 캘린더 등록됨';
  var body = '새로운 수업지원 요청이 캘린더에 등록되었습니다.\n\n' + lines.join('\n') +
    (editor ? '\n\n수정한 사람: ' + editor : '') +
    '\n\n시트 열기: ' + sheetUrl;
  try { MailApp.sendEmail(C('OWNER_EMAIL'), subject, body); } catch (e) { console.error('메일 전송 실패: ' + e); }
}

// ───────────────────────── 처리상태 ─────────────────────────
/** 시작 전 → 대기, 진행 중 → 진행중, 지난 뒤 → 완료 (보류만 직접 지정) */
function computeStatus_(start, end) {
  var now = new Date();
  if (now < start) return STATUS_WAITING;
  if (now <= end) return STATUS_ONGOING;
  return STATUS_DONE;
}

/** 완료/보류는 절대 덮어쓰지 않음. 값이 바뀌었으면 true */
function applyStatus_(sheet, row, current, desired) {
  if (MANUAL_STATUSES.indexOf(current) >= 0) return false;
  // 목록에 없는 잘못된 값(옛 '처리중', 오타 등)은 정리: 시간이 지난 행이면 기본값 '대기'로
  if (current && DROPDOWN_STATUSES.indexOf(current) < 0) desired = desired || STATUS_WAITING;
  if (!desired) return false;
  if (current === desired) return false;
  sheet.getRange(row, COL.STATUS).setValue(desired);
  return true;
}

/** 처리상태 드롭다운 목록 정리: 필요한 값 추가, 옛 값(처리중) 제거 */
function ensureStatusOptions_(sheet, lastRow) {
  try {
    var cell = sheet.getRange(FIRST_DATA_ROW, COL.STATUS);
    var rule = cell.getDataValidation();
    if (!rule || rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return;
    var args = rule.getCriteriaValues();
    var values = (args[0] || []).filter(function (v) { return OBSOLETE_STATUSES.indexOf(v) < 0; });
    var missing = DROPDOWN_STATUSES.filter(function (v) { return values.indexOf(v) < 0; });
    if (!missing.length && values.length === (args[0] || []).length) return;
    var newRule = rule.copy().requireValueInList(values.concat(missing), args[1] !== false).build();
    sheet.getRange(FIRST_DATA_ROW, COL.STATUS, lastRow - FIRST_DATA_ROW + 1, 1).setDataValidation(newRule);
  } catch (e) {
    console.error('드롭다운 옵션 정리 실패: ' + e);
  }
}

// ───────────────────────── 파서 ─────────────────────────
function buildStartEnd_(date, times) {
  return {
    start: new Date(date.getFullYear(), date.getMonth(), date.getDate(), times.sh, times.sm),
    end: new Date(date.getFullYear(), date.getMonth(), date.getDate(), times.eh, times.em)
  };
}

function parseDate_(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === 'number') { // 시리얼 넘버
    var base = new Date(1899, 11, 30);
    return new Date(base.getFullYear(), base.getMonth(), base.getDate() + Math.floor(v));
  }
  var s = String(v).trim();
  var m = s.match(/(\d{4})[-./년]\s*(\d{1,2})[-./월]\s*(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function parseTimeRange_(s) {
  var m = String(s).match(/(\d{1,2})\s*:\s*(\d{2})\s*[~\-–]\s*(\d{1,2})\s*:\s*(\d{2})/);
  if (!m) return null;
  return { sh: +m[1], sm: +m[2], eh: +m[3], em: +m[4] };
}


// ═══════════════════════════════════════════════════════════════════════
//  월별 제출 서류 자동 생성 (출근기록부 / 활동일지 / 급식확인서)
//  - "근무일정" 탭(날짜/요일/출근/퇴근/실근무시간/비고/활동)에서 출퇴근 시간을 읽고
//    (비고 = 근무일정표 원문(서류에 안 나감), 활동 = 서류에 표기할 특이 활동. 예: 동아리 활동, 진산제 지원)
//  - 월별 요청 시트(2026년 09월 ...)에서 지도반 / 교과·지도내용을 만들어
//  - 출근기록부_9월 / 활동일지_9월 / 급식확인서_9월 시트를 생성한다 (A4 1장씩)
//  - 메뉴 또는 매월 1일 08시 자동 실행 → PDF 3장을 OWNER_EMAIL 로 발송
// ═══════════════════════════════════════════════════════════════════════
var WORK_SHEET_NAME = '근무일정';
var DOC_ROWS = 14;                                             // 양식의 행 수
// 성명/학교/근무장소/급식 기준 등은 "설정" 탭 참고 (C('TUTOR_NAME') 등으로 읽음)                                             // 양식의 행 수
var WEEKDAYS_KO = ['일', '월', '화', '수', '목', '금', '토'];

function docSheetName_(kind, month) { return kind + '_' + month + '월'; }

/** 메뉴: 월 입력받아 생성 */
function generateMonthlyDocsPrompt() {
  var ui = SpreadsheetApp.getUi();
  var now = new Date();
  var res = ui.prompt('월별 서류 생성', '생성할 연-월을 입력하세요 (예: 2026-9)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var m = String(res.getResponseText() || (now.getFullYear() + '-' + (now.getMonth() + 1))).match(/(\d{4})\D+(\d{1,2})/);
  if (!m) { ui.alert('형식이 잘못되었습니다. 예: 2026-9'); return; }
  var r = generateDocsForMonth(+m[1], +m[2]);
  ui.alert(r.message);
}

/** 메뉴: 월 입력받아 PDF 메일 발송 */
function emailMonthlyDocsPrompt() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('서류 PDF 메일 발송', '연-월을 입력하세요 (예: 2026-9)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var m = String(res.getResponseText()).match(/(\d{4})\D+(\d{1,2})/);
  if (!m) { ui.alert('형식이 잘못되었습니다. 예: 2026-9'); return; }
  generateDocsForMonth(+m[1], +m[2]);
  emailMonthlyDocs(+m[1], +m[2]);
  ui.alert(C('OWNER_EMAIL') + ' 로 PDF 3장을 보냈습니다.');
}

/** 매월 1일 자동 실행: 전월 서류 생성 + PDF 메일 */
function monthlyDocsJob() {
  var now = new Date();
  var prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  generateDocsForMonth(prev.getFullYear(), prev.getMonth() + 1);
  emailMonthlyDocs(prev.getFullYear(), prev.getMonth() + 1);
}

// ───────────────────────── 데이터 수집 ─────────────────────────
/** 근무일정 탭 → [{date, start:'09:00', end:'14:30', hours, note}] (해당 월, 날짜순) */
function getWorkDays_(year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WORK_SHEET_NAME);
  if (!sheet) throw new Error('"' + WORK_SHEET_NAME + '" 탭이 없습니다. 날짜/요일/출근/퇴근/실근무시간/비고/활동 열로 만들어 주세요.');
  var values = sheet.getDataRange().getValues();
  var days = [];
  for (var i = 1; i < values.length; i++) {
    var d = parseDate_(values[i][0]);
    if (!d || d.getFullYear() !== year || d.getMonth() + 1 !== month) continue;
    var start = timeStr_(values[i][2]), end = timeStr_(values[i][3]);
    if (!start || !end) continue;
    days.push({
      date: d, start: start, end: end,
      hours: parseFloat(values[i][4]) || 0,
      note: String(values[i][5] || '').trim(),          // 근무일정표 비고 (서류에 미표기)
      activity: String(values[i][6] || '').trim()       // 서류 표기용 특이 활동
    });
  }
  days.sort(function (a, b) { return a.date - b.date; });
  return days;
}

/** 요청 시트 → { '2026-09-14': [{period, time, teacher, note}] } */
function getRequestsByDay_(year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var map = {};
  ss.getSheets().forEach(function (sheet) {
    if (!isScheduleSheet_(sheet)) return;
    var last = sheet.getLastRow();
    if (last < FIRST_DATA_ROW) return;
    var values = sheet.getRange(FIRST_DATA_ROW, 1, last - FIRST_DATA_ROW + 1, COL.STATUS).getValues();
    var currentDate = null;
    values.forEach(function (r) {
      var d = parseDate_(r[COL.DATE - 1]);
      if (d) currentDate = d;
      if (!currentDate || currentDate.getFullYear() !== year || currentDate.getMonth() + 1 !== month) return;
      var teacher = String(r[COL.TEACHER - 1] || '').trim();
      var note = String(r[COL.NOTE - 1] || '').trim();
      if (!teacher && !note) return;
      var key = dateKey_(currentDate);
      (map[key] = map[key] || []).push({
        period: String(r[COL.PERIOD - 1] || '').trim(),
        time: String(r[COL.TIME - 1] || '').trim(),
        teacher: teacher, note: note
      });
    });
  });
  return map;
}

/** "2학년 2반 수업지원" → "2-2" (없으면 null) */
function classShort_(text) {
  var m = String(text).match(/(\d)\s*학년\s*(\d{1,2})\s*반/);
  return m ? m[1] + '-' + m[2] : null;
}

/**
 * 하루치 요청 + 근무일정 비고 → 서류용 문구
 *  classes : "2-2, 2-10, 2-9"          (활동일지 지도반)
 *  detail  : "1교시 2학년 2반 수업 지원(김○○), ..."  (활동일지 교과/지도 내용)
 *  brief   : "2-2, 2-10, 2-9 수업지원"   (출근기록부 비고)
 */
function buildDayContent_(reqs, workNote) {
  reqs = reqs || [];
  var classes = [], details = [], extras = [];
  reqs.forEach(function (q) {
    var cls = classShort_(q.note);
    if (cls && classes.indexOf(cls) < 0) classes.push(cls);
    // 예: "1교시 2-2(김○○)" / 반 표기가 없으면 요청사항 원문 사용: "5교시 동아리 수업 지원(이○○)"
    var body = cls ? cls : (q.note || '수업 지원');
    details.push(q.period + ' ' + body + (q.teacher ? '(' + q.teacher + ')' : ''));
    if (!cls && q.note) extras.push(q.note.replace(/\s*(수업\s*)?지원\s*$/, ''));
  });
  var briefParts = [];
  if (workNote) briefParts.push(workNote);
  if (classes.length) briefParts.push(classes.join(', ') + ' 수업지원');
  if (extras.length) briefParts.push(extras.join(', ') + ' 지원');
  if (!briefParts.length) briefParts.push('기기 점검');

  var detailText = details.join(', ');
  if (detailText && classes.length) detailText += ' 수업 지원';
  if (workNote) detailText = workNote + (detailText ? ' / ' + detailText : '');
  if (!detailText) detailText = C('DEFAULT_ACTIVITY');

  return {
    classes: classes.length ? classes.join(', ') : C('DEFAULT_CLASS'),
    detail: detailText,
    brief: briefParts.join(' / ')
  };
}

// ───────────────────────── 서류 생성 ─────────────────────────
function generateDocsForMonth(year, month) {
  var days = getWorkDays_(year, month);
  if (!days.length) throw new Error(year + '년 ' + month + '월 근무일정이 "' + WORK_SHEET_NAME + '" 탭에 없습니다.');
  var reqs = getRequestsByDay_(year, month);
  var rows = days.map(function (w) {
    var c = buildDayContent_(reqs[dateKey_(w.date)], w.activity);
    return { work: w, content: c };
  });
  var totalHours = rows.reduce(function (s, r) { return s + r.work.hours; }, 0);

  var s1 = buildAttendanceSheet_(year, month, rows, totalHours);
  var s2 = buildActivitySheet_(year, month, rows, totalHours);
  var s3 = buildMealSheet_(year, month, rows);
  return {
    sheets: [s1, s2, s3],
    message: month + '월 서류 생성 완료\n· ' + s1.getName() + '\n· ' + s2.getName() + '\n· ' + s3.getName() +
      '\n근무 ' + rows.length + '일 / 총 ' + totalHours + '시간'
  };
}

function resetDocSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) { sh.clear(); sh.clearFormats(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); }
  else { sh = ss.insertSheet(name, ss.getSheets().length); }
  sh.setHiddenGridlines(true);
  return sh;
}

function styleTable_(sh, r1, c1, nRows, nCols, opts) {
  var rng = sh.getRange(r1, c1, nRows, nCols);
  rng.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  rng.setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true).setFontSize(opts && opts.fontSize || 10);
}

function docHeader_(sh, title, lastCol) {
  sh.getRange(1, 1, 1, lastCol).merge().setValue(title)
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 46);
  sh.getRange(2, 1, 1, lastCol).merge()
    .setValue('작성자: ' + C('SCHOOL_NAME') + ' ' + C('TUTOR_TITLE') + ' ' + C('TUTOR_NAME') + '            (서명)')
    .setHorizontalAlignment('right').setFontSize(10);
  sh.setRowHeight(2, 28);
}

function docFooter_(sh, row, month, totalHours, lastCol) {
  // ( 9 ) 월 총 근무시간 | ( 59.5 ) 시간 | 담당 확인 | (빈칸)
  var c = 1;
  sh.getRange(row, c, 1, 2).merge().setValue('( ' + month + ' ) 월 총 근무시간').setBackground('#d9d9d9'); c += 2;
  sh.getRange(row, c, 1, 2).merge().setValue('( ' + totalHours + ' ) 시간'); c += 2;
  sh.getRange(row, c).setValue('담당 확인').setBackground('#d9d9d9'); c += 1;
  sh.getRange(row, c, 1, lastCol - c + 1).merge().setValue('');
  sh.setRowHeight(row, 44);
  sh.getRange(row, 1, 1, lastCol).setFontWeight('bold').setWrap(false);
}

function fmtDay_(d) { return (d.getMonth() + 1) + '/' + d.getDate() + '(' + WEEKDAYS_KO[d.getDay()] + ')'; }

/** 출근 기록부 */
function buildAttendanceSheet_(year, month, rows, totalHours) {
  var name = docSheetName_('출근기록부', month);
  var sh = resetDocSheet_(name);
  var cols = 7;
  docHeader_(sh, C('SCHOOL_YEAR') + ' ' + C('TUTOR_TITLE') + ' 출근 기록부(' + month + '월)', cols);
  var hdr = ['순', '날짜\n(요일)', '근무장소', '출근 시간', '퇴근 시간', '비고\n(안전 점검)', '활동\n시간'];
  sh.getRange(3, 1, 1, cols).setValues([hdr]).setBackground('#d9d9d9').setFontWeight('bold');
  sh.setRowHeight(3, 40);
  var body = [];
  for (var i = 0; i < DOC_ROWS; i++) {
    var r = rows[i];
    body.push(r ? [i + 1, fmtDay_(r.work.date), C('DEFAULT_PLACE'), r.work.start, r.work.end, r.content.brief, r.work.hours]
                : [i + 1, '', '', '', '', '', '']);
  }
  sh.getRange(4, 1, DOC_ROWS, cols).setNumberFormat('@');   // "09:00" 이 시간값으로 바뀌지 않도록
  sh.getRange(4, 1, DOC_ROWS, cols).setValues(body);
  for (var k = 0; k < DOC_ROWS; k++) sh.setRowHeight(4 + k, 40);
  styleTable_(sh, 3, 1, DOC_ROWS + 2, cols, { fontSize: 10 });
  docFooter_(sh, 4 + DOC_ROWS, month, totalHours, cols);
  [40, 78, 88, 72, 72, 300, 52].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  return sh;
}

/** 활동 일지 */
function buildActivitySheet_(year, month, rows, totalHours) {
  var name = docSheetName_('활동일지', month);
  var sh = resetDocSheet_(name);
  var cols = 6;
  docHeader_(sh, C('SCHOOL_YEAR') + ' ' + C('TUTOR_TITLE') + ' 활동 일지(' + month + '월)', cols);
  var hdr = ['순', '날짜\n(요일)', '지도반', '장소', '시간', '교과 / 지도 내용'];
  sh.getRange(3, 1, 1, cols).setValues([hdr]).setBackground('#d9d9d9').setFontWeight('bold');
  sh.setRowHeight(3, 40);
  var body = [];
  for (var i = 0; i < DOC_ROWS; i++) {
    var r = rows[i];
    body.push(r ? [i + 1, fmtDay_(r.work.date), r.content.classes, C('DEFAULT_PLACE'), r.work.hours, r.content.detail]
                : [i + 1, '', '', '', '', '']);
  }
  sh.getRange(4, 1, DOC_ROWS, cols).setValues(body);
  for (var k = 0; k < DOC_ROWS; k++) sh.setRowHeight(4 + k, 40);
  styleTable_(sh, 3, 1, DOC_ROWS + 2, cols, { fontSize: 9 });
  sh.getRange(4, 6, DOC_ROWS, 1).setHorizontalAlignment('left');
  docFooter_(sh, 4 + DOC_ROWS, month, totalHours, cols);
  [40, 78, 120, 80, 62, 340].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  return sh;
}

/** 급식확인서 (퇴근이 13:00 이후인 근무일만) */
function buildMealSheet_(year, month, rows) {
  var name = docSheetName_('급식확인서', month);
  var sh = resetDocSheet_(name);
  var cols = 5;
  var mealDays = rows.filter(function (r) { return toMin_(r.work.end) > toMin_(C('MEAL_CUTOFF')) && toMin_(r.work.start) < toMin_(C('MEAL_CUTOFF')); });
  var dayList = mealDays.map(function (r) { return (r.work.date.getMonth() + 1) + '/' + r.work.date.getDate(); }).join(', ');
  var period = year + '년 ' + month + '월 근무일 중 ' + mealDays.length + '일\n(' + dayList + ')';
  var lastDay = rows.length ? rows[rows.length - 1].work.date : new Date(year, month, 0);

  sh.getRange(1, 1, 1, cols).merge().setValue('급 식 확 인 서')
    .setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontLine('underline');
  sh.setRowHeight(1, 60);
  sh.setRowHeight(2, 30);
  sh.getRange(3, 1, 1, cols).setValues([['직', '성명', '급식사유', '급식기간', '확인(영양사)']]).setBackground('#efefef').setFontWeight('bold');
  sh.setRowHeight(3, 40);
  sh.getRange(4, 1, 1, cols).setValues([[C('TUTOR_TITLE'), C('TUTOR_NAME'), C('MEAL_REASON'), period, '\n\n\n          (인)']]);
  sh.setRowHeight(4, 110);
  styleTable_(sh, 3, 1, 2, cols, { fontSize: 11 });
  sh.getRange(4, 5).setHorizontalAlignment('right').setVerticalAlignment('bottom');
  sh.setRowHeight(5, 40);
  sh.getRange(6, 1, 1, cols).merge().setValue('신청일 :  ' + lastDay.getFullYear() + '. ' + (lastDay.getMonth() + 1) + '. ' + lastDay.getDate() + '.')
    .setHorizontalAlignment('center').setFontSize(12);
  sh.setRowHeight(6, 40);
  sh.setRowHeight(7, 60);
  sh.getRange(8, 1, 1, cols).merge().setValue('신청인:  ' + C('TUTOR_NAME') + '          (서명)')
    .setHorizontalAlignment('right').setFontSize(12);
  [80, 80, 90, 320, 120].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  return sh;
}

// ───────────────────────── PDF / 메일 ─────────────────────────
function exportSheetPdf_(sh) {
  var ss = sh.getParent();
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf&gid=' + sh.getSheetId() +
    '&size=A4&portrait=true&fitw=true&scale=4&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED' +
    '&top_margin=0.6&bottom_margin=0.6&left_margin=0.5&right_margin=0.5&horizontal_alignment=CENTER&vertical_alignment=TOP';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('PDF 내보내기 실패(' + resp.getResponseCode() + '): ' + sh.getName());
  return resp.getBlob().setName(sh.getName() + '.pdf');
}

function emailMonthlyDocs(year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kinds = ['출근기록부', '활동일지', '급식확인서'];
  var blobs = [];
  kinds.forEach(function (k) {
    var sh = ss.getSheetByName(docSheetName_(k, month));
    if (sh) blobs.push(exportSheetPdf_(sh));
  });
  if (!blobs.length) throw new Error(month + '월 서류 시트가 없습니다. 먼저 생성하세요.');
  MailApp.sendEmail({
    to: C('OWNER_EMAIL'),
    subject: '[디지털튜터] ' + year + '년 ' + month + '월 제출 서류 (출근기록부/활동일지/급식확인서)',
    body: year + '년 ' + month + '월 제출 서류 PDF ' + blobs.length + '장을 첨부합니다.\n인쇄 후 서명해서 제출하세요.\n\n시트 열기: ' + ss.getUrl(),
    attachments: blobs
  });
}

// ───────────────────────── 유틸 ─────────────────────────
function dateKey_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function timeStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  if (typeof v === 'number') { var mins = Math.round((v % 1) * 24 * 60); return ('0' + Math.floor(mins / 60)).slice(-2) + ':' + ('0' + (mins % 60)).slice(-2); }
  var m = String(v || '').match(/(\d{1,2})\s*:\s*(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}
function toMin_(t) { var m = String(t).match(/(\d{1,2}):(\d{2})/); return m ? (+m[1]) * 60 + (+m[2]) : 0; }

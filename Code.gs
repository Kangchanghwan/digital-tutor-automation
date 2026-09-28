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
 * 근무일정 연동: 월별 요청 시트의 날짜/교시 틀은 "근무일정" 탭(+"교시표" 탭) 기준으로 자동 생성된다.
 *   근무일에만 날짜가 나오고, 근무시간 밖 교시는 회색(요청 불가). 근무일정을 고치면 요청표가 즉시 따라 바뀐다.
 *   → 아래 "근무일정 ↔ 월별 요청표 연동" 섹션 참고
 *
 * 서명: "설정" 탭 SIGNATURE 칸에 서명 이미지를 넣으면 서류의 (서명) 자리에 자동으로 찍힌다. (아래 "서명 이미지" 섹션)
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
  ['MEAL_CUTOFF',      '13:00',                '급식 포함 기준: 이 시각 이후에 퇴근하는 날만 급식'],
  ['PERIOD_TOLERANCE_MIN', '10',                '교시가 근무시간을 이 분(分)만큼 벗어나도 요청 가능으로 표시 (예: 1교시 08:55 시작 / 출근 09:00)'],
  ['SIGNATURE',        '',                     '서명 이미지: 이 행의 "값" 칸 선택 → 삽입 > 이미지 > 셀에 이미지 삽입. 서류의 (서명) 자리에 자동으로 찍힘. 비우면 직접 서명']
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
          if (k && r[1] !== '' && r[1] !== null && (typeof r[1] !== 'object' || r[1] instanceof Date)) configCache_[k] = (r[1] instanceof Date) ? Utilities.formatDate(r[1], TZ, 'HH:mm') : String(r[1]).trim();
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
  styleSignatureRow_(sh);
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
    .addItem('근무일정 → 요청표 맞추기 (지금 실행)', 'syncRequestSheetsFromWork')
    .addSeparator()
    .addItem('월별 서류 생성 (출근기록부/활동일지/급식확인서)', 'generateMonthlyDocsPrompt')
    .addItem('월별 서류 PDF 메일로 받기', 'emailMonthlyDocsPrompt')
    .addItem('서명 이미지 확인', 'checkSignature')
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
    '3) 월별 요청 시트(예: 2026년 09월)는 근무일정 기준으로 자동으로 만들어집니다. 날짜를 직접 넣지 마세요.\n' +
    '   (근무일정을 고치면 요청표가 따라 바뀌고, 교시 시간은 "교시표" 탭에서 바꿉니다)\n' +
    '4) 메뉴 > "트리거 설치 (최초 1회)" 를 누르고 Google 권한을 승인하세요.\n' +
    '5) (선택) "' + CONFIG_SHEET_NAME + '" 탭 SIGNATURE 행의 값 칸에 서명 이미지를 넣으면 (삽입 > 이미지 > 셀에 이미지 삽입)\n' +
    '   서류의 (서명) 자리에 자동으로 찍힙니다. 메뉴 > "서명 이미지 확인" 으로 점검할 수 있습니다.\n' +
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
    var name = sheet.getName();
    if (name === WORK_SHEET_NAME || name === PERIOD_SHEET_NAME || name === CONFIG_SHEET_NAME) {
      // 근무일정이 바뀌면 월별 요청표 틀을 즉시 다시 맞춘다
      var lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) return;
      try {
        var res = syncRequestSheetsFromWork_({ notify: true });
        if (res.changed.length) toast_(res.message);
      } finally {
        lock.releaseLock();
      }
      return;
    }
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
    // 근무일정과 요청표 틀이 어긋났으면 먼저 맞춘다 (캘린더 동기화는 아래에서 한꺼번에)
    try { syncRequestSheetsFromWork_({ notify: true, skipSync: true }); } catch (e) { console.error('근무일정 연동 오류: ' + e); }
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
        if (!eventId || !currentDate || !times || !isSlotAvailable_(currentDate, times)) {
          // 등록 전이거나, 근무일정상 근무일/근무시간이 아닌 요청 → 확인중
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

    var offSchedule = !isSlotAvailable_(currentDate, times);   // 근무일정상 근무일/근무시간이 아님
    var title = (offSchedule ? OFF_SCHEDULE_TAG + ' ' : '') + C('TITLE_PREFIX') + ' ' + period + (teacher ? ' ' + teacher : '') + (note ? ' - ' + note : '');
    if (status && AUTO_STATUSES.indexOf(status) < 0) title = '[' + status + '] ' + title;
    var desc = [
      '요청 선생님: ' + (teacher || '-'),
      '요청사항: ' + (note || '-'),
      '처리상태: ' + (status || '-'),
      offSchedule ? '※ 근무일정상 근무일/근무시간이 아닙니다. 요청 선생님과 일정 조정이 필요합니다.' : null,
      '',
      '시트: ' + sheet.getName() + ' ' + row + '행',
      sheetUrl
    ].filter(function (x) { return x !== null; }).join('\n');

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
      created.push({ when: Utilities.formatDate(start, TZ, 'yyyy-MM-dd (E) HH:mm') + '~' + Utilities.formatDate(end, TZ, 'HH:mm'), period: period, teacher: teacher, note: note, sheet: sheet.getName(), row: row, off: offSchedule });
    }
    // 등록 완료 → 시간 기준 상태 반영
    applyStatus_(sheet, row, status, offSchedule ? STATUS_CHECKING : computeStatus_(start, end));
  }

  if (created.length) notifyCreated_(created, editor, sheetUrl);
  return completed;
}

function notifyCreated_(created, editor, sheetUrl) {
  var lines = created.map(function (c) {
    return '• ' + c.when + ' ' + c.period + (c.teacher ? ' / ' + c.teacher + ' 선생님' : '') + (c.note ? ' / ' + c.note : '') +
      (c.off ? '  ※ 근무일정 확인 필요(근무일/근무시간 아님)' : '');
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
//    (비고 = 근무일정표 원문(서류에 안 나감), 활동 = 서류에 표기할 특이 활동. 예: 동아리 활동, 축제 지원)
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
  configCache_ = null;
  signatureCache_ = undefined;      // 설정 탭의 서명 이미지를 매번 새로 읽음 (사람이 바뀌면 바로 반영)
  ensureConfigKeys_();
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
      '\n근무 ' + rows.length + '일 / 총 ' + totalHours + '시간' +
      '\n서명: ' + (getSignatureBlob_() ? '설정 탭 이미지로 자동 입력' : '없음 (설정 탭 SIGNATURE 칸이 비어 있음)')
  };
}

function resetDocSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) {
    sh.clear(); sh.clearFormats(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    sh.getImages().forEach(function (im) { im.remove(); });     // 이전 서명 이미지 제거
  }
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
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontSize(10);
  sh.setRowHeight(2, 48);            // 서명 이미지가 들어갈 높이
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
  placeSignature_(sh, 2, cols);       // 작성자 (서명)
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
  placeSignature_(sh, 2, cols);       // 작성자 (서명)
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
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontSize(12);
  sh.setRowHeight(8, 52);
  [80, 80, 90, 320, 120].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  placeSignature_(sh, 8, cols);       // 신청인 (서명)
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


// ═══════════════════════════════════════════════════════════════════════
//  근무일정 ↔ 월별 요청표 연동
//  - 월별 요청 시트(2026년 09월 ...)의 날짜/교시 틀을 "근무일정" 탭 기준으로 자동으로 만든다
//    · 근무일에만 날짜 블록이 생긴다 (근무가 없는 날은 아예 안 보임)
//    · 근무시간 밖 교시는 회색 + "근무시간 외" 표시, 요청 입력이 막힌다
//    · 마지막 교시 뒤로 30분 이상 더 근무하는 날은 "방과후" 칸이 추가된다
//  - "근무일정" / "교시표" / "설정" 탭을 고치면 즉시 다시 맞추고, 매시간 한 번 더 점검한다
//  - 이미 들어온 요청은 (날짜, 교시) 기준으로 그대로 옮겨 담는다 (캘린더 일정도 유지)
//  - 근무일/근무시간이 바뀌어 들어갈 자리가 없어진 요청은 지우지 않고
//    표 맨 아래 "근무일정과 맞지 않는 요청" 구역으로 옮겨 "확인중"으로 표시 + 메일 알림
// ═══════════════════════════════════════════════════════════════════════
var PERIOD_SHEET_NAME = '교시표';
var PERIOD_DEFAULTS = [
  ['1교시', '08:55', '09:40'],
  ['2교시', '09:50', '10:35'],
  ['3교시', '10:45', '11:30'],
  ['4교시', '11:40', '12:25'],
  ['점심시간', '12:25', '13:25'],
  ['5교시', '13:25', '14:10'],
  ['6교시', '14:20', '15:05']
];
var AFTER_SCHOOL_LABEL = '방과후';
var AFTER_SCHOOL_MIN_GAP = 30;                     // 마지막 교시 종료 후 30분 이상 근무하면 "방과후" 칸 추가
var OFF_HOURS_MARK = '근무시간 외';                  // 회색 칸의 처리상태 칸 표시
var OFF_SCHEDULE_TAG = '[근무일정 확인]';           // 근무일/근무시간이 아닌 요청의 캘린더 제목 앞말
var ORPHAN_HEADER = '⚠ 근무일정과 맞지 않는 요청 : 근무일 또는 근무시간이 바뀌어 원래 칸이 없어진 요청입니다. 요청 선생님과 일정을 조정한 뒤 이 행의 요청 선생님/비고를 지우면 정리됩니다.';
var MONTH_SHEET_RE = /^(\d{4})년\s*(\d{1,2})월$/;
var LAYOUT_PROTECT_DESC = '근무일정 연동 (자동 생성 영역)';
var REQUEST_HEADERS = ['날짜', '근무시간', '요일', '교시', '시간', '요청 선생님', '비고 (요청사항)', '처리상태', EVENT_ID_HEADER];
var STATUS_LIST_ORDER = [STATUS_WAITING, STATUS_ONGOING, STATUS_DONE, STATUS_HOLD, STATUS_CHECKING];
var LOOK = {
  font: 'Malgun Gothic', navy: '#1f3864', red: '#ff0000', titleBg: '#dce3f0', white: '#ffffff', band: '#f2f5fa',
  input: '#fffde7', lunch: '#ededed', lunchFont: '#595959', off: '#e0e0e0', offFont: '#9e9e9e',
  grid: '#bfbfbf', blockLine: '#8ea9db', orphanHead: '#f4cccc', orphanHeadFont: '#990000', orphanBg: '#fce8e6'
};

/** 메뉴/편집기: 근무일정 기준으로 월별 요청표 다시 맞추기 */
function syncRequestSheetsFromWork() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    var r = syncRequestSheetsFromWork_({ notify: false, force: false });
    console.log(r.message);
    toast_(r.message);
  } finally {
    lock.releaseLock();
  }
}

/** 틀이 이미 맞아도 전부 새로 그리기 (서식이 깨졌을 때) */
function rebuildAllRequestSheets() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    var r = syncRequestSheetsFromWork_({ notify: false, force: true });
    console.log(r.message);
    toast_(r.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 핵심: 근무일정에 있는 달마다 요청표를 계산해서, 현재 시트와 다르면 다시 그린다.
 * opts.notify  : 새로 "맞지 않는 요청"이 생기면 메일 알림
 * opts.skipSync: 다시 그린 뒤 캘린더 동기화 생략 (syncAll 안에서 부를 때)
 * opts.force   : 틀이 같아도 다시 그리기
 * (잠금 없이 동작 — 호출하는 쪽에서 LockService 잡을 것)
 */
function syncRequestSheetsFromWork_(opts) {
  opts = opts || {};
  workMapCache_ = null;
  configCache_ = null;
  ensureConfigKeys_();
  ensurePeriodSheet_();
  var wm = getWorkMap_();
  var periods = getPeriods_();
  var tol = periodTolerance_();

  // 근무일정 → 달별 근무일 목록
  var months = {};
  Object.keys(wm.days).forEach(function (k) {
    var w = wm.days[k];
    var ym = w.date.getFullYear() * 100 + (w.date.getMonth() + 1);
    (months[ym] = months[ym] || []).push(w);
  });

  var changed = [], orphanTotal = 0, newOrphans = [];
  Object.keys(months).map(Number).sort(function (a, b) { return a - b; }).forEach(function (ym) {
    var year = Math.floor(ym / 100), month = ym % 100;
    var days = months[ym].sort(function (a, b) { return a.date - b.date; });
    var sheet = getOrCreateMonthSheet_(year, month);
    var cur = readRequestSheet_(sheet);
    var layout = buildMonthLayout_(days, periods, cur.records, tol);
    orphanTotal += layout.orphans.length;
    var wantSig = layout.rows.map(function (r) { return rowSig_(rowValues_(r)); });
    var same = !opts.force && sameList_(cur.sig, wantSig) &&
      String(sheet.getRange(1, 1).getDisplayValue()) === monthTitle_(year, month).join('') &&
      sheet.getRange(HEADER_ROW, 1, 1, REQUEST_HEADERS.length).getValues()[0].join('|') === REQUEST_HEADERS.join('|');
    if (same) return;

    writeMonthSheet_(sheet, year, month, layout);
    changed.push(month + '월(근무 ' + days.length + '일' + (layout.orphans.length ? ', 조정 필요 ' + layout.orphans.length + '건' : '') + ')');
    layout.orphans.forEach(function (o) { newOrphans.push({ sheet: sheet, o: o }); });
    if (!opts.skipSync) {
      var last = sheet.getLastRow();
      if (last >= FIRST_DATA_ROW) syncRows_(sheet, FIRST_DATA_ROW, last, '');
    }
  });

  if (opts.notify && newOrphans.length) notifyOrphans_(newOrphans);

  var message = changed.length
    ? '근무일정 기준으로 요청표를 맞췄습니다: ' + changed.join(', ')
    : '요청표가 이미 근무일정과 일치합니다.';
  if (orphanTotal) message += '\n근무일정과 맞지 않는 요청 ' + orphanTotal + '건은 각 달 표 맨 아래에 모아 두었습니다.';
  return { changed: changed, orphanTotal: orphanTotal, message: message };
}

// ───────────────────────── 근무일정 / 교시표 읽기 ─────────────────────────
var workMapCache_ = null;
/** 근무일정 탭 → { days: {'2026-09-01': {date, start, end, activity}}, months: {'2026-9': true} } */
function getWorkMap_() {
  if (workMapCache_) return workMapCache_;
  var map = { days: {}, months: {} };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WORK_SHEET_NAME);
  if (sh && sh.getLastRow() >= 2) {
    var values = sh.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      var d = parseDate_(r[0]);
      var s = timeStr_(r[2]), e = timeStr_(r[3]);
      if (!d || !s || !e) continue;
      var day = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      map.days[dateKey_(day)] = { date: day, start: s, end: e, activity: String(r[6] || '').trim() };
      map.months[day.getFullYear() + '-' + (day.getMonth() + 1)] = true;
    }
  }
  workMapCache_ = map;
  return map;
}

function periodTolerance_() {
  var t = parseInt(C('PERIOD_TOLERANCE_MIN'), 10);
  return isNaN(t) ? 10 : t;
}

/** 교시가 근무시간 안에 들어가는지 (앞뒤로 허용 오차 tol분) */
function periodFits_(startMin, endMin, work, tol) {
  return startMin >= toMin_(work.start) - tol && endMin <= toMin_(work.end) + tol;
}

/**
 * 이 날짜/시간의 요청을 디지털튜터가 받을 수 있는지 (캘린더·처리상태 판단용)
 * 근무일정이 아예 입력되지 않은 달은 제한하지 않는다.
 */
function isSlotAvailable_(date, times) {
  var wm = getWorkMap_();
  if (!wm.months[date.getFullYear() + '-' + (date.getMonth() + 1)]) return true;
  var w = wm.days[dateKey_(date)];
  if (!w) return false;
  return periodFits_(times.sh * 60 + times.sm, times.eh * 60 + times.em, w, periodTolerance_());
}

/** 교시표 탭이 없으면 기본 교시로 만든다 */
function ensurePeriodSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(PERIOD_SHEET_NAME);
  if (sh) return sh;
  var work = ss.getSheetByName(WORK_SHEET_NAME);
  sh = ss.insertSheet(PERIOD_SHEET_NAME, work ? work.getIndex() : ss.getSheets().length);
  var rows = [['교시', '시작', '종료']].concat(PERIOD_DEFAULTS);
  sh.getRange(1, 1, rows.length, 3).setNumberFormat('@').setValues(rows)
    .setHorizontalAlignment('center').setBorder(true, true, true, true, true, true, LOOK.grid, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground(LOOK.navy).setFontColor(LOOK.white);
  sh.getRange(2, 2, rows.length - 1, 2).setBackground('#fff8dc');
  sh.getRange(1, 5, 4, 1).setValues([
    ['안내'],
    ['· 여기 시간을 바꾸면 모든 월별 요청표의 교시/시간이 자동으로 바뀝니다.'],
    ['· 행을 추가하면 교시가 늘어납니다 (예: 7교시 15:15 16:00). 시간은 24시간제 HH:MM.'],
    ['· 근무시간 밖 교시는 요청표에서 회색(요청 불가)으로 표시됩니다. 허용 오차는 설정 탭 PERIOD_TOLERANCE_MIN.']
  ]);
  sh.getRange(1, 5).setFontWeight('bold');
  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 80); sh.setColumnWidth(3, 80); sh.setColumnWidth(5, 560);
  sh.setFrozenRows(1);
  return sh;
}

/** 교시표 → [{label, start, end}] (시작 시각 순) */
function getPeriods_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PERIOD_SHEET_NAME);
  var src = (sh && sh.getLastRow() >= 2) ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues() : PERIOD_DEFAULTS;
  var list = [];
  src.forEach(function (r) {
    var label = String(r[0] || '').trim(), s = timeStr_(r[1]), e = timeStr_(r[2]);
    if (label && s && e) list.push({ label: label, start: s, end: e });
  });
  if (!list.length) list = PERIOD_DEFAULTS.map(function (p) { return { label: p[0], start: p[1], end: p[2] }; });
  list.sort(function (a, b) { return toMin_(a.start) - toMin_(b.start); });
  return list;
}

/** 설정 탭에 새로 생긴 항목(PERIOD_TOLERANCE_MIN 등)이 없으면 아래에 추가 */
function ensureConfigKeys_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sh) { ensureConfigSheet_(); return; }
  var have = {};
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = true; });
  var missing = CONFIG_DEFAULTS.filter(function (d) { return !have[d[0]]; });
  if (!missing.length) return;
  var start = sh.getLastRow() + 1;
  sh.getRange(start, 2, missing.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, missing.length, 3).setValues(missing);
  sh.getRange(start, 2, missing.length, 1).setBackground('#fff8dc');
  styleSignatureRow_(sh);
  configCache_ = null;
}

// ───────────────────────── 현재 요청표 읽기 ─────────────────────────
/** 월별 요청 시트 → { records: 요청이 있는 행들, sig: 행별 틀 서명 } */
function readRequestSheet_(sheet) {
  var last = sheet.getLastRow();
  var out = { records: [], sig: [] };
  if (last < FIRST_DATA_ROW) return out;
  var values = sheet.getRange(FIRST_DATA_ROW, 1, last - FIRST_DATA_ROW + 1, COL.EVENT_ID).getValues();
  var cur = null;
  values.forEach(function (r) {
    var d = parseDate_(r[COL.DATE - 1]);
    if (d) cur = d;
    out.sig.push(rowSig_(r));
    var teacher = String(r[COL.TEACHER - 1] || '').trim();
    var note = String(r[COL.NOTE - 1] || '').trim();
    var eventId = String(r[COL.EVENT_ID - 1] || '').trim();
    if (!teacher && !note && !eventId) return;
    if (!cur) return;
    out.records.push({
      date: new Date(cur.getFullYear(), cur.getMonth(), cur.getDate()),
      dateKey: dateKey_(cur),
      period: String(r[COL.PERIOD - 1] || '').trim(),
      time: String(r[COL.TIME - 1] || '').trim(),
      teacher: teacher, note: note, eventId: eventId,
      status: String(r[COL.STATUS - 1] || '').trim()
    });
  });
  while (out.sig.length && out.sig[out.sig.length - 1] === EMPTY_SIG_) out.sig.pop();
  return out;
}

var EMPTY_SIG_ = '|||||';
/** 행의 "틀"만 비교하기 위한 서명: 날짜 | 근무시간 | 요일 | 교시 | 시간 | 근무시간외 여부 */
function rowSig_(v) {
  var a = (v[0] instanceof Date) ? dateKey_(v[0]) : String(v[0] === null || v[0] === undefined ? '' : v[0]).trim();
  var s = function (x) { return String(x === null || x === undefined ? '' : x).trim(); };
  return [a, s(v[1]), s(v[2]), s(v[3]), s(v[4]), s(v[7]) === OFF_HOURS_MARK ? 'off' : ''].join('|');
}

function sameList_(a, b) {
  if (a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

// ───────────────────────── 새 틀 계산 (순수 함수) ─────────────────────────
/**
 * workDays: [{date, start, end, activity}] (해당 월, 날짜순)
 * periods : [{label, start, end}]
 * records : readRequestSheet_ 의 records
 * 반환 rows: [{kind:'slot'|'off'|'sep'|'orphan', ...}], orphans: 자리 없는 요청들
 */
function buildMonthLayout_(workDays, periods, records, tol) {
  var byKey = {};
  records.forEach(function (rec) {
    var k = rec.dateKey + '|' + rec.period;
    (byKey[k] = byKey[k] || []).push(rec);
  });
  var lastEnd = 0;
  periods.forEach(function (p) { lastEnd = Math.max(lastEnd, toMin_(p.end)); });

  var rows = [];
  workDays.forEach(function (w, dayIdx) {
    var slots = periods.map(function (p) { return { label: p.label, start: p.start, end: p.end }; });
    if (toMin_(w.end) - lastEnd >= AFTER_SCHOOL_MIN_GAP) slots.push({ label: AFTER_SCHOOL_LABEL, start: minStr_(lastEnd), end: w.end });
    var dk = dateKey_(w.date);
    slots.forEach(function (s, i) {
      var avail = periodFits_(toMin_(s.start), toMin_(s.end), w, tol);
      var row = {
        kind: avail ? 'slot' : 'off', date: w.date, dateKey: dk, first: i === 0, last: i === slots.length - 1,
        blockLen: slots.length, band: dayIdx % 2 === 1, lunch: /점심/.test(s.label),
        workText: w.start + ' ~ ' + w.end + (w.activity ? '\n' + w.activity : ''),
        weekday: WEEKDAYS_KO[w.date.getDay()], period: s.label, time: s.start + ' ~ ' + s.end,
        teacher: '', note: '', eventId: '', status: avail ? STATUS_WAITING : OFF_HOURS_MARK
      };
      if (avail) {
        var list = byKey[dk + '|' + s.label];
        if (list && list.length) {
          var rec = list.shift();
          rec.placed = true;
          row.teacher = rec.teacher; row.note = rec.note; row.eventId = rec.eventId;
          row.status = DROPDOWN_STATUSES.indexOf(rec.status) >= 0 ? rec.status : STATUS_CHECKING;
        }
      }
      rows.push(row);
    });
  });

  var workByKey = {};
  workDays.forEach(function (w) { workByKey[dateKey_(w.date)] = w; });
  var orphans = records.filter(function (r) { return !r.placed; });
  orphans.sort(function (a, b) { return a.date - b.date || a.time.localeCompare(b.time); });
  if (orphans.length) {
    rows.push({ kind: 'sep' });
    orphans.forEach(function (o) {
      var w = workByKey[o.dateKey];
      rows.push({
        kind: 'orphan', date: o.date, dateKey: o.dateKey, period: o.period, time: o.time,
        workText: w ? w.start + ' ~ ' + w.end : '근무 없음', weekday: WEEKDAYS_KO[o.date.getDay()],
        teacher: o.teacher, note: o.note, eventId: o.eventId,
        status: o.status === STATUS_HOLD ? STATUS_HOLD : STATUS_CHECKING,
        reason: w ? '근무시간 외' : '근무일 아님'
      });
    });
  }
  return { rows: rows, orphans: orphans };
}

/**
 * 틀 행 → 시트에 쓸 9칸 값
 * 날짜는 '2026-09-01' 문자열로 쓴다: 시트 시간대와 스크립트 시간대가 달라도
 * 시트가 자기 시간대의 그날 자정으로 해석하므로 날짜가 하루 밀리지 않는다.
 */
function rowValues_(r) {
  if (r.kind === 'sep') return [ORPHAN_HEADER, '', '', '', '', '', '', '', ''];
  if (r.kind === 'orphan') return [r.dateKey, r.workText, r.weekday, r.period, r.time, r.teacher, r.note, r.status, r.eventId];
  return [r.first ? r.dateKey : '', r.first ? r.workText : '', r.first ? r.weekday : '', r.period, r.time, r.teacher, r.note, r.status, r.eventId];
}

function monthTitle_(year, month) {
  return [year + '년 ' + month + '월 교시별 요청사항  ', '(디지털튜터 근무일만 표시 · 회색 칸은 근무시간 외라 요청 불가)'];
}

// ───────────────────────── 시트에 그리기 ─────────────────────────
function getOrCreateMonthSheet_(year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = year + '년 ' + ('0' + month).slice(-2) + '월';
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  // 기존 이름 표기(예: "2026년 9월")도 인정
  var target = year * 100 + month, before = null, after = null;
  ss.getSheets().forEach(function (s) {
    var m = s.getName().match(MONTH_SHEET_RE);
    if (!m) return;
    var ym = (+m[1]) * 100 + (+m[2]);
    if (ym === target) sh = s;
    else if (ym < target && (!before || ym > before.ym)) before = { ym: ym, s: s };
    else if (ym > target && (!after || ym < after.ym)) after = { ym: ym, s: s };
  });
  if (sh) return sh;
  var idx = before ? before.s.getIndex() : (after ? after.s.getIndex() - 1 : ss.getSheets().length);
  sh = ss.insertSheet(name, idx);
  var ref = before ? before.s : (after ? after.s : null);
  if (ref && ref.getTabColor()) sh.setTabColor(ref.getTabColor());
  var widths = [100, 100, 50, 75, 115, 105, 380, 90];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  return sh;
}

function writeMonthSheet_(sheet, year, month, layout) {
  var rows = layout.rows, n = rows.length, W = COL.EVENT_ID;
  var BS = SpreadsheetApp.BorderStyle;

  // 1) 제목 / 헤더
  var t = monthTitle_(year, month);
  var titleRange = sheet.getRange(1, 1, 1, 8);
  titleRange.breakApart().merge();
  var base = SpreadsheetApp.newTextStyle().setFontFamily(LOOK.font).setFontSize(14).setBold(true);
  sheet.getRange(1, 1).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(t[0] + t[1])
    .setTextStyle(0, t[0].length, base.setForegroundColor(LOOK.navy).build())
    .setTextStyle(t[0].length, t[0].length + t[1].length, base.setForegroundColor(LOOK.red).build())
    .build());
  titleRange.setBackground(LOOK.titleBg).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 37);
  sheet.getRange(HEADER_ROW, 1, 1, W).setValues([REQUEST_HEADERS]);
  sheet.getRange(HEADER_ROW, 1, 1, 8).setBackground(LOOK.navy).setFontColor(LOOK.white).setFontWeight('bold')
    .setFontFamily(LOOK.font).setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, LOOK.grid, BS.SOLID);
  sheet.setRowHeight(HEADER_ROW, 32);

  // 2) 기존 본문 싹 정리 (병합/서식/검증/보호)
  var maxRows = sheet.getMaxRows();
  var need = FIRST_DATA_ROW + n + 5;
  if (maxRows < need) { sheet.insertRowsAfter(maxRows, need - maxRows); maxRows = need; }
  var body = sheet.getRange(FIRST_DATA_ROW, 1, maxRows - FIRST_DATA_ROW + 1, W);
  body.breakApart();
  body.clearDataValidations();
  body.clear();
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (p.getDescription() === LAYOUT_PROTECT_DESC) p.remove();
  });
  if (!n) return;

  // 3) 값 + 서식 (한 번에)
  var values = [], bg = [], fc = [], fw = [], ha = [], nf = [], hv = [], fgv = [];
  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_LIST_ORDER, true).setAllowInvalid(false).build();
  var offRule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied('=FALSE').setAllowInvalid(false)
    .setHelpText('디지털튜터 근무시간이 아니라 요청할 수 없습니다. (근무일정 탭 기준)').build();
  var blocks = [], sepRow = 0;
  rows.forEach(function (r, i) {
    var row = FIRST_DATA_ROW + i;
    values.push(rowValues_(r));
    nf.push(['yyyy-mm-dd', '@', '@', '@', '@', '@', '@', '@', '@']);
    ha.push(['center', 'center', 'center', 'center', 'center', 'center', 'left', 'center', 'left']);
    var left = r.band ? LOOK.band : LOOK.white;
    if (r.kind === 'sep') {
      sepRow = row;
      bg.push([LOOK.orphanHead, LOOK.orphanHead, LOOK.orphanHead, LOOK.orphanHead, LOOK.orphanHead, LOOK.orphanHead, LOOK.orphanHead, LOOK.orphanHead, null]);
      fc.push(fill9_(LOOK.orphanHeadFont)); fw.push(fill9_('bold'));
      hv.push([null]); fgv.push([null, null]);
      return;
    }
    if (r.kind === 'orphan') {
      bg.push([LOOK.orphanBg, LOOK.orphanBg, LOOK.orphanBg, LOOK.orphanBg, LOOK.orphanBg, LOOK.input, LOOK.input, LOOK.input, null]);
      fc.push(fill9_('#000000')); fw.push(fill9_('normal'));
      hv.push([statusRule]); fgv.push([null, null]);
      return;
    }
    if (r.first) blocks.push({ row: row, len: r.blockLen });
    if (r.kind === 'off') {
      bg.push([left, left, left, LOOK.off, LOOK.off, LOOK.off, LOOK.off, LOOK.off, null]);
      fc.push(['#000000', '#000000', '#000000', LOOK.offFont, LOOK.offFont, LOOK.offFont, LOOK.offFont, LOOK.offFont, '#000000']);
      fw.push(fill9_('normal'));
      hv.push([null]); fgv.push([offRule, offRule]);
      return;
    }
    var de = r.lunch ? LOOK.lunch : left;
    bg.push([left, left, left, de, de, LOOK.input, LOOK.input, LOOK.input, null]);
    fc.push(['#000000', '#000000', '#000000', r.lunch ? LOOK.lunchFont : '#000000', r.lunch ? LOOK.lunchFont : '#000000', '#000000', '#000000', '#000000', '#000000']);
    fw.push(['normal', 'normal', 'normal', r.lunch ? 'bold' : 'normal', r.lunch ? 'bold' : 'normal', 'normal', 'normal', 'normal', 'normal']);
    hv.push([statusRule]); fgv.push([null, null]);
  });

  var rng = sheet.getRange(FIRST_DATA_ROW, 1, n, W);
  rng.setNumberFormats(nf);
  rng.setValues(values);
  rng.setBackgrounds(bg).setFontColors(fc).setFontWeights(fw).setHorizontalAlignments(ha)
    .setVerticalAlignment('middle').setFontFamily(LOOK.font).setFontSize(11);
  sheet.getRange(FIRST_DATA_ROW, 2, n, 1).setWrap(true).setFontSize(10);          // 근무시간
  sheet.getRange(FIRST_DATA_ROW, COL.NOTE, n, 1).setWrap(true);                  // 비고
  sheet.getRange(FIRST_DATA_ROW, 1, n, 8).setBorder(true, true, true, true, true, true, LOOK.grid, BS.SOLID);
  sheet.setRowHeights(FIRST_DATA_ROW, n, 26);

  // 4) 날짜 블록 병합 + 블록 구분선
  blocks.forEach(function (b) {
    sheet.getRange(b.row, 1, b.len, 3).mergeVertically();
    sheet.getRange(b.row + b.len - 1, 1, 1, 8).setBorder(null, null, true, null, null, null, LOOK.blockLine, BS.SOLID_MEDIUM);
  });
  if (sepRow) {
    sheet.getRange(sepRow, 1, 1, 8).merge().setWrap(true).setHorizontalAlignment('left').setFontSize(10);
    sheet.setRowHeight(sepRow, 44);
  }

  // 5) 입력 규칙: 처리상태 드롭다운 / 근무시간 외 칸 입력 차단
  sheet.getRange(FIRST_DATA_ROW, COL.STATUS, n, 1).setDataValidations(hv);
  sheet.getRange(FIRST_DATA_ROW, COL.TEACHER, n, 2).setDataValidations(fgv);

  // 6) 날짜~시간(A~E)은 자동 생성 영역: 실수로 고치려 하면 경고
  sheet.getRange(HEADER_ROW, 1, n + 1, 5).protect().setDescription(LAYOUT_PROTECT_DESC).setWarningOnly(true);

  // 7) 열 모양
  if (sheet.getColumnWidth(2) < 90) sheet.setColumnWidth(2, 100);
  sheet.hideColumns(COL.EVENT_ID);
  sheet.setFrozenRows(HEADER_ROW);
}

function fill9_(v) { return [v, v, v, v, v, v, v, v, v]; }

/** 새로 "맞지 않는 요청"이 된 건만 메일 (같은 건은 한 번만) */
function notifyOrphans_(items) {
  var props = PropertiesService.getDocumentProperties();
  var seen = [];
  try { seen = JSON.parse(props.getProperty('NOTIFIED_ORPHANS') || '[]'); } catch (e) { seen = []; }
  var fresh = items.filter(function (it) {
    var k = it.o.dateKey + '|' + it.o.period + '|' + it.o.teacher + '|' + it.o.note;
    it.key = k;
    return seen.indexOf(k) < 0;
  });
  if (!fresh.length) return;
  var lines = fresh.map(function (it) {
    var o = it.o;
    return '• ' + o.dateKey + ' (' + WEEKDAYS_KO[o.date.getDay()] + ') ' + o.period + ' ' + o.time +
      (o.teacher ? ' / ' + o.teacher + ' 선생님' : '') + (o.note ? ' / ' + o.note : '') + '  → ' + (o.reason || '확인 필요');
  });
  var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  try {
    MailApp.sendEmail(C('OWNER_EMAIL'), '[디지털튜터 요청] 근무일정과 맞지 않는 요청 ' + fresh.length + '건',
      '근무일정이 바뀌어 아래 요청이 들어갈 칸이 없어졌습니다.\n각 달 요청표 맨 아래 "근무일정과 맞지 않는 요청" 구역에 모아 두었으니, 요청 선생님과 일정을 조정해 주세요.\n\n' +
      lines.join('\n') + '\n\n시트 열기: ' + url);
  } catch (e) { console.error('메일 전송 실패: ' + e); }
  seen = seen.concat(fresh.map(function (it) { return it.key; })).slice(-300);
  props.setProperty('NOTIFIED_ORPHANS', JSON.stringify(seen));
}

function toast_(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, '디지털튜터 캘린더', 10); } catch (e) {}
}
function minStr_(mins) { return ('0' + Math.floor(mins / 60)).slice(-2) + ':' + ('0' + (mins % 60)).slice(-2); }


// ═══════════════════════════════════════════════════════════════════════
//  서명 이미지
//  - "설정" 탭 SIGNATURE 행의 값 칸에 서명 이미지를 넣어 두면
//    (그 칸 선택 → 삽입 > 이미지 > 셀에 이미지 삽입)
//    출근기록부 / 활동일지 / 급식확인서를 만들 때 (서명) 자리에 자동으로 찍힌다.
//  - 사람이 바뀌면 설정 탭의 이미지만 바꾸면 된다. 비워 두면 (서명) 칸만 남아 직접 서명.
//  - 이미지 대신 이미지 주소(https://..., 공개된 Google Drive 링크)를 적어도 된다.
//  - 배경이 투명한 PNG가 가장 자연스럽다 (글자 "(서명)" 위에 겹쳐 찍힘).
// ═══════════════════════════════════════════════════════════════════════
var SIGNATURE_KEY = 'SIGNATURE';
var SIGNATURE_HEIGHT = 34;       // 서류에 찍히는 서명 높이(px)
var signatureCache_;             // undefined = 아직 안 읽음, null = 서명 없음

/** 메뉴: 설정 탭의 서명 이미지를 읽을 수 있는지 확인 */
function checkSignature() {
  signatureCache_ = undefined;
  ensureConfigKeys_();
  var b = getSignatureBlob_();
  var msg = b
    ? '서명 이미지 확인됨 (' + b.getContentType() + ', ' + Math.round(b.getBytes().length / 102.4) / 10 + 'KB). 서류를 만들면 (서명) 자리에 들어갑니다.'
    : '설정 탭 SIGNATURE 칸에 서명 이미지가 없습니다. 그 칸을 선택하고 삽입 > 이미지 > 셀에 이미지 삽입으로 넣어 주세요.';
  console.log(msg);
  toast_(msg);
  return msg;
}

/** 설정 탭에서 key 가 있는 행 번호 (없으면 0) */
function findConfigRow_(key, sh) {
  sh = sh || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return 0;
  var keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) if (String(keys[i][0]).trim() === key) return i + 2;
  return 0;
}

/** 설정 탭 SIGNATURE 행: 이미지가 잘 보이도록 행 높이 확보 */
function styleSignatureRow_(sh) {
  var row = findConfigRow_(SIGNATURE_KEY, sh);
  if (!row) return;
  sh.setRowHeight(row, 72);
  sh.getRange(row, 1, 1, 3).setVerticalAlignment('middle');
  sh.getRange(row, 2).setHorizontalAlignment('center');
}

/** 설정 탭의 서명 → 이미지 Blob (없으면 null). 한 번 실행 안에서는 캐시 */
function getSignatureBlob_() {
  if (signatureCache_ !== undefined) return signatureCache_;
  signatureCache_ = null;
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
    var row = sh ? findConfigRow_(SIGNATURE_KEY, sh) : 0;
    if (row) signatureCache_ = imageValueToBlob_(sh.getRange(row, 2).getValue());
  } catch (e) {
    console.error('서명 이미지 읽기 실패: ' + e);
  }
  return signatureCache_;
}

/** 셀 값(셀 이미지 / 이미지 주소) → Blob */
function imageValueToBlob_(v) {
  if (v === null || v === undefined || v === '') return null;
  var urls = [];
  if (typeof v === 'object' && typeof v.getContentUrl === 'function') {   // 셀에 삽입한 이미지 (CellImage)
    try { urls.push(v.getContentUrl()); } catch (e) {}
    try { urls.push(v.getUrl()); } catch (e) {}
  } else {
    urls.push(String(v).trim());
  }
  for (var i = 0; i < urls.length; i++) {
    var b = urlToImageBlob_(urls[i]);
    if (b) return b.setName('signature');
  }
  return null;
}

function urlToImageBlob_(url) {
  if (!url) return null;
  url = String(url);
  var m = url.match(/^data:(image\/[\w.+-]+);base64,(.+)$/);
  if (m) return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1]);
  if (!/^https?:\/\//.test(url)) return null;
  var drive = url.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:export=\w+&)?id=)([\w-]{20,})/);
  if (drive) url = 'https://drive.google.com/uc?export=download&id=' + drive[1];
  // 로그인 토큰은 Google 주소에만 붙인다 (외부 주소로 토큰이 새지 않도록)
  var isGoogle = /^https:\/\/([\w-]+\.)*(googleusercontent|google)\.com\//.test(url);
  var tries = isGoogle ? [{ headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }, {}] : [{}];
  for (var i = 0; i < tries.length; i++) {
    try {
      var opt = tries[i];
      opt.muteHttpExceptions = true;
      opt.followRedirects = true;
      var res = UrlFetchApp.fetch(url, opt);
      if (res.getResponseCode() !== 200) continue;
      var blob = res.getBlob();
      if (/^image\//.test(String(blob.getContentType()))) return blob;
    } catch (e) {}
  }
  return null;
}

/**
 * 서류의 (서명) 자리에 서명 이미지를 겹쳐 넣는다.
 * row 행의 1~lastCol 칸 오른쪽 끝(= 오른쪽 정렬된 "(서명)" 글자 위치)에 맞춰 배치.
 * 서명이 없으면 아무것도 하지 않음 → (서명) 칸에 직접 서명.
 */
function placeSignature_(sh, row, lastCol) {
  var blob = getSignatureBlob_();
  if (!blob) return false;
  try {
    var img = sh.insertImage(blob.copyBlob(), lastCol, row);
    var iw = img.getInherentWidth(), ih = img.getInherentHeight();
    var h = SIGNATURE_HEIGHT, w = Math.max(1, Math.round(iw * h / ih));
    var maxW = SIGNATURE_HEIGHT * 3;                      // 가로로 긴 서명도 과하게 넓어지지 않게
    if (w > maxW) { w = maxW; h = Math.max(1, Math.round(ih * w / iw)); }
    img.setWidth(w).setHeight(h);
    // 오른쪽 끝에서 w 만큼 왼쪽이 이미지 시작점. 좁은 칸이면 앞 칸에 걸쳐 놓는다.
    var right = 0;
    for (var c = 1; c <= lastCol; c++) right += sh.getColumnWidth(c);
    var left = Math.max(0, right - w - 2);
    var col = 1, acc = 0;
    while (col < lastCol && acc + sh.getColumnWidth(col) <= left) { acc += sh.getColumnWidth(col); col++; }
    img.setAnchorCell(sh.getRange(row, col))
      .setAnchorCellXOffset(Math.round(left - acc))
      .setAnchorCellYOffset(Math.max(0, Math.round((sh.getRowHeight(row) - h) / 2)));
    img.setAltTextTitle('서명');
    return true;
  } catch (e) {
    console.error('서명 넣기 실패(' + sh.getName() + '): ' + e);
    return false;
  }
}

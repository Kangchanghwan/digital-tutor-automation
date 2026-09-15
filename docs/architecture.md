# 아키텍처와 데이터 흐름

## 1. 구성 요소

| 구성 | 위치 | 역할 |
|---|---|---|
| 월별 요청 시트 (`2026년 09월` ...) | Google Sheets | 선생님들이 교시별로 요청을 적는 곳. 시스템의 단일 진실 원천 |
| `근무일정` 탭 | Google Sheets | 날짜 / 출근 / 퇴근 / 실근무시간 / 비고 / 활동. 서류의 출퇴근 시간과 급식일 판정 근거 |
| `설정` 탭 | Google Sheets | 성명, 이메일, 학교명, 기본 근무장소, 급식 기준 시각 등 11개 항목 |
| `Code.gs` | Apps Script (시트 바인딩) | 모든 로직. 약 700줄 |
| Google Calendar | 사용자 기본 캘린더 | 요청 1건 = 일정 1개 |
| Gmail | MailApp | 새 요청 알림, 월별 서류 PDF 발송 |

## 2. 트리거와 흐름

```mermaid
sequenceDiagram
    participant T as 선생님
    participant S as 요청 시트
    participant E as onEditInstallable
    participant C as Calendar
    participant M as Gmail
    T->>S: F/G열에 요청 입력
    S-->>E: onEdit(e.range)
    E->>S: 해당 행 읽기 (날짜 병합셀 → 위로 탐색)
    alt I열에 이벤트ID 없음
        E->>C: createEvent(title, start, end)
        C-->>E: eventId
        E->>S: I열에 eventId 기록, H열 상태 계산
        E->>M: "새 요청 N건 등록" 메일
    else 이벤트ID 있음
        E->>C: getEventById → setTitle / setTime / setDescription
    end
    alt F/G열이 비워짐
        E->>C: deleteEvent
        E->>S: I열 비움, 상태 대기로
    end
```

시간 트리거:

| 함수 | 주기 | 하는 일 |
|---|---|---|
| `updateStatuses` | 5분 | 캘린더 호출 없이 시트만 읽어 상태 판정. `확인중 / 대기 / 진행중 / 완료`. `보류`는 건드리지 않음 |
| `syncAll` | 1시간 | 모든 요청 시트를 처음부터 재동기화. 놓친 수정/외부 삭제 복구 |
| `monthlyDocsJob` | 매월 1일 08:00 | 전월 서류 3종 생성 → PDF 3장 메일 |

## 3. 시트 파싱 규칙

- 요청 시트 판별: 2행 A열 값이 `날짜`인 시트만. 서류 탭·설정 탭·근무일정 탭은 자동 제외
- 날짜 열(A)은 하루치 7행이 병합되어 있어 값이 첫 행에만 존재 → 위에서 아래로 스캔하며 `currentDate`를 유지
- 시간 열(E)은 `08:55 ~ 09:40` 문자열 → 정규식 `(\d{1,2}):(\d{2})\s*[~\-–]\s*(\d{1,2}):(\d{2})`
- 날짜 셀은 Date / 시리얼 넘버 / `2026-09-14` 문자열 세 가지 모두 허용 (`parseDate_`)
- 요청 여부: F(요청 선생님) 또는 G(비고) 중 하나라도 비어 있지 않으면 요청

## 4. 멱등성과 복구

| 상황 | 처리 |
|---|---|
| 같은 행을 여러 번 수정 | I열의 eventId로 기존 일정 찾아 업데이트. 중복 생성 없음 |
| 캘린더에서 일정을 직접 지움 | `getEventById`가 null → 다음 동기화에서 재생성 |
| 요청을 지움 | 일정 삭제 + I열 비움 + 상태 `대기` |
| 동시 실행 (onEdit 연타 + 시간 트리거) | `LockService.getScriptLock()`으로 직렬화 |
| 6분 실행 제한 초과 | 4.5분 예산 소진 시 중단, `after(60s)` 일회성 트리거로 이어서 실행. 완료된 행은 I열에 남아 있으므로 재실행 시 건너뜀 |
| 스크립트가 H열을 바꾼 것 | 설치형 onEdit은 스크립트 편집에는 발화하지 않음 → 무한 루프 없음 |

## 5. 월별 서류 생성

```
근무일정(해당 월) ──┐
                    ├─→ 날짜별 rows[{work, content}] ──→ 출근기록부 / 활동일지 / 급식확인서 시트
요청 시트(해당 월) ─┘         │
                              └─ buildDayContent_():
                                   classes: "2-2, 2-10, 2-9"      ← G열에서 "N학년 M반" 정규식 추출
                                   detail : "1교시 2-2(○○○), ..."  ← 활동일지 교과/지도내용
                                   brief  : "2-2, 2-10 수업지원"   ← 출근기록부 비고
```

- 급식확인서: `출근 < 급식기준(13:00) < 퇴근`인 날만 포함
- 서류 시트는 매번 지우고 다시 그림 (`resetDocSheet_`). 같은 달을 여러 번 생성해도 안전
- PDF: `https://docs.google.com/spreadsheets/d/{id}/export?format=pdf&gid={gid}&size=A4&portrait=true&fitw=true&gridlines=false` + `Authorization: Bearer ScriptApp.getOAuthToken()`

## 6. 필요한 OAuth 스코프

| 스코프 | 용도 |
|---|---|
| `spreadsheets` | 시트 읽기/쓰기 |
| `calendar` | 일정 CRUD |
| `script.send_mail` | 알림 및 PDF 메일 |
| `script.external_request` | PDF export URL 호출 |
| `script.scriptapp` | 트리거 설치/정리 |

## 7. 알려진 제약

- Sheets 쿠키 기반 외부 읽기(비공식 API)는 병합 셀이 있는 시트에서 값이 어긋날 수 있음. 스크립트 내부 `getValues()`는 정확함
- `SpreadsheetApp.getUi().alert()`를 편집기에서 실행하면 열려 있는 시트 탭에 모달이 떠서 실행이 블로킹됨 → `toast()`로 교체
- Google Workspace 조직 정책에 따라 "확인하지 않은 앱" 경고가 뜰 수 있음 (자기 계정 스크립트이므로 고급 > 이동으로 진행)

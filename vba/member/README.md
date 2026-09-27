# 팀원용 업무일정 템플릿 – VBA 코드

계획안 v2(`docs/팀스케줄러_계획안_v2.md`)의 **팀원 파일**을 구현한 코드입니다.
빈 엑셀 파일에 코드를 넣고 `SetupTemplate` 매크로를 한 번 실행하면 시트, 표, 드롭다운, 버튼이 자동으로 만들어집니다.

## 폴더 구성

| 폴더 | 용도 |
|------|------|
| `src/` | 원본 코드 (UTF-8). GitHub에서 읽거나 수정할 때 사용합니다 |
| `import/` | **엑셀에 넣을 때 사용하는 파일**입니다. `.bas`·`.cls`는 한글 Windows VBA 편집기용(CP949, CRLF)이고, `.txt`는 붙여넣기용입니다 |

> `src`를 고친 뒤에는 `python tools/build_vba_import.py`를 실행해 `import` 폴더를 다시 만드세요.
> `src`의 `.bas` 파일을 그대로 가져오면 한글이 깨집니다.

| 파일 | 종류 | 내용 |
|------|------|------|
| `modConfig.bas` | 모듈 | 상수(시트·표·열 번호), 설정 읽기/쓰기, 날짜·색상 유틸 |
| `modTask.bas` | 모듈 | Task 저장·삭제, TaskID 발급, 상태/진척 자동 보정, 가중 진척률 |
| `modCalendar.bas` | 모듈 | 주간형 월간 달력: 막대 줄 배치, 보고 줄(◆), 주간 요약, 메모 보존, 월 이동 |
| `modDetail.bas` | 모듈 | 오른쪽 Task 상세 패널, 빠른 버튼(수정·완료 처리·+1일 연장), 입력폼 열기 |
| `modProject.bas` | 모듈 | PM 여부, 임시 프로젝트 요청, 단계(Phase) 조회·저장 |
| `modMasterLink.bas` | 모듈 | 팀장 `ProjectList.xlsx` 읽기(프로젝트·단계·코드 교체·코멘트·보고일정), 드롭다운 목록 |
| `modReport.bas` | 모듈 | 주간보고 초안 (텍스트 + 표) |
| `modMyProject.bas` | 모듈 | MyProject 주 단위 간트 (단계 → Task) |
| `modEvents.bas` | 모듈 | 파일 열기·저장·클릭·더블클릭·변경 처리 |
| `modSetup.bas` | 모듈 | **템플릿 자동 구성** (`SetupTemplate`) |
| `clsCtl.cls` | 클래스 | 입력폼 버튼·목록 이벤트 연결 |
| `frmTask.txt` | 폼 코드 | Task 등록·수정 폼 |
| `frmRequest.txt` | 폼 코드 | 새 프로젝트 요청 폼 |
| `frmPhase.txt` | 폼 코드 | 단계 편집 폼 (PM 전용) |
| `ThisWorkbook.txt` | 통합문서 코드 | 이벤트 연결 |

## 설치 순서 (템플릿 1회 제작, 약 10분)

1. 엑셀에서 **새 통합 문서**를 엽니다.
2. `Alt + F11`로 VBA 편집기를 엽니다.
3. **사용자 정의 폼 3개를 먼저 만듭니다.** 이렇게 해야 `Microsoft Forms 2.0` 참조가 자동으로 추가됩니다.
   - 메뉴 `삽입 → 사용자 정의 폼`을 누르고, 속성 창에서 `(이름)`을 **`frmTask`**로 바꿉니다.
   - 같은 방법으로 **`frmRequest`**와 **`frmPhase`**를 만듭니다.
   - 폼 위에 컨트롤을 그리지 않아도 됩니다. 코드가 자동으로 만듭니다.
4. 각 폼을 더블클릭해 코드 창을 열고, 이미 있는 내용을 지운 뒤 `import/frmTask.txt`, `frmRequest.txt`, `frmPhase.txt`의 내용을 **각각 전부 붙여넣습니다.**
5. 메뉴 `파일 → 파일 가져오기`로 `import/` 폴더의 **`.bas` 10개**와 **`clsCtl.cls`**를 가져옵니다.
6. 왼쪽 프로젝트 창에서 `현재_통합_문서`(ThisWorkbook)를 더블클릭하고 `import/ThisWorkbook.txt` 내용을 붙여넣습니다.
7. 메뉴 `디버그 → VBAProject 컴파일`을 실행해 오류가 없는지 확인합니다.
8. 엑셀로 돌아와 `Alt + F8` → **`SetupTemplate`** → 실행을 누릅니다.
9. **`업무일정_Template.xlsm`**(Excel 매크로 사용 통합 문서)으로 저장합니다.
10. 공유폴더에 아래처럼 배치합니다.
    ```
    \\회사서버\...\업무관리\
     ├─ Master\ProjectList.xlsx        ← 팀장 Master가 생성 (없어도 팀원 파일은 동작)
     ├─ Member\홍길동_업무일정.xlsm      ← 템플릿을 복사해 이름만 변경
     └─ Template\업무일정_Template.xlsm
    ```
11. 팀원은 파일을 처음 열 때 **이름**과 **영문 이니셜**(Task 번호 앞자리, 예: `KIM-0001`)을 한 번 입력합니다.

> 매크로가 막히면 IT 부서에 공유폴더를 **신뢰할 수 있는 위치**로 등록해 달라고 요청하세요. (계획안 10절 참고)

## 만들어지는 시트

| 시트 | 설명 |
|------|------|
| `Calendar` | 주간형 월간 달력. 버튼: `+ 새 Task` `◀` `▶` `이번 주` `주간보고` `새로고침`, 상단 **프로젝트 필터** |
| `MyProject` | 프로젝트 선택 → 단계/Task 주 단위 간트. `단계 편집 (PM)` 버튼 |
| `TaskDB` | Task 원본 표 `tblTask` (A~R열). 여러 건을 한꺼번에 고칠 때 직접 수정 가능 |
| `WeeklyReport` | `주간보고 만들기` → 금주 실적 / 차주 계획 / 이슈 (텍스트 + 표) |
| `Setting` (숨김) | 설정, 프로젝트·단계 캐시, 내 요청, 보고일정, 공휴일, 메모 |
| `CalMap` (완전 숨김) | 캘린더 칸 ↔ TaskID 매핑 |

## 사용법 요약

| 하고 싶은 일 | 방법 |
|--------------|------|
| 새 업무 등록 | 날짜 칸이나 빈 막대 칸을 **더블클릭** (또는 `+ 새 Task`) |
| 업무 수정·삭제 | 막대를 **더블클릭** |
| 업무 상세 보기 | 막대를 **한 번 클릭** → 오른쪽 상세 패널 (앞뒤 Task, 프로젝트 진척) |
| 빠른 완료·연장 | 상세 패널의 `완료 처리` / `+1일 연장` |
| 그날 업무 전체 보기 | 날짜 칸이나 빈 칸을 한 번 클릭 |
| 보고·마감 일정 | 날짜 바로 아래 **보고 줄**을 더블클릭 (◆ 표시) |
| 자유 메모 | 각 날짜 칸 아래 2줄에 직접 입력 (달을 넘겨도 유지됨) |
| 새 프로젝트 필요 | 입력폼 프로젝트 목록 맨 아래 `[+ 새 프로젝트 요청]` → 임시 코드로 바로 사용 |
| 단계 만들기 (PM) | 입력폼 `단계 편집` 또는 MyProject의 `단계 편집 (PM)` |
| 특정 프로젝트만 보기 | Calendar의 `프로젝트` 드롭다운 |
| 주간보고 | `주간보고` → 기준 날짜 입력 → WeeklyReport의 A열 복사 |

## 팀장 Master와 주고받는 데이터 (Master 개발 시 기준)

**Master가 팀원 파일에서 읽는 표** (읽기 전용으로 열기)

| 표 (시트) | 열 |
|-----------|----|
| `tblTask` (TaskDB) | TaskID, ProjectID, ProjectName, PhaseID, TaskName, StartDate, DueDate, Completion, Status, Priority, Note, Owner, PlanDays, DoneDate, CreatedDate, LastUpdated, LeaderComment, Milestone |
| `tblMyRequest` (Setting) | ReqCode, ProjectName, PlanStart, PlanEnd, Memo, ReqDate, Result, MappedTo, Reason |
| `tblMyPhase` (Setting) | PhaseID, ProjectID, Order, PhaseName, PlanStart, PlanEnd, Author, ModifiedAt — **PM인 프로젝트만 반영** |
| `tblMyPhaseMap` (Setting) | FromPhaseID, ToPhaseID, MergedAt — 단계 합치기 기록 → Master가 CodeMap(Type=Phase)으로 배포 |

**Master가 만드는 `Master\ProjectList.xlsx`** (각 시트 1행은 머리글)

| 시트 | 열 |
|------|----|
| `Projects` | ProjectID, ProjectName, PM(1인 프로젝트는 담당자 자동), Color(RGB 숫자), Status |
| `Phases` | PhaseID, ProjectID, Order, PhaseName, PlanStart, PlanEnd, Author, ModifiedAt |
| `CodeMap` | FromCode, ToCode, Type(`Project`/`Phase`), Result(`승인`/`통합`/`반려`), Reason |
| `Comments` | TaskID, Comment, Date |
| `Events` | Date, Time, Title, Target(`전체` 또는 이름) — 반복 일정은 날짜별로 펼쳐서 기록 |
| `Holidays` (선택) | Date, Name — 있으면 팀원 공휴일 표를 교체 |

## 알려진 제약 / 확인 필요

- 이 코드는 Excel이 없는 환경에서 작성해서 **실제 엑셀로 실행해 보지는 않았습니다.** 구조 검사(블록 짝 맞춤, 호출 대상 존재 여부)만 통과했습니다. 처음 설치할 때 `디버그 → 컴파일`과 파일럿 테스트가 필요합니다.
- 주 시작 요일은 **월요일로 고정**입니다.
- 공휴일은 2026년 기본 공휴일만 넣었습니다. **대체공휴일·선거일은 Setting 시트 `tblHoliday`에 직접 추가**하거나, Master의 `Holidays` 시트로 배포하세요.
- 메모 칸에서 여러 셀을 한꺼번에 붙여넣으면 60칸까지만 저장합니다.
- PM의 **다른 참여자 Task 보기(PMView)**와 **새 단계 제안**은 계획안의 2단계 기능이라 아직 넣지 않았습니다.

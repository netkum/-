팀 스케줄러 (엑셀 VBA) - 설치 폴더
==============================================

이 폴더를 통째로 원하는 위치에 두세요. 예: F:\claude\Scheduler
모든 경로는 이 폴더 기준 "상대 경로"라서 위치가 바뀌어도 그대로 동작합니다.

  Scheduler\
   ├─ build.bat              ← 더블클릭하면 엑셀 파일 2개를 자동 생성
   ├─ build.ps1              ← build.bat이 실행하는 스크립트
   ├─ vba\member\            ← 팀원용 템플릿 코드 (가져오기용, CP949)
   ├─ vba\master\            ← 팀장 Master 코드
   ├─ Template\              → 업무일정_Template.xlsm 이 생성됨
   ├─ Master\                → 연구팀_Master.xlsm 이 생성됨 (+ 나중에 ProjectList.xlsx)
   └─ Member\                → 팀원 파일(이름_업무일정.xlsm)을 두는 곳

[방법 1] 자동 생성 (권장)
----------------------------------------------
1. 엑셀 > 파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정
   > "VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음" 체크 (생성 후 다시 꺼도 됩니다)
2. build.bat 더블클릭 → 엑셀이 잠깐 열렸다 닫히며 파일 2개 생성
3. Template\업무일정_Template.xlsm 을 Member 폴더에 복사하고
   "홍길동_업무일정.xlsm" 처럼 팀원 이름으로 바꿉니다. (팀원 수만큼)
4. Master\연구팀_Master.xlsm 을 열어 ProjectMaster 시트에 프로젝트 등록 → [ProjectList 배포]
5. 팀원 파일을 열면 처음에 이름·이니셜을 묻습니다 → Task 등록 후 저장
6. Master에서 [전체 동기화] → Dashboard 확인

* build.bat / build.ps1 이 보이지 않거나(보안 프로그램이 삭제) 실행이 막히면 [방법 1-B]를 사용하세요.

[방법 1-B] 엑셀 안에서 자동 생성 (bat·PowerShell 없이)
----------------------------------------------
1. 위 1번의 "VBA 프로젝트 개체 모델에 안전하게 액세스" 체크
2. 엑셀에서 새 통합 문서 → Alt+F11 (VBA 편집기)
3. 파일 > 파일 가져오기 → Scheduler 폴더의 Builder.bas 선택
4. 엑셀로 돌아와 Alt+F8 → BuildAll → 실행 → 폴더 경로 확인(기본 F:\claude\Scheduler) → 확인
5. "완료!" 창이 뜨면 Template·Master 폴더에 파일 2개가 생성됨 (이 새 통합 문서는 저장하지 않고 닫아도 됨)

[방법 2] 수동 설치
----------------------------------------------
팀원 템플릿:
 1) 새 통합 문서 → Alt+F11
 2) 삽입 > 사용자 정의 폼 3개 추가, (이름)을 frmTask / frmRequest / frmPhase 로 변경
 3) 각 폼 코드 창에 vba\member\frmTask.txt / frmRequest.txt / frmPhase.txt 내용 붙여넣기
 4) 파일 > 파일 가져오기: vba\member 의 .bas 10개 + clsCtl.cls
 5) ThisWorkbook(현재_통합_문서)에 vba\member\ThisWorkbook.txt 붙여넣기
 6) 디버그 > 컴파일 → Alt+F8 > SetupTemplate 실행
 7) Template\업무일정_Template.xlsm (매크로 사용 통합 문서)으로 저장
팀장 Master:
 1) 새 통합 문서 → Alt+F11 → 파일 가져오기: vba\master 의 .bas 9개
 2) ThisWorkbook에 vba\master\ThisWorkbook.txt 붙여넣기
 3) 디버그 > 컴파일 → Alt+F8 > SetupMaster 실행
 4) Master\연구팀_Master.xlsm 으로 저장

참고
----------------------------------------------
- 팀원 파일은 기본으로 "..\Master\ProjectList.xlsx"를 읽고,
  Master는 기본으로 "..\Member\*_업무일정.xlsm"을 수집합니다. (폴더 구조를 유지하세요)
- 공유폴더로 옮길 때도 Scheduler 폴더 구조 그대로 복사하면 됩니다.
- 매크로가 막히면 이 폴더를 "신뢰할 수 있는 위치"로 등록하세요.
- 코드는 Excel 없이 작성되어 실제 실행 검증 전입니다. 오류 메시지가 나오면 알려주세요.

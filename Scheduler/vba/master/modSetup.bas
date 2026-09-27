Attribute VB_Name = "modSetup"
'==============================================================
' 팀장 Master 최초 구성: SetupMaster 한 번 실행
'==============================================================
Option Explicit

Private mSilent As Boolean

' 자동 생성 스크립트(build.ps1)용: 확인창 없이 실행
Public Sub SetupMasterSilent()
    mSilent = True
    SetupMaster
    mSilent = False
End Sub

Public Sub SetupMaster()
    If Not mSilent Then
        If MsgBox("팀장 Master 구조(시트·표·버튼)를 만듭니다." & vbLf & _
              "같은 이름의 시트가 있으면 내용이 초기화됩니다. 계속할까요?", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub
    End If
    On Error GoTo EH
    SpeedOn
    EnsureSheet SH_DASH
    EnsureSheet SH_PV
    EnsureSheet SH_TS
    EnsureSheet SH_ALL
    EnsureSheet SH_PM
    EnsureSheet SH_SET
    OrderSheets

    BuildSetting
    BuildProjectMaster
    RefreshLists
    AddNamedValidations
    BuildTaskAll
    BuildDashboard
    BuildProjectView
    BuildTeamSchedule
    RemoveOtherSheets
    SpeedOff

    RefreshViews
    GetWS(SH_DASH).Activate
    If Not mSilent Then
        Msg "Master 구성이 끝났습니다." & vbLf & vbLf & _
            "1) '연구팀_Master.xlsm'(매크로 사용 통합 문서)으로 공유폴더의 Master 폴더에 저장하세요." & vbLf & _
            "2) Setting 시트에서 MemberFolder(팀원 파일 폴더)를 확인하세요. 비우면 '상위 폴더\Member'." & vbLf & _
            "3) ProjectMaster 시트에 프로젝트를 등록하고 [ProjectList 배포]를 누르세요." & vbLf & _
            "4) 팀 공통 보고일정·공휴일은 Setting 시트에서 관리합니다."
    End If
    Exit Sub
EH:
    SpeedReset
    If Not mSilent Then Msg "구성 중 오류: " & Err.Description, vbCritical
End Sub

Private Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Set ws = GetWS(nm)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    ws.Visible = xlSheetVisible
    Set EnsureSheet = ws
End Function

Private Sub OrderSheets()
    Dim names As Variant, i As Long
    names = Array(SH_DASH, SH_PV, SH_TS, SH_ALL, SH_PM, SH_SET)
    GetWS(names(0)).Move Before:=ThisWorkbook.Worksheets(1)
    For i = 1 To UBound(names)
        GetWS(names(i)).Move After:=GetWS(names(i - 1))
    Next i
End Sub

Private Sub RemoveOtherSheets()
    Dim sh As Worksheet, keep As String
    keep = "|" & SH_DASH & "|" & SH_PV & "|" & SH_TS & "|" & SH_ALL & "|" & SH_PM & "|" & SH_SET & "|"
    Application.DisplayAlerts = False
    For Each sh In ThisWorkbook.Worksheets
        If InStr(keep, "|" & sh.Name & "|") = 0 Then
            If Application.WorksheetFunction.CountA(sh.UsedRange) = 0 Then sh.Delete
        End If
    Next sh
    Application.DisplayAlerts = True
End Sub

Private Sub ResetSheet(ByVal ws As Worksheet)
    Dim lo As ListObject, shp As Shape
    For Each lo In ws.ListObjects: lo.Delete: Next lo
    For Each shp In ws.Shapes: shp.Delete: Next shp
    ws.Cells.Clear
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 9
End Sub

Private Function MakeTable(ByVal ws As Worksheet, ByVal topLeft As String, ByVal tblName As String, _
                           ByVal headers As Variant, Optional ByVal title As String = "") As ListObject
    Dim rng As Range, lo As ListObject, n As Long
    n = UBound(headers) - LBound(headers) + 1
    Set rng = ws.Range(topLeft).Resize(1, n)
    rng.Value = headers
    If title <> "" And rng.Row > 1 Then
        rng.Cells(1, 1).Offset(-1, 0).Value = title
        rng.Cells(1, 1).Offset(-1, 0).Font.Bold = True
    End If
    Set lo = ws.ListObjects.Add(xlSrcRange, rng.Resize(2), , xlYes)
    lo.Name = tblName
    lo.TableStyle = "TableStyleLight9"
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    Set MakeTable = lo
End Function

Private Sub AddListValidation(ByVal rng As Range, ByVal formula As String, Optional ByVal strict As Boolean = True)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=IIf(strict, xlValidAlertStop, xlValidAlertInformation), Formula1:=formula
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Function AddButton(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal macro As String, _
                           ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal col As Long) As Double
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, 21)
    With shp
        .Name = nm
        .OnAction = macro
        .Fill.ForeColor.RGB = col
        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
        With .TextFrame2
            .TextRange.Text = caption
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 1: .MarginRight = 1: .MarginTop = 0: .MarginBottom = 0
        End With
    End With
    AddButton = x + w + 4
End Function

' 공통 상단 버튼 (전체 동기화 / 요청 처리 / 배포)
Private Sub TopButtons(ByVal ws As Worksheet, ByVal x As Double)
    x = AddButton(ws, "btnSync_" & ws.Index, "전체 동기화", "SyncAll", x, 4, 80, RGB(47, 117, 181))
    x = AddButton(ws, "btnReq_" & ws.Index, "요청 처리 적용", "ProcessRequests", x, 4, 90, RGB(84, 130, 53))
    x = AddButton(ws, "btnPub_" & ws.Index, "ProjectList 배포", "PublishButton", x, 4, 95, RGB(127, 127, 127))
End Sub

Private Sub Title(ByVal ws As Worksheet, ByVal addr As String, ByVal text As String)
    ws.Rows(1).RowHeight = 30
    ws.Range(addr).Value = text
    ws.Range(addr).Font.Size = 14
    ws.Range(addr).Font.Bold = True
End Sub

Private Sub Freeze(ByVal ws As Worksheet, ByVal addr As String)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range(addr).Select
    ActiveWindow.FreezePanes = True
End Sub

'--- Setting ------------------------------------------------
Private Sub BuildSetting()
    Dim ws As Worksheet, lo As ListObject, h As Variant, i As Long
    Set ws = GetWS(SH_SET)
    ResetSheet ws
    Title ws, "A1", "설정"
    Set lo = MakeTable(ws, "A3", TB_SET, Array("Key", "Value"), "기본 설정")
    AppendRow lo, Array("MemberFolder", "")
    AppendRow lo, Array("ExportFolder", "")
    AppendRow lo, Array("StaleDays", 7)
    AppendRow lo, Array("ProjectPrefix", "P")
    AppendRow lo, Array("LastSync", "")
    AppendRow lo, Array("LastExport", "")
    AppendRow lo, Array("CommentDirty", "")
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 44
    ws.Range("A12").Value = "MemberFolder: 팀원 파일 폴더 (비우면 이 파일 폴더의 상위\Member)"
    ws.Range("A13").Value = "ExportFolder: ProjectList.xlsx 저장 폴더 (비우면 이 파일이 있는 폴더)"
    ws.Range("A14").Value = "StaleDays: 이 일수 이상 수정 없는 진행 Task → 미갱신"
    ws.Range("A12:A14").Font.Color = RGB(128, 128, 128)

    ws.Range("M:M").NumberFormat = "@"
    MakeTable ws, "D3", TB_MEMBER, Array("Name", "File", "LastSync", "Result", "Rows", "LastUpdated", "Version"), "팀원 파일 (동기화 시 자동)"
    Set lo = MakeTable(ws, "L3", TB_EVENT, Array("Title", "Time", "Target", "Repeat", "Date", "Weekday", "DayOfMonth", "From", "To"), _
                       "팀 공통 보고일정 (Repeat: 한번/매주/매월/매월마지막, Target 비우면 전체)")
    AppendRow lo, Array("주간회의", "10:00", "전체", "매주", Empty, "월", Empty, Empty, Empty)
    AppendRow lo, Array("주간보고 제출", "", "전체", "매주", Empty, "금", Empty, Empty, Empty)
    Set lo = MakeTable(ws, "V3", TB_HOL, Array("HolDate", "HolName"), "공휴일 (대체공휴일 등 확인 후 추가)")
    h = Array("2026-01-01", "신정", "2026-02-16", "설날", "2026-02-17", "설날", "2026-02-18", "설날", _
              "2026-03-01", "삼일절", "2026-05-05", "어린이날", "2026-05-24", "부처님오신날", _
              "2026-06-06", "현충일", "2026-08-15", "광복절", "2026-09-24", "추석", "2026-09-25", "추석", _
              "2026-09-26", "추석", "2026-10-03", "개천절", "2026-10-09", "한글날", "2026-12-25", "성탄절")
    For i = 0 To UBound(h) Step 2
        AppendRow lo, Array(CDate(h(i)), h(i + 1))
    Next i
    MakeTable ws, "Y3", TB_LOG, Array("SyncTime", "File", "Owner", "Result", "Rows", "Message"), "동기화 기록 (최근 500건)"

    ws.Range("P:P,S:T,V:V").NumberFormat = "yyyy-mm-dd"
    ws.Range("F:F,I:I").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Range("Y:Y").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Range("D:D").ColumnWidth = 10
    ws.Range("E:E").ColumnWidth = 24
    ws.Range("G:G").ColumnWidth = 22
    ws.Range("L:L").ColumnWidth = 16
    ws.Range("Y:Y").ColumnWidth = 16
    ws.Range("AD:AD").ColumnWidth = 30
    ws.Range("AH2").Value = "드롭다운 목록 (자동)"
    ws.Range("AH2").Font.Bold = True
End Sub

'--- ProjectMaster ------------------------------------------
Private Sub BuildProjectMaster()
    Dim ws As Worksheet, x As Double
    Set ws = GetWS(SH_PM)
    ResetSheet ws
    Title ws, "A1", "프로젝트 관리"
    MakeTable ws, "A4", TB_PROJ, Array("ProjectID", "ProjectName", "PM(지정)", "PM(적용)", "참여자(자동)", "StartDate", "EndDate", "Color", "Status"), _
              "프로젝트 - PM(지정)은 참여자 2명 이상일 때 팀장이 선택 / Color 칸에 채우기 색을 칠하면 그 색 사용"
    MakeTable ws, "K4", TB_PHASE, Array("PhaseID", "ProjectID", "Order", "PhaseName", "PlanStart", "PlanEnd", "Author", "ModifiedAt"), _
              "단계 - PM이 팀원 파일에서 작성 (팀장이 직접 고치면 팀장 수정 우선)"
    MakeTable ws, "T4", TB_REQ, Array("ReqCode", "요청자", "프로젝트명", "예상 시작", "예상 종료", "메모", "Task 수", "요청일", _
                                      "처리▼", "정식코드/통합대상", "반려 사유", "처리일시"), _
              "승인대기 - 처리▼ 선택(승인/기존 통합/반려) → [요청 처리 적용]  (승인 시 정식코드 비우면 자동 발급)"
    MakeTable ws, "AG4", TB_MAP, Array("FromCode", "ToCode", "Type", "Result", "Reason", "Date"), "코드 변경 기록 (자동)"
    MakeTable ws, "AN4", TB_CMT, Array("TaskID", "Comment", "Date", "Owner", "TaskName"), "팀장 코멘트 (TaskAll·Dashboard에서 입력)"

    AddListValidation ws.Range("I5:I1000"), "진행,보류,완료,종료"

    ws.Range("A:A,K:K,L:L,T:T,AG:AH,AN:AN").NumberFormat = "@"
    ws.Range("F:G,O:P,W:X").NumberFormat = "yyyy-mm-dd"
    ws.Range("R:R,AA:AA,AE:AE,AL:AL,AP:AP").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Range("D5:E1000").Interior.Color = RGB(242, 242, 242)
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C:D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F:G").ColumnWidth = 11
    ws.Columns("H:I").ColumnWidth = 8
    ws.Columns("J").ColumnWidth = 2
    ws.Columns("K").ColumnWidth = 10
    ws.Columns("N").ColumnWidth = 18
    ws.Columns("V").ColumnWidth = 20
    ws.Columns("AB:AC").ColumnWidth = 14
    ws.Columns("AO").ColumnWidth = 30
    x = ws.Range("C1").Left
    TopButtons ws, x
    ws.Range("A2").Value = "※ 프로젝트 추가: 표 아래 빈 행에 ProjectID(예: P001)·이름 입력 → [ProjectList 배포]. 팀원 드롭다운에 반영됩니다."
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
End Sub

' 이름 정의(목록)가 만들어진 뒤에 거는 드롭다운
Private Sub AddNamedValidations()
    Dim ws As Worksheet
    Set ws = GetWS(SH_SET)
    AddListValidation ws.Range("O4:O500"), "=lstRepeat"
    AddListValidation ws.Range("Q4:Q500"), "=lstWeekday"
    Set ws = GetWS(SH_PM)
    AddListValidation ws.Range("C5:C1000"), "=lstMember", False
    AddListValidation ws.Range("AB5:AB1000"), "=lstAction"
    AddListValidation ws.Range("L5:L2000"), "=lstProjectID", False
End Sub

'--- TaskAll ------------------------------------------------
Private Sub BuildTaskAll()
    Dim ws As Worksheet, lo As ListObject
    Set ws = GetWS(SH_ALL)
    ResetSheet ws
    ws.Range("A1").Resize(1, A_COLS).Value = Array("TaskID", "ProjectID", "ProjectName", "PhaseID", "TaskName", "StartDate", "DueDate", _
        "Completion", "Status", "Priority", "Note", "Owner", "PlanDays", "DoneDate", "CreatedDate", "LastUpdated", "팀장코멘트(입력)", _
        "Milestone", "SourceOwner", "SyncedAt", "지연", "계획진척", "차이", "D-day", "미수정일수")
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, A_COLS), , xlYes)
    lo.Name = TB_ALL
    lo.TableStyle = "TableStyleLight9"
    lo.DataBodyRange.Delete
    ws.Range("F:G,N:N").NumberFormat = "yyyy-mm-dd"
    ws.Range("O:P,T:T").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Range("A:A,B:B,D:D").NumberFormat = "@"
    ws.Columns("A:D").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 26
    ws.Columns("K").ColumnWidth = 26
    ws.Columns("Q").ColumnWidth = 30
    ws.Range("Q1").Interior.Color = RGB(255, 192, 0)
    ws.Range("Q2:Q5000").Interior.Color = RGB(255, 242, 204)
    With ws.Range("A2:Y5000").FormatConditions.Add(Type:=xlExpression, Formula1:="=$U2=""Y""")
        .Font.Color = RGB(192, 0, 0)
    End With
    Freeze ws, "F2"
End Sub

'--- Dashboard ----------------------------------------------
Private Sub BuildDashboard()
    Dim ws As Worksheet, widths As Variant, i As Long
    Set ws = GetWS(SH_DASH)
    ResetSheet ws
    Title ws, "B1", "팀 업무 Dashboard"
    widths = Array(2, 26, 12, 16, 16, 12, 12, 9, 9, 8, 8, 6, 26, 4, 10)
    For i = 0 To UBound(widths)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i
    ws.Columns(DB_KEYCOL).Hidden = True
    ws.Rows(5).RowHeight = 30
    TopButtons ws, ws.Range("D1").Left
End Sub

'--- ProjectView --------------------------------------------
Private Sub BuildProjectView()
    Dim ws As Worksheet
    Set ws = GetWS(SH_PV)
    ResetSheet ws
    Title ws, "B1", "프로젝트 보기"
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("A").Hidden = True
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D:E").ColumnWidth = 7
    ws.Range(ws.Columns(6), ws.Columns(6 + 30)).ColumnWidth = 3.3
    ws.Range("B2").Value = "프로젝트 ▶"
    ws.Range("B2").HorizontalAlignment = xlRight
    ws.Range("C2:P2").Merge
    AddListValidation ws.Range(PV_FILTER), "=lstProjectAll"
    ws.Range("C2:P2").Interior.Color = RGB(255, 242, 204)
    TopButtons ws, ws.Range("F1").Left
End Sub

'--- TeamSchedule -------------------------------------------
Private Sub BuildTeamSchedule()
    Dim ws As Worksheet, x As Double
    Set ws = GetWS(SH_TS)
    ResetSheet ws
    Title ws, "B1", "팀 일정 (월간)"
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 12
    ws.Range(ws.Columns(3), ws.Columns(33)).ColumnWidth = 5.5
    ws.Range("B2").Value = "월 ▶"
    ws.Range("B2").HorizontalAlignment = xlRight
    ws.Range("C2:E2").Merge
    ws.Range(TS_MONTH).Value = DateSerial(Year(Date), Month(Date), 1)
    ws.Range(TS_MONTH).NumberFormat = "yyyy-mm"
    ws.Range("C2:E2").Interior.Color = RGB(255, 242, 204)
    x = ws.Range("G1").Left
    x = AddButton(ws, "btnTSPrev", "◀ 이전달", "TSPrevMonth", x, 4, 60, RGB(89, 89, 89))
    x = AddButton(ws, "btnTSNext", "다음달 ▶", "TSNextMonth", x, 4, 60, RGB(89, 89, 89))
    TopButtons ws, x + 10
    Freeze ws, "C7"
End Sub

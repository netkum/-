Attribute VB_Name = "modSetup"
'==============================================================
' 템플릿 최초 구성: 시트·표·서식·드롭다운·버튼을 자동으로 만든다
'   새 통합문서에 모듈을 넣은 뒤 SetupTemplate 을 한 번 실행
'==============================================================
Option Explicit

Public Sub SetupTemplate()
    If MsgBox("업무일정 템플릿 구조(시트·표·버튼)를 만듭니다." & vbLf & _
              "같은 이름의 시트가 있으면 내용이 초기화됩니다. 계속할까요?", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub
    On Error GoTo EH
    SpeedOn
    EnsureSheet SH_CAL
    EnsureSheet SH_MP
    EnsureSheet SH_DB
    EnsureSheet SH_RPT
    EnsureSheet SH_SET
    EnsureSheet SH_MAP
    OrderSheets

    BuildSetting
    RefreshLists
    BuildTaskDB
    BuildCalendar
    BuildMyProject
    BuildReport
    BuildMap
    RemoveOtherSheets

    CfgSet "Version", APP_VERSION
    CfgSet "CurYear", Year(Date)
    CfgSet "CurMonth", Month(Date)
    SpeedOff

    RenderCalendar
    GetWS(SH_CAL).Activate
    Msg "템플릿 구성이 끝났습니다." & vbLf & vbLf & _
        "1) 파일을 '업무일정_Template.xlsm'(매크로 사용 통합 문서)으로 저장하세요." & vbLf & _
        "2) 팀원은 이 파일을 복사해 '이름_업무일정.xlsm'으로 이름만 바꿔 사용합니다." & vbLf & _
        "3) 공휴일은 Setting 시트의 tblHoliday에서 확인·추가하세요."
    Exit Sub
EH:
    SpeedReset
    Msg "템플릿 구성 중 오류: " & Err.Description, vbCritical
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
    names = Array(SH_CAL, SH_MP, SH_DB, SH_RPT, SH_SET, SH_MAP)
    GetWS(names(0)).Move Before:=ThisWorkbook.Worksheets(1)
    For i = 1 To UBound(names)
        GetWS(names(i)).Move After:=GetWS(names(i - 1))
    Next i
End Sub

Private Sub RemoveOtherSheets()
    Dim sh As Worksheet, keep As String
    keep = "|" & SH_CAL & "|" & SH_MP & "|" & SH_DB & "|" & SH_RPT & "|" & SH_SET & "|" & SH_MAP & "|"
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

' 머리글만 있는 빈 표 만들기
Private Function MakeTable(ByVal ws As Worksheet, ByVal topLeft As String, ByVal tblName As String, _
                           ByVal headers As Variant, Optional ByVal title As String = "") As ListObject
    Dim rng As Range, lo As ListObject, n As Long
    n = UBound(headers) - LBound(headers) + 1
    Set rng = ws.Range(topLeft).Resize(1, n)
    rng.Value = headers
    If title <> "" Then
        rng.Cells(1, 1).Offset(-1, 0).Value = title
        rng.Cells(1, 1).Offset(-1, 0).Font.Bold = True
    End If
    Set lo = ws.ListObjects.Add(xlSrcRange, rng.Resize(2), , xlYes)
    lo.Name = tblName
    lo.TableStyle = "TableStyleLight9"
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    Set MakeTable = lo
End Function

'--- Setting ------------------------------------------------
Private Sub BuildSetting()
    Dim ws As Worksheet, lo As ListObject, def As String
    Set ws = GetWS(SH_SET)
    ResetSheet ws
    ws.Range("A1").Value = "설정 (VBA가 사용하는 데이터입니다. 표 머리글·위치를 바꾸지 마세요)"
    ws.Range("A1").Font.Bold = True

    Set lo = MakeTable(ws, "A3", TB_SET, Array("Key", "Value"), "기본 설정")
    def = ThisWorkbook.Name
    If InStr(def, "_") > 0 Then def = Left$(def, InStr(def, "_") - 1) Else def = ""
    AppendRow lo, Array("OwnerName", def)
    AppendRow lo, Array("Initial", "")
    AppendRow lo, Array("MasterFolder", "")
    AppendRow lo, Array("CurYear", Year(Date))
    AppendRow lo, Array("CurMonth", Month(Date))
    AppendRow lo, Array("LastSeq", 0)
    AppendRow lo, Array("Version", APP_VERSION)
    AppendRow lo, Array("LastMasterSync", "")
    AppendRow lo, Array("DetailTaskID", "")
    AppendRow lo, Array("DetailRow", "")
    AppendRow lo, Array("CalDirty", "")
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 40
    ws.Range("C4").Value = "← MasterFolder 비우면 자동: 내 파일 폴더의 상위\Master"
    ws.Range("C4").Font.Color = RGB(128, 128, 128)

    MakeTable ws, "D3", TB_PROJ, Array("ProjectID", "ProjectName", "PM", "Color", "Status"), "프로젝트 (팀장 파일에서 자동)"
    MakeTable ws, "J3", TB_PHASE, Array("PhaseID", "ProjectID", "Order", "PhaseName", "PlanStart", "PlanEnd", "Author", "ModifiedAt"), "단계 (팀장 파일에서 자동)"
    MakeTable ws, "S3", TB_MYPHASE, Array("PhaseID", "ProjectID", "Order", "PhaseName", "PlanStart", "PlanEnd", "Author", "ModifiedAt"), "내가 PM인 프로젝트 단계"
    MakeTable ws, "AB3", TB_PHMAP, Array("FromPhaseID", "ToPhaseID", "MergedAt"), "단계 합치기 기록"
    MakeTable ws, "AF3", TB_REQ, Array("ReqCode", "ProjectName", "PlanStart", "PlanEnd", "Memo", "ReqDate", "Result", "MappedTo", "Reason"), "내 프로젝트 요청"
    MakeTable ws, "AP3", TB_EVENT, Array("EventDate", "EventTime", "Title", "Target"), "팀 공통 보고일정 (팀장 파일에서 자동)"
    Set lo = MakeTable(ws, "AU3", TB_HOL, Array("HolDate", "HolName"), "공휴일 (확인 후 추가·수정)")
    AddHolidays lo
    MakeTable ws, "AX3", TB_MEMO, Array("MemoDate", "Line", "Text"), "캘린더 메모"

    ws.Range("D:D,J:K,S:T,AB:AC,AF:AF").NumberFormat = "@"
    ws.Range("N:O,W:X,AH:AI,AK:AK,AP:AP,AU:AU,AX:AX").NumberFormat = "yyyy-mm-dd"
    ws.Range("Q:Q,Z:Z,AD:AD").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Range("BB2").Value = "드롭다운 목록 (자동)"
    ws.Range("BB2").Font.Bold = True
End Sub

' 2026년 공휴일 (대체공휴일·임시공휴일은 매년 확인 후 추가)
Private Sub AddHolidays(ByVal lo As ListObject)
    Dim h As Variant, i As Long
    h = Array("2026-01-01", "신정", "2026-02-16", "설날", "2026-02-17", "설날", "2026-02-18", "설날", _
              "2026-03-01", "삼일절", "2026-05-05", "어린이날", "2026-05-24", "부처님오신날", _
              "2026-06-06", "현충일", "2026-08-15", "광복절", "2026-09-24", "추석", "2026-09-25", "추석", _
              "2026-09-26", "추석", "2026-10-03", "개천절", "2026-10-09", "한글날", "2026-12-25", "성탄절")
    For i = 0 To UBound(h) Step 2
        AppendRow lo, Array(CDate(h(i)), h(i + 1))
    Next i
End Sub

'--- TaskDB -------------------------------------------------
Private Sub BuildTaskDB()
    Dim ws As Worksheet, lo As ListObject, i As Long, notes As Variant, heads As Variant
    Set ws = GetWS(SH_DB)
    ResetSheet ws
    heads = Array("TaskID", "ProjectID", "ProjectName", "PhaseID", "TaskName", "StartDate", "DueDate", _
                  "Completion", "Status", "Priority", "Note", "Owner", "PlanDays", "DoneDate", _
                  "CreatedDate", "LastUpdated", "LeaderComment", "Milestone")
    notes = Array("자동: 업무번호", "입력: 프로젝트 코드 (드롭다운)", "자동: 프로젝트명", "선택: 단계 ID (입력폼에서 선택 권장)", _
                  "입력: 업무명", "입력: 시작일", "입력: 종료(예정)일", "입력: 진척률 0~100", "입력: 예정/진행/완료/보류/취소", _
                  "선택: 상/중/하", "선택: 지연 사유·이슈 (팀장 지연목록에 표시)", "자동: 담당자", "자동: 공휴일 제외 작업일수", _
                  "자동: 완료일", "자동: 등록일시", "자동: 최종수정일시", "자동: 팀장 코멘트 (읽기 전용)", "선택: Y = 보고·마감 일정(◆)")
    ws.Range("A1").Resize(1, T_COLS).Value = heads
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, T_COLS), , xlYes)
    lo.Name = TB_TASK
    lo.TableStyle = "TableStyleLight9"
    lo.DataBodyRange.Delete
    For i = 0 To T_COLS - 1
        ws.Cells(1, i + 1).AddComment CStr(notes(i))
    Next i

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 26
    ws.Columns("F:G").ColumnWidth = 11
    ws.Columns("H:J").ColumnWidth = 8
    ws.Columns("K").ColumnWidth = 30
    ws.Columns("L:M").ColumnWidth = 8
    ws.Columns("N:P").ColumnWidth = 15
    ws.Columns("Q").ColumnWidth = 30
    ws.Columns("R").ColumnWidth = 8
    ws.Range("F:G,N:N").NumberFormat = "yyyy-mm-dd"
    ws.Range("O:P").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Range("A:A,B:B,D:D").NumberFormat = "@"

    AddListValidation ws.Range("B2:B5000"), "=lstProject"
    AddListValidation ws.Range("H2:H5000"), "=lstCompletion"
    AddListValidation ws.Range("I2:I5000"), "=lstStatus"
    AddListValidation ws.Range("J2:J5000"), "=lstPriority"
    AddListValidation ws.Range("R2:R5000"), "=lstYN"
    With ws.Range("F2:G5000").Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=CStr(CLng(DateSerial(2000, 1, 1)))
        .ErrorMessage = "날짜를 입력하세요 (예: 2026-10-07)"
    End With

    ' 필수값 누락 행 연한 빨강
    With ws.Range("A2:R5000").FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND(COUNTA($B2,$E2,$F2,$G2)>0,COUNTA($B2,$E2,$F2,$G2)<4)")
        .Interior.Color = RGB(255, 220, 220)
    End With
    ' 자동 열 회색 머리글
    ws.Range("A1,C1,L1:Q1").Interior.Color = RGB(166, 166, 166)

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub AddListValidation(ByVal rng As Range, ByVal formula As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formula
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

'--- Calendar -----------------------------------------------
Private Sub BuildCalendar()
    Dim ws As Worksheet, x As Double
    Set ws = GetWS(SH_CAL)
    ResetSheet ws

    ws.Columns(COL_WEEK).ColumnWidth = 7
    ws.Range(ws.Columns(COL_DAY1), ws.Columns(COL_DAY1 + 4)).ColumnWidth = 20
    ws.Columns(COL_WKND).ColumnWidth = 8
    ws.Columns(COL_SUM).ColumnWidth = 30
    ws.Columns(COL_SUM + 1).ColumnWidth = 2
    ws.Columns(COL_DETAIL).ColumnWidth = 44

    ws.Rows(1).RowHeight = 28
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True

    ws.Range("B2").Value = "프로젝트"
    ws.Range("B2").HorizontalAlignment = xlRight
    ws.Range(CAL_FILTER).Value = FILTER_ALL
    AddListValidation ws.Range(CAL_FILTER), "=lstFilter"
    ws.Range(CAL_FILTER).Interior.Color = RGB(255, 242, 204)
    ws.Range("D2").Value = "'" & SYM_MILE & "보고·마감   " & SYM_DELAY & "지연   " & SYM_DONE & "완료   " & _
                           SYM_COMMENT & "팀장 코멘트   *승인대기   " & SYM_PREV & SYM_NEXT & "앞/뒤 주로 이어짐"
    ws.Range("D2").Font.Color = RGB(89, 89, 89)
    ws.Range("D2").Font.Size = 8

    ws.Range("A3:J3").Value = Array("주차", "월", "화", "수", "목", "금", "토·일", "주간 요약", "", "Task 상세 (막대를 클릭)")
    With ws.Range("A3:H3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(68, 84, 106)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("G3").Font.Color = RGB(255, 199, 206)
    ws.Range("J3").Font.Bold = True
    ws.Range("J3").Font.Color = RGB(89, 89, 89)

    ' 상단 버튼
    x = ws.Range("D1").Left
    x = AddButton(ws, "btnNew", "+ 새 Task", "NewTaskButton", x, 3, 70, RGB(47, 117, 181))
    x = AddButton(ws, "btnPrev", SYM_PREV, "PrevMonth", x, 3, 28, RGB(89, 89, 89))
    x = AddButton(ws, "btnNext", SYM_NEXT, "NextMonth", x, 3, 28, RGB(89, 89, 89))
    x = AddButton(ws, "btnToday", "이번 주", "GoThisWeek", x, 3, 55, RGB(89, 89, 89))
    x = AddButton(ws, "btnReport", "주간보고", "MakeWeeklyReport", x, 3, 60, RGB(84, 130, 53))
    x = AddButton(ws, "btnRefresh", "새로고침", "RefreshAll", x, 3, 60, RGB(127, 127, 127))
    ' 상세 패널 버튼 (선택 시에만 보임)
    x = ws.Cells(1, COL_DETAIL).Left
    x = AddButton(ws, "btnDetailEdit", "수정", "DetailEdit", x, 3, 50, RGB(47, 117, 181))
    x = AddButton(ws, "btnDetailDone", "완료 처리", "DetailComplete", x, 3, 60, RGB(84, 130, 53))
    x = AddButton(ws, "btnDetailExtend", "+1일 연장", "DetailExtend", x, 3, 60, RGB(191, 144, 0))
    ws.Shapes("btnDetailEdit").Visible = msoFalse
    ws.Shapes("btnDetailDone").Visible = msoFalse
    ws.Shapes("btnDetailExtend").Visible = msoFalse

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A" & CAL_TOP).Select
    ActiveWindow.FreezePanes = True

    With ws.PageSetup
        .PrintArea = "$A$1:$H$" & LastCalRow()
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
End Sub

' 둥근 사각형 버튼 → 다음 버튼의 x 위치 반환
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

'--- MyProject ----------------------------------------------
Private Sub BuildMyProject()
    Dim ws As Worksheet, x As Double
    Set ws = GetWS(SH_MP)
    ResetSheet ws
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("A").Hidden = True
    ws.Columns("B").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 7
    ws.Range(ws.Columns(4), ws.Columns(4 + 30)).ColumnWidth = 3.3
    ws.Rows(1).RowHeight = 28
    ws.Range("B1").Value = "내 프로젝트 흐름 (주 단위)"
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Bold = True
    ws.Range("C2").Value = "프로젝트"
    ws.Range("C2").HorizontalAlignment = xlRight
    ws.Range("D2:P2").Merge
    ws.Range(MP_FILTER).Value = FILTER_ALL
    AddListValidation ws.Range(MP_FILTER), "=lstFilter"
    ws.Range("D2:P2").Interior.Color = RGB(255, 242, 204)
    ws.Range("B3").Font.Bold = True
    x = ws.Range("D1").Left
    x = AddButton(ws, "btnPhaseEdit", "단계 편집 (PM)", "EditPhasesButton", x, 3, 90, RGB(47, 117, 181))
    x = AddButton(ws, "btnGoCal", "캘린더로", "GoCalendarButton", x, 3, 60, RGB(89, 89, 89))
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("D5").Select
    ActiveWindow.FreezePanes = True
End Sub

'--- WeeklyReport -------------------------------------------
Private Sub BuildReport()
    Dim ws As Worksheet
    Set ws = GetWS(SH_RPT)
    ResetSheet ws
    ws.Columns("A").ColumnWidth = 90
    ws.Columns("B").ColumnWidth = 2
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D:E").ColumnWidth = 36
    ws.Columns("F").ColumnWidth = 36
    ws.Rows(1).RowHeight = 28
    ws.Range("A1").Value = "주간보고 초안"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    AddButton ws, "btnMakeReport", "주간보고 만들기", "MakeWeeklyReport", ws.Range("C1").Left, 3, 100, RGB(84, 130, 53)
End Sub

'--- CalMap (숨김) ------------------------------------------
Private Sub BuildMap()
    Dim ws As Worksheet
    Set ws = GetWS(SH_MAP)
    ResetSheet ws
    ws.Range("A1").Value = "캘린더 막대 ↔ TaskID 매핑 (VBA 전용)"
    ws.Visible = xlSheetVeryHidden
    GetWS(SH_SET).Visible = xlSheetHidden
End Sub

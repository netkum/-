Attribute VB_Name = "modConfig"
'==============================================================
' 팀장 Master - 공통 상수 / 설정 / 유틸리티
'  (팀원 템플릿과 같은 규칙: TaskDB 열 순서, 특수 프로젝트 코드 등)
'==============================================================
Option Explicit

Public Const APP_VERSION As String = "1.0.0"
Public Const APP_TITLE As String = "팀 업무관리 Master"
Public Const MEMBER_PATTERN As String = "*_업무일정.xlsm"

'--- 시트 이름 -------------------------------------------------
Public Const SH_DASH As String = "Dashboard"
Public Const SH_PV As String = "ProjectView"
Public Const SH_TS As String = "TeamSchedule"
Public Const SH_ALL As String = "TaskAll"
Public Const SH_PM As String = "ProjectMaster"
Public Const SH_SET As String = "Setting"

'--- 표 이름 ---------------------------------------------------
Public Const TB_ALL As String = "tblAll"
Public Const TB_SET As String = "tblSetting"
Public Const TB_MEMBER As String = "tblMember"
Public Const TB_EVENT As String = "tblTeamEvent"
Public Const TB_HOL As String = "tblHoliday"
Public Const TB_LOG As String = "tblSyncLog"
Public Const TB_PROJ As String = "tblProject"
Public Const TB_PHASE As String = "tblPhase"
Public Const TB_REQ As String = "tblRequest"
Public Const TB_MAP As String = "tblCodeMap"
Public Const TB_CMT As String = "tblComment"

'--- 팀원 TaskDB 열 (1~18) + Master 계산 열 (19~25) ------------
Public Const T_ID As Long = 1
Public Const T_PID As Long = 2
Public Const T_PNAME As Long = 3
Public Const T_PHASE As Long = 4
Public Const T_NAME As Long = 5
Public Const T_START As Long = 6
Public Const T_DUE As Long = 7
Public Const T_PCT As Long = 8
Public Const T_STATUS As Long = 9
Public Const T_PRI As Long = 10
Public Const T_NOTE As Long = 11
Public Const T_OWNER As Long = 12
Public Const T_PLAN As Long = 13
Public Const T_DONE As Long = 14
Public Const T_CREATED As Long = 15
Public Const T_UPDATED As Long = 16
Public Const T_COMMENT As Long = 17
Public Const T_MILE As Long = 18
Public Const T_COLS As Long = 18
Public Const A_SRC As Long = 19       ' 수집한 파일의 담당자(파일명 기준)
Public Const A_SYNC As Long = 20      ' 수집 일시
Public Const A_LATE As Long = 21      ' 지연 Y
Public Const A_PLANPCT As Long = 22   ' 계획 진척률
Public Const A_GAP As Long = 23       ' 실적 - 계획
Public Const A_DDAY As Long = 24      ' 종료일 - 오늘
Public Const A_STALE As Long = 25     ' 마지막 수정 후 경과일
Public Const A_COLS As Long = 25

'--- tblProject ------------------------------------------------
Public Const PR_ID As Long = 1
Public Const PR_NAME As Long = 2
Public Const PR_PM As Long = 3        ' 팀장이 지정한 PM
Public Const PR_AUTOPM As Long = 4    ' 자동 (1인 프로젝트 담당자 / 지정 필요)
Public Const PR_PART As Long = 5      ' 참여자 (자동)
Public Const PR_START As Long = 6
Public Const PR_END As Long = 7
Public Const PR_COLOR As Long = 8     ' 셀에 색을 칠하면 그 색 사용
Public Const PR_STATUS As Long = 9
Public Const PR_COLS As Long = 9

'--- tblPhase (팀원 tblMyPhase와 동일) -------------------------
Public Const PH_ID As Long = 1
Public Const PH_PID As Long = 2
Public Const PH_ORDER As Long = 3
Public Const PH_NAME As Long = 4
Public Const PH_START As Long = 5
Public Const PH_END As Long = 6
Public Const PH_AUTHOR As Long = 7
Public Const PH_MOD As Long = 8
Public Const PH_COLS As Long = 8

'--- tblRequest ------------------------------------------------
Public Const RQ_CODE As Long = 1
Public Const RQ_OWNER As Long = 2
Public Const RQ_NAME As Long = 3
Public Const RQ_START As Long = 4
Public Const RQ_END As Long = 5
Public Const RQ_MEMO As Long = 6
Public Const RQ_TASKS As Long = 7
Public Const RQ_DATE As Long = 8
Public Const RQ_ACTION As Long = 9    ' 승인 / 기존 통합 / 반려
Public Const RQ_TARGET As Long = 10   ' 정식 코드 / 통합 대상
Public Const RQ_REASON As Long = 11
Public Const RQ_DONE As Long = 12     ' 처리 일시
Public Const RQ_COLS As Long = 12

'--- tblCodeMap / tblComment / tblMember / tblTeamEvent --------
Public Const CM_COLS As Long = 6      ' From, To, Type, Result, Reason, Date
Public Const CO_ID As Long = 1
Public Const CO_TEXT As Long = 2
Public Const CO_DATE As Long = 3
Public Const CO_OWNER As Long = 4
Public Const CO_TASK As Long = 5
Public Const CO_COLS As Long = 5
Public Const M_NAME As Long = 1
Public Const M_FILE As Long = 2
Public Const M_SYNC As Long = 3
Public Const M_RESULT As Long = 4
Public Const M_ROWS As Long = 5
Public Const M_LASTUPD As Long = 6
Public Const M_VER As Long = 7
Public Const M_COLS As Long = 7

'--- 상태값 / 특수 프로젝트 -----------------------------------
Public Const ST_PLAN As String = "예정"
Public Const ST_DOING As String = "진행"
Public Const ST_DONE As String = "완료"
Public Const ST_HOLD As String = "보류"
Public Const ST_CANCEL As String = "취소"
Public Const PJ_COMMON As String = "공통"
Public Const PJ_LEAVE As String = "휴가"
Public Const PJ_TRIP As String = "출장"
Public Const PJ_EDU As String = "교육"
Public Const FILTER_ALL As String = "전체"
Public Const PV_FILTER As String = "C2"
Public Const TS_MONTH As String = "C2"

Public Const SYM_MILE As String = "◆"
Public Const SYM_DELAY As String = "▲"
Public Const SYM_DONE As String = "√"
Public Const SYM_COMMENT As String = "※"

Private mSpeed As Long

Public Function IsReady() As Boolean
    IsReady = Not GetLO(TB_ALL) Is Nothing And Not GetLO(TB_SET) Is Nothing And Not GetLO(TB_PROJ) Is Nothing
End Function

Public Function MemberFolder() As String
    Dim p As String
    p = CStr(CfgGet("MemberFolder", ""))
    If p = "" Then
        p = ThisWorkbook.Path
        If InStrRev(p, "\") > 0 Then p = Left$(p, InStrRev(p, "\") - 1)
        p = p & "\Member"
    End If
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    MemberFolder = p
End Function

' ProjectList.xlsx 를 저장할 폴더 (기본: 이 파일이 있는 Master 폴더)
Public Function ExportFolder() As String
    Dim p As String
    p = CStr(CfgGet("ExportFolder", ""))
    If p = "" Then p = ThisWorkbook.Path
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    ExportFolder = p
End Function

' 프로젝트 기본 색 (셀에 색이 없을 때)
Public Function PaletteColor(ByVal pid As String) As Long
    Dim pal As Variant, i As Long, h As Long
    If IsAbsence(pid) Then PaletteColor = RGB(166, 166, 166): Exit Function
    If pid = PJ_COMMON Then PaletteColor = RGB(128, 128, 128): Exit Function
    pal = Array(RGB(68, 114, 196), RGB(237, 125, 49), RGB(112, 173, 71), RGB(191, 144, 0), _
                RGB(91, 155, 213), RGB(165, 105, 189), RGB(38, 166, 154), RGB(214, 69, 65))
    For i = 1 To Len(pid): h = h + AscW(Mid$(pid, i, 1)): Next i
    PaletteColor = pal(Abs(h) Mod 8)
End Function

'==============================================================
' 공통 유틸리티 (팀원 템플릿 modConfig와 동일)
'==============================================================
Public Function GetWS(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWS = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function GetLO(ByVal tblName As String) As ListObject
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        On Error Resume Next
        Set GetLO = sh.ListObjects(tblName)
        On Error GoTo 0
        If Not GetLO Is Nothing Then Exit Function
    Next sh
End Function

' 표 데이터 → 2차원 배열(1-based). 데이터 없으면 Empty
Public Function ReadTable(ByVal lo As ListObject) As Variant
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ReadTable = lo.DataBodyRange.Value
End Function

Public Function RowCount(ByVal data As Variant) As Long
    If IsEmpty(data) Then RowCount = 0 Else RowCount = UBound(data, 1)
End Function

' 표 전체 교체
Public Sub WriteTable(ByVal lo As ListObject, ByVal data As Variant)
    Dim n As Long, c As Long
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    If IsEmpty(data) Then Exit Sub
    n = UBound(data, 1) - LBound(data, 1) + 1
    c = UBound(data, 2) - LBound(data, 2) + 1
    If n <= 0 Then Exit Sub
    lo.Resize lo.HeaderRowRange.Resize(n + 1)
    lo.DataBodyRange.Resize(n, c).Value = data
End Sub

' 1차원 배열 1행 추가 → 추가된 행 번호(DataBody 기준)
Public Function AppendRow(ByVal lo As ListObject, ByVal v As Variant) As Long
    Dim lr As ListRow
    Set lr = lo.ListRows.Add
    lr.Range.Resize(1, UBound(v) - LBound(v) + 1).Value = v
    AppendRow = lr.Index
End Function

Public Function RowValues(ByVal lo As ListObject, ByVal idx As Long) As Variant
    Dim a As Variant, r() As Variant, i As Long, n As Long
    a = lo.DataBodyRange.Rows(idx).Value
    n = UBound(a, 2)
    ReDim r(1 To n)
    For i = 1 To n: r(i) = a(1, i): Next i
    RowValues = r
End Function

Public Sub SetRowValues(ByVal lo As ListObject, ByVal idx As Long, ByVal v As Variant)
    lo.DataBodyRange.Rows(idx).Resize(1, UBound(v) - LBound(v) + 1).Value = v
End Sub

' 표의 특정 열에서 key를 찾아 DataBody 행 번호 반환 (없으면 0)
Public Function FindRow(ByVal lo As ListObject, ByVal colIdx As Long, ByVal key As String) As Long
    Dim r As Variant
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If key = "" Then Exit Function
    r = Application.Match(key, lo.ListColumns(colIdx).DataBodyRange, 0)
    If Not IsError(r) Then FindRow = CLng(r)
End Function

' 컬렉션(1차원 배열 모음) → 2차원 배열
Public Function CollTo2D(ByVal coll As Collection, ByVal nCols As Long) As Variant
    Dim a() As Variant, i As Long, j As Long, v As Variant
    If coll.Count = 0 Then Exit Function
    ReDim a(1 To coll.Count, 1 To nCols)
    For i = 1 To coll.Count
        v = coll(i)
        For j = 1 To nCols
            a(i, j) = v(LBound(v) + j - 1)
        Next j
    Next i
    CollTo2D = a
End Function

' 2차원 배열의 i행 → 1차원 배열(1-based)
Public Function RowOf(ByVal data As Variant, ByVal i As Long) As Variant
    Dim r() As Variant, j As Long
    ReDim r(1 To UBound(data, 2))
    For j = 1 To UBound(data, 2): r(j) = data(i, j): Next j
    RowOf = r
End Function

Public Function CfgGet(ByVal key As String, Optional ByVal def As Variant = "") As Variant
    Dim lo As ListObject, r As Long, v As Variant
    Set lo = GetLO(TB_SET)
    r = FindRow(lo, 1, key)
    If r = 0 Then CfgGet = def: Exit Function
    v = lo.DataBodyRange.Cells(r, 2).Value
    If IsEmpty(v) Then
        CfgGet = def
    ElseIf CStr(v) = "" Then
        CfgGet = def
    Else
        CfgGet = v
    End If
End Function

Public Sub CfgSet(ByVal key As String, ByVal v As Variant)
    Dim lo As ListObject, r As Long
    Set lo = GetLO(TB_SET)
    If lo Is Nothing Then Exit Sub
    r = FindRow(lo, 1, key)
    If r = 0 Then
        AppendRow lo, Array(key, v)
    Else
        lo.DataBodyRange.Cells(r, 2).Value = v
    End If
End Sub

Public Sub SpeedOn()
    mSpeed = mSpeed + 1
    If mSpeed = 1 Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    End If
End Sub

Public Sub SpeedOff()
    If mSpeed > 0 Then mSpeed = mSpeed - 1
    If mSpeed = 0 Then
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub

Public Sub SpeedReset()
    mSpeed = 0
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Function IsDateVal(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        IsDateVal = True
    ElseIf IsNumeric(v) Then
        IsDateVal = (CDbl(v) > 1)
    ElseIf VarType(v) = vbString Then
        If Trim$(v) <> "" Then IsDateVal = IsDate(v)
    End If
End Function

Public Function ToDate(ByVal v As Variant) As Date
    If Not IsDateVal(v) Then Exit Function
    If VarType(v) = vbDate Then
        ToDate = Int(CDbl(v))
    ElseIf IsNumeric(v) Then
        ToDate = CDate(Int(CDbl(v)))
    Else
        ToDate = Int(CDbl(CDate(v)))
    End If
End Function

' "2026-10-07", "10/7", "10-7", "10.7" 등을 날짜로 (실패 시 0)
Public Function ParseDate(ByVal s As String) As Date
    s = Trim$(Replace(s, ".", "-"))
    If s = "" Then Exit Function
    If IsDate(s) Then
        ParseDate = Int(CDbl(CDate(s)))
    ElseIf IsDate(Year(Date) & "-" & s) Then
        ParseDate = Int(CDbl(CDate(Year(Date) & "-" & s)))
    End If
End Function

Public Function WeekMonday(ByVal d As Date) As Date
    WeekMonday = d - Weekday(d, vbMonday) + 1
End Function

Public Function HolidayRange() As Range
    Dim lo As ListObject
    Set lo = GetLO(TB_HOL)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set HolidayRange = lo.ListColumns(1).DataBodyRange
End Function

' 공휴일 제외 작업일수 (최소 1)
Public Function WorkDays(ByVal s As Date, ByVal e As Date) As Long
    Dim h As Range, n As Long
    If e < s Then e = s
    Set h = HolidayRange()
    On Error Resume Next
    If h Is Nothing Then
        n = Application.WorksheetFunction.NetworkDays(s, e)
    Else
        n = Application.WorksheetFunction.NetworkDays(s, e, h)
    End If
    On Error GoTo 0
    If n < 1 Then n = 1
    WorkDays = n
End Function

' 공휴일 사전 (key = CLng(날짜), value = 이름)
Public Function HolidayDict() As Object
    Dim d As Object, a As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    a = ReadTable(GetLO(TB_HOL))
    For i = 1 To RowCount(a)
        If IsDateVal(a(i, 1)) Then d(CLng(ToDate(a(i, 1)))) = CStr(a(i, 2))
    Next i
    Set HolidayDict = d
End Function

Public Function DDayText(ByVal due As Date) As String
    Dim n As Long
    n = due - Date
    If n = 0 Then
        DDayText = "D-day"
    ElseIf n > 0 Then
        DDayText = "D-" & n
    Else
        DDayText = "D+" & (-n)
    End If
End Function

Public Function IsAbsence(ByVal pid As String) As Boolean
    IsAbsence = (pid = PJ_LEAVE Or pid = PJ_TRIP Or pid = PJ_EDU)
End Function

Public Function IsSpecialProject(ByVal pid As String) As Boolean
    IsSpecialProject = (pid = PJ_COMMON Or IsAbsence(pid))
End Function

' 색을 흰색 쪽으로 섞어 연한 톤 만들기 (f: 0~1, 클수록 연함)
Public Function LightColor(ByVal c As Long, Optional ByVal f As Double = 0.7) As Long
    Dim r As Long, g As Long, b As Long
    r = c Mod 256: g = (c \ 256) Mod 256: b = (c \ 65536) Mod 256
    r = r + (255 - r) * f: g = g + (255 - g) * f: b = b + (255 - b) * f
    LightColor = RGB(r, g, b)
End Function

Public Function Nz(ByVal v As Variant, Optional ByVal def As Variant = "") As Variant
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        Nz = def
    Else
        Nz = v
    End If
End Function

Public Function NumOr(ByVal v As Variant, Optional ByVal def As Double = 0) As Double
    If IsEmpty(v) Or IsError(v) Then
        NumOr = def
    ElseIf VarType(v) = vbDate Then
        NumOr = CDbl(v)
    ElseIf IsNumeric(v) Then
        NumOr = CDbl(v)
    Else
        NumOr = def
    End If
End Function

Public Sub Msg(ByVal text As String, Optional ByVal icon As VbMsgBoxStyle = vbInformation)
    MsgBox text, icon, APP_TITLE
End Sub

' 필터 셀 값 "P001 저도주..." → "P001" / "전체" → ""
Public Function FilterPid(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(Nz(v, "")))
    If s = "" Or s = FILTER_ALL Then Exit Function
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
    FilterPid = s
End Function

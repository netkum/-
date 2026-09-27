Attribute VB_Name = "modConfig"
'==============================================================
' 팀원용 업무일정 템플릿 - 공통 상수 / 설정 / 유틸리티
'==============================================================
Option Explicit

Public Const APP_VERSION As String = "1.0.0"
Public Const APP_TITLE As String = "업무일정"

'--- 시트 이름 -------------------------------------------------
Public Const SH_CAL As String = "Calendar"
Public Const SH_MP As String = "MyProject"
Public Const SH_DB As String = "TaskDB"
Public Const SH_RPT As String = "WeeklyReport"
Public Const SH_SET As String = "Setting"
Public Const SH_MAP As String = "CalMap"

'--- 표 이름 ---------------------------------------------------
Public Const TB_TASK As String = "tblTask"
Public Const TB_SET As String = "tblSetting"
Public Const TB_PROJ As String = "tblProject"
Public Const TB_PHASE As String = "tblPhase"
Public Const TB_MYPHASE As String = "tblMyPhase"
Public Const TB_PHMAP As String = "tblMyPhaseMap"
Public Const TB_REQ As String = "tblMyRequest"
Public Const TB_EVENT As String = "tblEvent"
Public Const TB_HOL As String = "tblHoliday"
Public Const TB_MEMO As String = "tblMemo"

'--- TaskDB 열 번호 (팀장 Master도 같은 순서로 읽음) -----------
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

'--- tblProject 열 --------------------------------------------
Public Const P_ID As Long = 1
Public Const P_NAME As Long = 2
Public Const P_PM As Long = 3
Public Const P_COLOR As Long = 4
Public Const P_STATUS As Long = 5

'--- tblPhase / tblMyPhase 열 ---------------------------------
Public Const PH_ID As Long = 1
Public Const PH_PID As Long = 2
Public Const PH_ORDER As Long = 3
Public Const PH_NAME As Long = 4
Public Const PH_START As Long = 5
Public Const PH_END As Long = 6
Public Const PH_AUTHOR As Long = 7
Public Const PH_MOD As Long = 8
Public Const PH_COLS As Long = 8

'--- tblMyRequest 열 ------------------------------------------
Public Const RQ_CODE As Long = 1
Public Const RQ_NAME As Long = 2
Public Const RQ_START As Long = 3
Public Const RQ_END As Long = 4
Public Const RQ_MEMO As Long = 5
Public Const RQ_DATE As Long = 6
Public Const RQ_RESULT As Long = 7
Public Const RQ_MAPPED As Long = 8
Public Const RQ_REASON As Long = 9
Public Const RQ_COLS As Long = 9

'--- 상태값 ---------------------------------------------------
Public Const ST_PLAN As String = "예정"
Public Const ST_DOING As String = "진행"
Public Const ST_DONE As String = "완료"
Public Const ST_HOLD As String = "보류"
Public Const ST_CANCEL As String = "취소"

'--- 특수 프로젝트 코드 ---------------------------------------
Public Const PJ_COMMON As String = "공통"
Public Const PJ_LEAVE As String = "휴가"
Public Const PJ_TRIP As String = "출장"
Public Const PJ_EDU As String = "교육"
Public Const PJ_PENDING As String = "승인대기"

'--- 캘린더 격자 (모든 파일 동일) -----------------------------
Public Const CAL_TOP As Long = 4          ' 첫 주 블록 시작 행
Public Const CAL_BLOCK As Long = 9        ' 주 블록 1개 행 수
Public Const CAL_WEEKS As Long = 6        ' 최대 주 수
Public Const CAL_LANES As Long = 5        ' 업무 막대 줄 수
Public Const OFF_DATE As Long = 0         ' 블록 내 날짜 행
Public Const OFF_REPORT As Long = 1       ' 보고 줄
Public Const OFF_LANE As Long = 2         ' 첫 막대 줄
Public Const OFF_MEMO As Long = 7         ' 첫 메모 줄
Public Const MEMO_LINES As Long = 2
Public Const COL_WEEK As Long = 1         ' A 주차
Public Const COL_DAY1 As Long = 2         ' B 월 ~ F 금
Public Const COL_WKND As Long = 7         ' G 토·일
Public Const COL_SUM As Long = 8          ' H 주간 요약
Public Const COL_DETAIL As Long = 10      ' J 상세 패널
Public Const DETAIL_ROWS As Long = 22
Public Const CAL_FILTER As String = "C2"
Public Const MP_FILTER As String = "D2"
Public Const FILTER_ALL As String = "전체"

' 기호 (CP949에 있는 문자만 사용)
Public Const SYM_MILE As String = "◆"
Public Const SYM_DELAY As String = "▲"
Public Const SYM_DONE As String = "√"
Public Const SYM_COMMENT As String = "※"
Public Const SYM_NEXT As String = "▶"
Public Const SYM_PREV As String = "◀"

Private mSpeed As Long

'==============================================================
' 시트 / 표 접근
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

Public Function IsReady() As Boolean
    IsReady = Not GetLO(TB_SET) Is Nothing And Not GetLO(TB_TASK) Is Nothing _
              And Not GetWS(SH_CAL) Is Nothing And Not GetWS(SH_MAP) Is Nothing
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

'==============================================================
' 설정 (Setting 시트 tblSetting: Key / Value)
'==============================================================
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

Public Function OwnerName() As String
    OwnerName = CStr(CfgGet("OwnerName", ""))
End Function

Public Function MasterFolder() As String
    Dim p As String
    p = CStr(CfgGet("MasterFolder", ""))
    If p = "" Then
        ' 기본값: ...\업무관리\Member\ 의 상위 폴더 아래 Master
        p = ThisWorkbook.Path
        If InStrRev(p, "\") > 0 Then p = Left$(p, InStrRev(p, "\") - 1)
        p = p & "\Master"
    End If
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    MasterFolder = p
End Function

'==============================================================
' 속도 최적화 (중첩 호출 가능)
'==============================================================
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

'==============================================================
' 날짜
'==============================================================
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

Public Function AddWorkDays(ByVal d As Date, ByVal n As Long) As Date
    Dim h As Range
    Set h = HolidayRange()
    On Error Resume Next
    If h Is Nothing Then
        AddWorkDays = Application.WorksheetFunction.WorkDay(d, n)
    Else
        AddWorkDays = Application.WorksheetFunction.WorkDay(d, n, h)
    End If
    If Err.Number <> 0 Then AddWorkDays = d + n
    On Error GoTo 0
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

' 주차 표시: "그 주 목요일이 속한 달의 n주차"
Public Function WeekLabel(ByVal monday As Date) As String
    Dim thu As Date
    thu = monday + 3
    WeekLabel = Month(thu) & "월" & vbLf & ((Day(thu) - 1) \ 7 + 1) & "주차"
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

'==============================================================
' 프로젝트 정보
'==============================================================
Public Function IsAbsence(ByVal pid As String) As Boolean
    IsAbsence = (pid = PJ_LEAVE Or pid = PJ_TRIP Or pid = PJ_EDU)
End Function

Public Function IsSpecialProject(ByVal pid As String) As Boolean
    IsSpecialProject = (pid = PJ_COMMON Or IsAbsence(pid))
End Function

' key = ProjectID, value = Array(이름, PM, 색, 상태)
Public Function ProjectDict() As Object
    Dim d As Object, a As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    a = ReadTable(GetLO(TB_PROJ))
    For i = 1 To RowCount(a)
        If CStr(a(i, P_ID)) <> "" Then
            d(CStr(a(i, P_ID))) = Array(CStr(a(i, P_NAME)), CStr(a(i, P_PM)), a(i, P_COLOR), CStr(a(i, P_STATUS)))
        End If
    Next i
    Set ProjectDict = d
End Function

Public Function ProjectNameOf(ByVal pid As String, Optional ByVal pd As Object = Nothing) As String
    If IsSpecialProject(pid) Then ProjectNameOf = pid: Exit Function
    If pd Is Nothing Then Set pd = ProjectDict()
    If pd.Exists(pid) Then ProjectNameOf = pd(pid)(0)
End Function

' 캘린더 표시용 코드 (승인대기 임시 프로젝트는 * 표시)
Public Function ProjectLabel(ByVal pid As String, Optional ByVal pd As Object = Nothing) As String
    If pd Is Nothing Then Set pd = ProjectDict()
    ProjectLabel = pid
    If pd.Exists(pid) Then
        If pd(pid)(3) = PJ_PENDING Then ProjectLabel = pid & "*"
    End If
End Function

Public Function ProjectColor(ByVal pid As String, Optional ByVal pd As Object = Nothing) As Long
    Dim pal As Variant, i As Long, h As Long
    If IsAbsence(pid) Then ProjectColor = RGB(166, 166, 166): Exit Function
    If pid = PJ_COMMON Then ProjectColor = RGB(128, 128, 128): Exit Function
    If pd Is Nothing Then Set pd = ProjectDict()
    If pd.Exists(pid) Then
        If IsNumeric(pd(pid)(2)) And CStr(pd(pid)(2)) <> "" Then
            ProjectColor = CLng(pd(pid)(2)): Exit Function
        End If
    End If
    pal = Array(RGB(68, 114, 196), RGB(237, 125, 49), RGB(112, 173, 71), RGB(191, 144, 0), _
                RGB(91, 155, 213), RGB(165, 105, 189), RGB(38, 166, 154), RGB(214, 69, 65))
    For i = 1 To Len(pid): h = h + AscW(Mid$(pid, i, 1)): Next i
    ProjectColor = pal(Abs(h) Mod 8)
End Function

' 색을 흰색 쪽으로 섞어 연한 톤 만들기 (f: 0~1, 클수록 연함)
Public Function LightColor(ByVal c As Long, Optional ByVal f As Double = 0.7) As Long
    Dim r As Long, g As Long, b As Long
    r = c Mod 256: g = (c \ 256) Mod 256: b = (c \ 65536) Mod 256
    r = r + (255 - r) * f: g = g + (255 - g) * f: b = b + (255 - b) * f
    LightColor = RGB(r, g, b)
End Function

'==============================================================
' 기타
'==============================================================
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

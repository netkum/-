Attribute VB_Name = "modEvents"
'==============================================================
' ThisWorkbook 이벤트에서 호출하는 처리 (시트 모듈에 코드 불필요)
'==============================================================
Option Explicit

Public Sub OnOpen()
    If Not IsReady() Then Exit Sub
    EnsureOwner
    CfgSet "CurYear", Year(Date)
    CfgSet "CurMonth", Month(Date)
    SyncFromMaster True
    RenderCalendar
    GoThisWeek
End Sub

' 처음 열 때 이름·이니셜 등록
Public Sub EnsureOwner()
    Dim nm As String, ini As String, def As String
    If OwnerName() <> "" And CStr(CfgGet("Initial", "")) <> "" Then Exit Sub
    If InStr(1, ThisWorkbook.Name, "Template", vbTextCompare) > 0 Then Exit Sub   ' 원본 템플릿은 건너뜀
    def = ThisWorkbook.Name
    If InStr(def, "_") > 0 Then def = Left$(def, InStr(def, "_") - 1) Else def = ""
    nm = InputBox("처음 사용하시네요. 이름을 입력하세요." & vbLf & "(팀장 화면에 담당자로 표시됩니다)", APP_TITLE, def)
    If Trim$(nm) = "" Then Exit Sub
    ini = InputBox("Task 번호 앞에 붙일 영문 이니셜을 입력하세요. (예: KIM)", APP_TITLE, "")
    ini = UCase$(Trim$(ini))
    If ini = "" Then ini = "T"
    CfgSet "OwnerName", Trim$(nm)
    CfgSet "Initial", ini
    RefreshLists
End Sub

Public Sub OnBeforeSave()
    If Not IsReady() Then Exit Sub
    If CStr(CfgGet("CalDirty", "")) = "Y" Then RenderCalendar
End Sub

Public Sub OnSheetChange(ByVal sh As Object, ByVal target As Range)
    Dim c As Range, d As Date, kind As String, n As Long
    If Not IsReady() Then Exit Sub
    On Error GoTo EH
    Select Case sh.Name
        Case SH_DB
            FixEditedRows target
            CfgSet "CalDirty", "Y"
        Case SH_CAL
            If Not Intersect(target, sh.Range(CAL_FILTER)) Is Nothing Then
                RenderCalendar
                Exit Sub
            End If
            For Each c In target.Cells
                n = n + 1
                If n > 60 Then Exit For
                kind = CellKind(c.Row, c.Column, d)
                If kind = "memo" Then
                    Application.EnableEvents = False
                    SaveMemoCell c
                    Application.EnableEvents = True
                ElseIf kind = "lane" Or kind = "date" Or kind = "report" Then
                    Msg "막대·날짜 칸은 직접 입력할 수 없습니다." & vbLf & _
                        "날짜 칸을 더블클릭해서 Task를 등록하세요. (자유 메모는 아래 2줄에 쓰세요)", vbExclamation
                    RenderCalendar
                    Exit Sub
                End If
            Next c
        Case SH_MP
            If Not Intersect(target, sh.Range(MP_FILTER)) Is Nothing Then RenderMyProject
    End Select
    Exit Sub
EH:
    SpeedReset
End Sub

Public Sub OnSelectionChange(ByVal sh As Object, ByVal target As Range)
    Dim d As Date, kind As String, id As String, top As Long, b As Long, lane As Long
    If sh.Name <> SH_CAL Then Exit Sub
    If Not IsReady() Then Exit Sub
    If target.Cells.CountLarge > 1 Then Exit Sub
    On Error GoTo EH
    kind = CellKind(target.Row, target.Column, d, lane, b)
    If kind = "" Or kind = "memo" Then Exit Sub
    top = BlockTop(b)
    If kind = "lane" Then id = CStr(Nz(GetWS(SH_MAP).Cells(target.Row, target.Column).Value))
    If id <> "" Then
        ShowDetail id, top
    Else
        ShowDayList d, top
    End If
    Exit Sub
EH:
    SpeedReset
End Sub

Public Sub OnDoubleClick(ByVal sh As Object, ByVal target As Range, ByRef cancel As Boolean)
    Dim d As Date, kind As String, id As String
    If Not IsReady() Then Exit Sub
    If sh.Name = SH_MP Then MPDoubleClick target, cancel: Exit Sub
    If sh.Name <> SH_CAL Then Exit Sub
    kind = CellKind(target.Row, target.Column, d)
    Select Case kind
        Case "memo"
            Exit Sub                    ' 메모 줄은 그냥 셀 편집
        Case "lane"
            cancel = True
            id = CStr(Nz(GetWS(SH_MAP).Cells(target.Row, target.Column).Value))
            If id <> "" Then OpenTaskForm id Else OpenTaskForm "", d
        Case "date"
            cancel = True
            OpenTaskForm "", d
        Case "report"
            cancel = True
            OpenTaskForm "", d, True
        Case Else
            If target.Column = COL_SUM Or target.Column = COL_WEEK Or target.Column = COL_DETAIL Then cancel = True
    End Select
End Sub

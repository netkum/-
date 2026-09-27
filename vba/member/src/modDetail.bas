Attribute VB_Name = "modDetail"
'==============================================================
' 캘린더 오른쪽 Task 상세 패널 (J열)
'  - 막대 한 번 클릭 → Task 상세
'  - 빈 칸/날짜 클릭 → 그날 전체 목록
'  - 패널은 클릭한 주 블록 옆에 표시 (스크롤해도 보이도록)
'==============================================================
Option Explicit

Private Const BTN_EDIT As String = "btnDetailEdit"
Private Const BTN_DONE As String = "btnDetailDone"
Private Const BTN_EXT As String = "btnDetailExtend"

Public Sub ClearDetailPanel()
    Dim ws As Worksheet, rng As Range
    Set ws = GetWS(SH_CAL)
    Set rng = ws.Range(ws.Cells(3, COL_DETAIL), ws.Cells(LastCalRow() + DETAIL_ROWS, COL_DETAIL))
    rng.ClearContents
    rng.Interior.Pattern = xlNone
    rng.Font.Bold = False
    rng.Font.Color = RGB(38, 38, 38)
    rng.Borders.LineStyle = xlNone
    ShowButtons False, 0
    CfgSet "DetailTaskID", ""
End Sub

Private Sub ShowButtons(ByVal visible As Boolean, ByVal atRow As Long)
    Dim ws As Worksheet, nm As Variant, x As Double, shp As Shape
    Set ws = GetWS(SH_CAL)
    x = ws.Cells(1, COL_DETAIL).Left + 2
    For Each nm In Array(BTN_EDIT, BTN_DONE, BTN_EXT)
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(CStr(nm))
        On Error GoTo 0
        If Not shp Is Nothing Then
            shp.Visible = IIf(visible, msoTrue, msoFalse)
            If visible Then
                shp.Top = ws.Cells(atRow, COL_DETAIL).Top + 2
                shp.Left = x
                x = x + shp.Width + 4
            End If
        End If
    Next nm
End Sub

' 패널에 줄 단위로 쓰기
Private Sub WritePanel(ByVal topRow As Long, ByVal lines As Collection, ByVal title As String)
    Dim ws As Worksheet, i As Long, r As Long
    Set ws = GetWS(SH_CAL)
    SpeedOn
    ClearDetailPanel
    ws.Cells(topRow, COL_DETAIL).Value = "'" & title
    With ws.Cells(topRow, COL_DETAIL)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 84, 106)
    End With
    r = topRow + 1
    For i = 1 To lines.Count
        If r > topRow + DETAIL_ROWS - 2 Then Exit For
        ws.Cells(r, COL_DETAIL).Value = "'" & lines(i)
        If Left$(lines(i), 1) = "─" Then ws.Cells(r, COL_DETAIL).Font.Color = RGB(128, 128, 128)
        r = r + 1
    Next i
    With ws.Range(ws.Cells(topRow, COL_DETAIL), ws.Cells(r + 1, COL_DETAIL))
        .BorderAround xlContinuous, xlThin, , RGB(166, 166, 166)
        .Offset(1, 0).Resize(.Rows.Count - 1, 1).Interior.Color = RGB(250, 250, 250)
    End With
    ws.Cells(r + 1, COL_DETAIL).Interior.Color = RGB(250, 250, 250)
    CfgSet "DetailRow", r + 1
    SpeedOff
End Sub

'--- 1) Task 상세 -------------------------------------------
Public Sub ShowDetail(ByVal id As String, ByVal topRow As Long)
    Dim t As Variant, i As Long, k As Long, lines As Collection, pd As Object
    Dim pid As String, s As Date, e As Date, pct As Double, st As String
    Dim prevI As Long, nextI As Long, same As Collection, nDone As Long

    t = LoadTasks()
    For k = 1 To RowCount(t)
        If CStr(t(k, T_ID)) = id Then i = k: Exit For
    Next k
    If i = 0 Then Exit Sub

    Set pd = ProjectDict()
    pid = CStr(t(i, T_PID))
    s = ToDate(t(i, T_START)): e = ToDate(t(i, T_DUE))
    pct = NumOr(t(i, T_PCT)): st = CStr(t(i, T_STATUS))

    Set lines = New Collection
    lines.Add "[" & ProjectLabel(pid, pd) & "] " & ProjectNameOf(pid, pd)
    If CStr(Nz(t(i, T_PHASE))) <> "" Then lines.Add "단계: " & PhaseNameOf(CStr(t(i, T_PHASE)))
    lines.Add CStr(t(i, T_NAME)) & "      " & id
    lines.Add Format$(s, "m/d(aaa)") & " ~ " & Format$(e, "m/d(aaa)") & "   " & NumOr(t(i, T_PLAN)) & "일"
    lines.Add "진척 " & pct & "% " & ProgressBar(pct) & "  " & IIf(st = ST_DONE, "", DDayText(e))
    lines.Add "상태: " & st & "   우선순위: " & Nz(t(i, T_PRI)) & IIf(UCase$(CStr(Nz(t(i, T_MILE)))) = "Y", "   " & SYM_MILE & "보고·마감", "")
    If IsDelayed(t, i) Then lines.Add SYM_DELAY & " 지연 중 (" & DDayText(e) & ")"
    If CStr(Nz(t(i, T_NOTE))) <> "" Then lines.Add "메모: " & t(i, T_NOTE)
    If CStr(Nz(t(i, T_COMMENT))) <> "" Then lines.Add SYM_COMMENT & " 팀장: " & t(i, T_COMMENT)

    ' 같은 프로젝트 앞뒤 Task
    If Not IsSpecialProject(pid) Then
        Set same = New Collection
        For k = 1 To RowCount(t)
            If IsValidTask(t, k) And CStr(t(k, T_PID)) = pid And CStr(t(k, T_STATUS)) <> ST_CANCEL Then
                same.Add k
                If CStr(t(k, T_STATUS)) = ST_DONE Then nDone = nDone + 1
                If k <> i Then
                    If ToDate(t(k, T_START)) < s Or (ToDate(t(k, T_START)) = s And k < i) Then
                        If prevI = 0 Then
                            prevI = k
                        ElseIf ToDate(t(k, T_START)) >= ToDate(t(prevI, T_START)) Then
                            prevI = k
                        End If
                    Else
                        If nextI = 0 Then
                            nextI = k
                        ElseIf ToDate(t(k, T_START)) < ToDate(t(nextI, T_START)) Then
                            nextI = k
                        End If
                    End If
                End If
            End If
        Next k
        lines.Add "─ 같은 프로젝트 앞뒤 Task ─"
        If prevI > 0 Then lines.Add IIf(CStr(t(prevI, T_STATUS)) = ST_DONE, SYM_DONE, "·") & " " & t(prevI, T_NAME) & _
                                    " (" & Format$(t(prevI, T_START), "m/d") & "~" & Format$(t(prevI, T_DUE), "m/d") & ")"
        lines.Add "● " & t(i, T_NAME) & " (지금)"
        If nextI > 0 Then lines.Add "→ " & t(nextI, T_NAME) & " (" & Format$(t(nextI, T_START), "m/d") & "~" & _
                                    Format$(t(nextI, T_DUE), "m/d") & ")"
        lines.Add "─ 프로젝트 현황 (내 Task 기준) ─"
        lines.Add pid & " 진척 " & Format$(WeightedPct(t, same), "0") & "% · Task " & same.Count & "건 중 " & nDone & "건 완료"
    End If

    WritePanel topRow, lines, "Task 상세"
    CfgSet "DetailTaskID", id
    ShowButtons True, CLng(CfgGet("DetailRow", topRow + 12))
End Sub

'--- 2) 그날 전체 목록 ---------------------------------------
Public Sub ShowDayList(ByVal d As Date, ByVal topRow As Long)
    Dim t As Variant, i As Long, lines As Collection, pd As Object, ev As Variant
    Dim memo As Object, k As Long, n As Long

    Set lines = New Collection
    Set pd = ProjectDict()

    ev = ReadTable(GetLO(TB_EVENT))
    For i = 1 To RowCount(ev)
        If IsDateVal(ev(i, 1)) Then
            If ToDate(ev(i, 1)) = d Then lines.Add SYM_MILE & " " & Trim$(Nz(ev(i, 2)) & " " & Nz(ev(i, 3)))
        End If
    Next i

    t = LoadTasks()
    For i = 1 To RowCount(t)
        If IsValidTask(t, i) Then
            If CStr(t(i, T_STATUS)) <> ST_CANCEL And ToDate(t(i, T_START)) <= d And ToDate(t(i, T_DUE)) >= d Then
                lines.Add BarLabel(t, i, pd) & " (" & Format$(t(i, T_START), "m/d") & "~" & Format$(t(i, T_DUE), "m/d") & ")"
                n = n + 1
            End If
        End If
    Next i
    If n = 0 Then lines.Add "(등록된 업무 없음)"

    Set memo = MemoDict()
    For k = 1 To MEMO_LINES
        If memo.Exists(MemoKey(d, k)) Then lines.Add "메모: " & memo(MemoKey(d, k))
    Next k
    lines.Add "─ 더블클릭하면 이 날짜로 새 Task 등록 ─"

    WritePanel topRow, lines, Format$(d, "m/d (aaa)") & " 일정  " & n & "건"
    ShowButtons False, 0
End Sub

Private Function ProgressBar(ByVal pct As Double) As String
    Dim n As Long
    n = Int(pct / 10 + 0.5)
    If n < 0 Then n = 0
    If n > 10 Then n = 10
    ProgressBar = String$(n, "■") & String$(10 - n, "□")
End Function

'--- 상세 패널 버튼 (OnAction) ------------------------------
Public Sub DetailEdit()
    Dim id As String
    id = CStr(CfgGet("DetailTaskID", ""))
    If id = "" Then Exit Sub
    OpenTaskForm id
End Sub

Public Sub DetailComplete()
    Dim id As String
    id = CStr(CfgGet("DetailTaskID", ""))
    If id = "" Then Exit Sub
    If MsgBox("이 업무를 완료 처리할까요?", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub
    CompleteTask id
    AfterTaskChanged id
End Sub

Public Sub DetailExtend()
    Dim id As String
    id = CStr(CfgGet("DetailTaskID", ""))
    If id = "" Then Exit Sub
    ExtendTask id, 1
    AfterTaskChanged id
End Sub

' Task 변경 후 캘린더 다시 그리고 상세 패널 유지
Public Sub AfterTaskChanged(ByVal id As String)
    Dim topRow As Long, ws As Worksheet
    Set ws = GetWS(SH_CAL)
    topRow = ActiveCell.Row
    If ActiveSheet.Name <> SH_CAL Then topRow = CAL_TOP
    topRow = CAL_TOP + ((topRow - CAL_TOP) \ CAL_BLOCK) * CAL_BLOCK
    If topRow < CAL_TOP Then topRow = CAL_TOP
    RenderCalendar
    If id <> "" Then
        If Not IsEmpty(GetTaskArr(id)) Then ShowDetail id, topRow
    End If
End Sub

'--- 입력폼 열기 ---------------------------------------------
Public Sub OpenTaskForm(Optional ByVal id As String = "", Optional ByVal defDate As Date = 0, _
                        Optional ByVal defMile As Boolean = False)
    Dim f As frmTask, savedID As String, changed As Boolean
    Set f = New frmTask
    f.Init id, defDate, defMile
    f.Show
    changed = f.Saved
    If changed Then savedID = f.SavedID
    Unload f
    If changed Then AfterTaskChanged savedID
End Sub

Public Sub NewTaskButton()
    Dim d As Date, kind As String
    If ActiveSheet.Name = SH_CAL Then kind = CellKind(ActiveCell.Row, ActiveCell.Column, d)
    If kind = "" Then d = Date
    OpenTaskForm "", d
End Sub

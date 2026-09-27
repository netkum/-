Attribute VB_Name = "modProjectView"
'==============================================================
' ProjectView : 프로젝트 1개 선택 → 요약 + 단계 현황표 + 주 단위 Gantt (단계 → Task/담당자)
'  열: A 숨김(키) | B 항목 | C 담당 | D 진척 | E 계획 | F~ 주(week)
'==============================================================
Option Explicit

Private Const PV_COL1 As Long = 6
Private Const PV_MAXW As Long = 30

Public Sub ShowProject(ByVal pid As String)
    Dim ws As Worksheet, lo As ListObject, r As Long
    Set ws = GetWS(SH_PV)
    Set lo = GetLO(TB_PROJ)
    r = FindRow(lo, PR_ID, pid)
    Application.EnableEvents = False
    If r > 0 Then
        ws.Range(PV_FILTER).Value = pid & " " & lo.DataBodyRange.Cells(r, PR_NAME).Value
    Else
        ws.Range(PV_FILTER).Value = pid
    End If
    Application.EnableEvents = True
    RenderProjectView
    ws.Activate
End Sub

Public Sub RenderProjectView()
    Dim ws As Worksheet, a As Variant, nA As Long, pid As String, pr As ListObject, pRow As Long
    Dim st As Object, v As Variant, ph As Variant, p As Long, s As Variant, r As Long, i As Long
    Dim startW As Date, endW As Date, nW As Long, w As Long, col As Long, pm As String
    Dim grp As Collection, used As Object, k As Variant, sig As String

    If Not IsReady() Then Exit Sub
    Set ws = GetWS(SH_PV)
    pid = FilterPid(ws.Range(PV_FILTER).Value)
    On Error GoTo EH
    SpeedOn
    With ws.Range(ws.Cells(3, 1), ws.Cells(600, PV_COL1 + PV_MAXW))
        .ClearContents
        .ClearFormats
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
    End With
    If pid = "" Then
        ws.Range("B4").Value = "위쪽 프로젝트 목록에서 프로젝트를 선택하세요."
        SpeedOff
        Exit Sub
    End If

    a = ReadTable(GetLO(TB_ALL)): nA = RowCount(a)
    Set st = ProjectStats(a)
    Set pr = GetLO(TB_PROJ)
    pRow = FindRow(pr, PR_ID, pid)
    If pRow > 0 Then
        pm = CStr(Nz(pr.DataBodyRange.Cells(pRow, PR_PM).Value))
        If pm = "" Then pm = CStr(Nz(pr.DataBodyRange.Cells(pRow, PR_AUTOPM).Value))
        col = ColorOf(pr.DataBodyRange.Cells(pRow, PR_COLOR), pid)
    Else
        col = PaletteColor(pid)
    End If

    '--- 요약 -----------------------------------------------
    If st.Exists(pid) Then
        v = st(pid)
        ws.Range("B3").Value = "'PM " & IIf(pm = "", "-", pm) & "   │ 참여 " & v(6) & "   │ 기간 " & Format$(v(7), "m/d") & "~" & _
            Format$(v(8), "m/d") & "   │ 진척 " & Format$(v(0), "0") & "% (계획 " & Format$(v(1), "0") & "%, 차이 " & _
            Format$(v(0) - v(1), "+0;-0;0") & "%p)   │ Task " & v(2) & " (완료 " & v(3) & " · 진행 " & v(4) & " · 지연 " & v(5) & _
            ")   │ " & v(12)
        sig = CStr(v(12))
    Else
        ws.Range("B3").Value = "'PM " & IIf(pm = "", "-", pm) & "   │ 등록된 Task가 없습니다."
    End If
    ws.Range("B3").Font.Bold = True
    If Left$(pm, 1) = SYM_DELAY Then ws.Range("B3").Font.Color = RGB(192, 0, 0)

    '--- 단계 현황표 ----------------------------------------
    r = 5
    ws.Cells(r, 2).Value = "■ 단계 현황"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    HeaderRow ws, r, Array("단계", "Task", "진척", "계획", "계획기간 · 담당 · 상태")
    ph = GetPhases(pid)
    For p = 1 To RowCount(ph)
        s = PhaseStat(a, pid, CStr(ph(p, PH_ID)), ph(p, PH_START), ph(p, PH_END))
        r = r + 1
        ws.Cells(r, 2).Value = ph(p, PH_ORDER) & " " & ph(p, PH_NAME)
        ws.Cells(r, 3).Value = s(2)
        ws.Cells(r, 4).Value = s(0) / 100
        ws.Cells(r, 5).Value = s(1) / 100
        ws.Cells(r, PV_COL1).Value = "'" & IIf(s(4) > 0, Format$(s(4), "m/d") & "~" & Format$(s(5), "m/d"), "-") & _
                                    "   " & s(3) & "   " & PhaseSignal(s)
        If s(6) Then ws.Cells(r, PV_COL1).Font.Color = RGB(192, 0, 0)
    Next p
    s = PhaseStat(a, pid, "", Empty, Empty)
    If s(2) > 0 Then
        r = r + 1
        ws.Cells(r, 2).Value = "(단계 미지정)"
        ws.Cells(r, 3).Value = s(2)
        ws.Cells(r, 4).Value = s(0) / 100
        ws.Cells(r, PV_COL1).Value = "'" & s(3)
    End If
    ws.Range(ws.Cells(7, 4), ws.Cells(r, 5)).NumberFormat = "0%"
    If RowCount(ph) = 0 And s(2) = 0 Then r = r + 1: ws.Cells(r, 2).Value = "단계가 없습니다. (PM이 팀원 파일에서 단계를 만들거나, ProjectMaster의 tblPhase에 직접 입력)"

    '--- Gantt ----------------------------------------------
    If Not st.Exists(pid) Then SpeedOff: Exit Sub
    v = st(pid)
    startW = WeekMonday(v(7)): endW = WeekMonday(v(8))
    For p = 1 To RowCount(ph)
        If IsDateVal(ph(p, PH_START)) Then
            If WeekMonday(ToDate(ph(p, PH_START))) < startW Then startW = WeekMonday(ToDate(ph(p, PH_START)))
        End If
        If IsDateVal(ph(p, PH_END)) Then
            If WeekMonday(ToDate(ph(p, PH_END))) > endW Then endW = WeekMonday(ToDate(ph(p, PH_END)))
        End If
    Next p
    If WeekMonday(Date) < startW Then startW = WeekMonday(Date)
    If WeekMonday(Date) > endW Then endW = WeekMonday(Date)
    nW = (endW - startW) \ 7 + 1
    If nW > PV_MAXW Then startW = WeekMonday(Date) - 7 * 8: nW = PV_MAXW

    r = r + 2
    ws.Cells(r, 2).Value = "■ 일정 (주 단위, 진한 부분 = 진척)"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    For w = 0 To nW - 1
        If w = 0 Or Month(startW + 7 * w) <> Month(startW + 7 * (w - 1)) Then ws.Cells(r, PV_COL1 + w).Value = "'" & Month(startW + 7 * w) & "월"
        ws.Cells(r + 1, PV_COL1 + w).Value = Day(startW + 7 * w)
    Next w
    ws.Range(ws.Cells(r, PV_COL1), ws.Cells(r, PV_COL1 + nW - 1)).Font.Bold = True
    ws.Range(ws.Cells(r + 1, PV_COL1), ws.Cells(r + 1, PV_COL1 + nW - 1)).Font.Size = 7
    ws.Range(ws.Cells(r + 1, PV_COL1), ws.Cells(r + 1, PV_COL1 + nW - 1)).HorizontalAlignment = xlCenter
    HeaderRow ws, r + 1, Array("단계 / 업무", "담당", "진척", "계획")
    r = r + 1
    Dim topG As Long
    topG = r

    Set used = CreateObject("Scripting.Dictionary")
    For p = 1 To RowCount(ph)
        Set grp = TasksOf(a, pid, CStr(ph(p, PH_ID)))
        s = PhaseStat(a, pid, CStr(ph(p, PH_ID)), ph(p, PH_START), ph(p, PH_END))
        r = r + 1
        ws.Cells(r, 2).Value = "▣ " & ph(p, PH_ORDER) & " " & ph(p, PH_NAME)
        ws.Cells(r, 2).Font.Bold = True
        ws.Cells(r, 3).Value = s(3)
        ws.Cells(r, 4).Value = s(0) / 100
        ws.Cells(r, 5).Value = s(1) / 100
        If s(4) > 0 Then GanttBar ws, r, s(4), s(5), s(0), startW, nW, col, s(6), True
        For Each k In grp
            r = r + 1
            TaskGanttRow ws, r, a, CLng(k), startW, nW, col
        Next k
    Next p
    Set grp = TasksOf(a, pid, "")
    If grp.Count > 0 Then
        r = r + 1
        ws.Cells(r, 2).Value = "▣ (단계 미지정)"
        ws.Cells(r, 2).Font.Bold = True
        For Each k In grp
            r = r + 1
            TaskGanttRow ws, r, a, CLng(k), startW, nW, col
        Next k
    End If
    ws.Range(ws.Cells(topG + 1, 4), ws.Cells(r, 5)).NumberFormat = "0%"

    ' 오늘선
    w = (WeekMonday(Date) - startW) \ 7
    If w >= 0 And w < nW Then
        With ws.Range(ws.Cells(topG - 1, PV_COL1 + w), ws.Cells(r, PV_COL1 + w)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(192, 0, 0)
        End With
    End If
    SpeedOff
    Exit Sub
EH:
    SpeedReset
    Msg "ProjectView를 그리는 중 오류: " & Err.Description, vbExclamation
End Sub

Private Function PhaseSignal(ByVal s As Variant) As String
    If s(2) = 0 Then
        PhaseSignal = "Task 없음"
    ElseIf s(0) >= 99.5 Then
        PhaseSignal = SYM_DONE & " 완료"
    ElseIf s(6) Then
        PhaseSignal = SYM_DELAY & " 지연"
    ElseIf s(0) - s(1) < -10 Then
        PhaseSignal = "주의"
    ElseIf s(7) Then
        PhaseSignal = "진행"
    Else
        PhaseSignal = "예정"
    End If
End Function

Private Sub HeaderRow(ByVal ws As Worksheet, ByVal r As Long, ByVal names As Variant)
    Dim n As Long
    n = UBound(names) - LBound(names) + 1
    With ws.Cells(r, 2).Resize(1, n)
        .Value = names
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 84, 106)
    End With
End Sub

' 프로젝트·단계의 Task 행 번호 (시작일 순)
Private Function TasksOf(ByVal a As Variant, ByVal pid As String, ByVal phaseId As String) As Collection
    Dim res As New Collection, idx() As Long, n As Long, i As Long, j As Long, tmp As Long
    ReDim idx(1 To RowCount(a) + 1)
    For i = 1 To RowCount(a)
        If IsWorkTask(a, i) Then
            If CStr(a(i, T_PID)) = pid And CStr(Nz(a(i, T_PHASE))) = phaseId Then n = n + 1: idx(n) = i
        End If
    Next i
    For i = 2 To n
        tmp = idx(i): j = i - 1
        Do While j >= 1
            If ToDate(a(idx(j), T_START)) <= ToDate(a(tmp, T_START)) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = tmp
    Next i
    For i = 1 To n: res.Add idx(i): Next i
    Set TasksOf = res
End Function

Private Sub TaskGanttRow(ByVal ws As Worksheet, ByVal r As Long, ByVal a As Variant, ByVal i As Long, _
                         ByVal startW As Date, ByVal nW As Long, ByVal col As Long)
    Dim lbl As String, late As Boolean, w As Long, c As Long
    late = IsLateRow(a, i)
    lbl = "   "
    If CStr(a(i, T_STATUS)) = ST_DONE Then
        lbl = lbl & SYM_DONE & " "
    ElseIf late Then
        lbl = lbl & SYM_DELAY & " "
    End If
    If CStr(Nz(a(i, T_COMMENT))) <> "" Then lbl = lbl & SYM_COMMENT
    lbl = lbl & a(i, T_NAME) & "  (" & Format$(a(i, T_START), "m/d") & "~" & Format$(a(i, T_DUE), "m/d") & ")"
    ws.Cells(r, 1).Value = a(i, T_ID)
    ws.Cells(r, 2).Value = "'" & lbl
    ws.Cells(r, 3).Value = a(i, T_OWNER)
    ws.Cells(r, 4).Value = NumOr(a(i, T_PCT)) / 100
    ws.Cells(r, 5).Value = NumOr(a(i, A_PLANPCT)) / 100
    If late Then ws.Cells(r, 2).Font.Color = RGB(192, 0, 0)
    c = col
    If CStr(a(i, T_STATUS)) = ST_DONE Then c = RGB(166, 166, 166)
    GanttBar ws, r, ToDate(a(i, T_START)), ToDate(a(i, T_DUE)), NumOr(a(i, T_PCT)), startW, nW, c, late, False
    If UCase$(CStr(Nz(a(i, T_MILE)))) = "Y" Then
        w = (WeekMonday(ToDate(a(i, T_DUE))) - startW) \ 7
        If w >= 0 And w < nW Then ws.Cells(r, PV_COL1 + w).Value = SYM_MILE
    End If
End Sub

Public Sub GanttBar(ByVal ws As Worksheet, ByVal r As Long, ByVal s As Date, ByVal e As Date, ByVal pct As Double, _
                    ByVal startW As Date, ByVal nW As Long, ByVal col As Long, ByVal late As Boolean, ByVal isGroup As Boolean)
    Dim w1 As Long, w2 As Long, w As Long, nDone As Long
    w1 = (WeekMonday(s) - startW) \ 7
    w2 = (WeekMonday(e) - startW) \ 7
    If w2 < 0 Or w1 > nW - 1 Then Exit Sub
    nDone = Int((w2 - w1 + 1) * pct / 100 + 0.5)
    For w = w1 To w2
        If w >= 0 And w < nW Then
            With ws.Cells(r, PV_COL1 + w).Interior
                If w - w1 < nDone Then
                    .Color = IIf(isGroup, col, LightColor(col, 0.2))
                ElseIf late Then
                    .Color = RGB(244, 176, 176)
                Else
                    .Color = LightColor(col, IIf(isGroup, 0.55, 0.72))
                End If
            End With
        End If
    Next w
End Sub

Public Function ColorOf(ByVal cell As Range, ByVal pid As String) As Long
    If cell.Interior.ColorIndex = xlNone Then ColorOf = PaletteColor(pid) Else ColorOf = cell.Interior.Color
End Function

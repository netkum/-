Attribute VB_Name = "modMyProject"
'==============================================================
' MyProject 시트 - 몇 개월 단위 프로젝트 흐름 (주 단위 간트)
'  열: A 숨김(TaskID) | B 항목 | C 진척 | D~ 주(week)
'  행 5부터: 단계 요약 행 → Task 행 / "전체" 선택 시 프로젝트 요약 행
'==============================================================
Option Explicit

Private Const MP_TOP As Long = 5
Private Const MP_COL1 As Long = 4
Private Const MP_MAXW As Long = 30

Public Sub RenderMyProject()
    Dim ws As Worksheet, t As Variant, nT As Long, i As Long, pid As String, pd As Object
    Dim startW As Date, endW As Date, nW As Long, r As Long, w As Long
    Dim ph As Variant, p As Long, grp As Collection, ungrouped As Collection, sel As Collection
    Dim projIds As Collection, seen As Object, k As Variant, nDone As Long, nDoing As Long, nLate As Long

    If Not IsReady() Then Exit Sub
    Set ws = GetWS(SH_MP)
    On Error GoTo EH
    SpeedOn
    ClearMP ws

    pid = FilterPid(ws.Range(MP_FILTER).Value)
    t = LoadTasks(): nT = RowCount(t)
    Set pd = ProjectDict()

    ' 대상 Task
    Set sel = New Collection
    For i = 1 To nT
        If IsValidTask(t, i) Then
            If CStr(t(i, T_STATUS)) <> ST_CANCEL And Not IsSpecialProject(CStr(t(i, T_PID))) Then
                If pid = "" Or CStr(t(i, T_PID)) = pid Then sel.Add i
            End If
        End If
    Next i
    If sel.Count = 0 Then
        ws.Range("B" & MP_TOP).Value = "표시할 Task가 없습니다. (캘린더에서 Task를 등록하세요)"
        SpeedOff
        Exit Sub
    End If

    ' 기간 (최대 30주, 오늘 포함)
    startW = WeekMonday(ToDate(t(sel(1), T_START))): endW = WeekMonday(ToDate(t(sel(1), T_DUE)))
    For Each k In sel
        If WeekMonday(ToDate(t(k, T_START))) < startW Then startW = WeekMonday(ToDate(t(k, T_START)))
        If WeekMonday(ToDate(t(k, T_DUE))) > endW Then endW = WeekMonday(ToDate(t(k, T_DUE)))
    Next k
    If pid <> "" Then
        ph = GetPhases(pid)
        For p = 1 To RowCount(ph)
            If IsDateVal(ph(p, PH_START)) And WeekMonday(ToDate(ph(p, PH_START))) < startW Then startW = WeekMonday(ToDate(ph(p, PH_START)))
            If IsDateVal(ph(p, PH_END)) And WeekMonday(ToDate(ph(p, PH_END))) > endW Then endW = WeekMonday(ToDate(ph(p, PH_END)))
        Next p
    End If
    If WeekMonday(Date) < startW Then startW = WeekMonday(Date)
    If WeekMonday(Date) > endW Then endW = WeekMonday(Date)
    nW = (endW - startW) \ 7 + 1
    If nW > MP_MAXW Then
        ' 너무 길면 오늘 기준 8주 전부터 30주만 표시
        startW = WeekMonday(Date) - 7 * 8
        nW = MP_MAXW
    End If
    DrawHeader ws, startW, nW

    ' 요약 줄
    For Each k In sel
        If CStr(t(k, T_STATUS)) = ST_DONE Then nDone = nDone + 1
        If CStr(t(k, T_STATUS)) = ST_DOING Then nDoing = nDoing + 1
        If IsDelayed(t, k) Then nLate = nLate + 1
    Next k
    ws.Range("B3").Value = "'" & IIf(pid = "", "전체 프로젝트", pid & " " & ProjectNameOf(pid, pd)) & _
        "   진척 " & Format$(WeightedPct(t, sel), "0") & "%   내 Task " & sel.Count & "건 (완료 " & nDone & _
        " · 진행 " & nDoing & " · 지연 " & nLate & ")" & IIf(pid <> "" And IsPM(pid), "   [PM]", "")

    r = MP_TOP
    If pid = "" Then
        ' 프로젝트별 요약 + Task
        Set projIds = New Collection
        Set seen = CreateObject("Scripting.Dictionary")
        For Each k In sel
            If Not seen.Exists(CStr(t(k, T_PID))) Then seen(CStr(t(k, T_PID))) = True: projIds.Add CStr(t(k, T_PID))
        Next k
        For Each k In projIds
            Set grp = New Collection
            For i = 1 To sel.Count
                If CStr(t(sel(i), T_PID)) = k Then grp.Add sel(i)
            Next i
            DrawGroupRow ws, r, "▣ " & ProjectLabel(CStr(k), pd) & " " & ProjectNameOf(CStr(k), pd), t, grp, startW, nW, _
                         ProjectColor(CStr(k), pd), 0, 0
            r = r + 1
            DrawTaskRows ws, r, t, grp, startW, nW, pd
        Next k
    Else
        ' 단계별
        Set ungrouped = New Collection
        For Each k In sel
            ungrouped.Add k
        Next k
        For p = 1 To RowCount(ph)
            Set grp = New Collection
            For i = ungrouped.Count To 1 Step -1
                If CStr(Nz(t(ungrouped(i), T_PHASE))) = CStr(ph(p, PH_ID)) Then grp.Add ungrouped(i): ungrouped.Remove i
            Next i
            Set grp = Reverse(grp)
            DrawGroupRow ws, r, "▣ " & ph(p, PH_ORDER) & " " & ph(p, PH_NAME), t, grp, startW, nW, _
                         ProjectColor(pid, pd), ToDate(ph(p, PH_START)), ToDate(ph(p, PH_END))
            r = r + 1
            DrawTaskRows ws, r, t, grp, startW, nW, pd
        Next p
        If ungrouped.Count > 0 Then
            DrawGroupRow ws, r, "▣ (단계 미지정)", t, ungrouped, startW, nW, RGB(128, 128, 128), 0, 0
            r = r + 1
            DrawTaskRows ws, r, t, ungrouped, startW, nW, pd
        End If
    End If

    ' 오늘 세로선
    w = (WeekMonday(Date) - startW) \ 7
    If w >= 0 And w < nW Then
        With ws.Range(ws.Cells(MP_TOP - 1, MP_COL1 + w), ws.Cells(r, MP_COL1 + w)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(192, 0, 0)
        End With
    End If
    SpeedOff
    Exit Sub
EH:
    SpeedReset
    Msg "MyProject를 그리는 중 오류: " & Err.Description, vbExclamation
End Sub

Private Function Reverse(ByVal c As Collection) As Collection
    Dim r As New Collection, i As Long
    For i = c.Count To 1 Step -1: r.Add c(i): Next i
    Set Reverse = r
End Function

Private Sub ClearMP(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(3, 1), ws.Cells(500, MP_COL1 + MP_MAXW))
        .ClearContents
        .ClearComments
        .Interior.Pattern = xlNone
        .Font.Bold = False
        .Font.Color = RGB(38, 38, 38)
        .Borders.LineStyle = xlNone
    End With
End Sub

Private Sub DrawHeader(ByVal ws As Worksheet, ByVal startW As Date, ByVal nW As Long)
    Dim w As Long, d As Date
    For w = 0 To nW - 1
        d = startW + 7 * w
        If w = 0 Or Month(d) <> Month(d - 7) Then
            ws.Cells(MP_TOP - 2, MP_COL1 + w).Value = "'" & Month(d) & "월"
            ws.Cells(MP_TOP - 2, MP_COL1 + w).Font.Bold = True
        End If
        ws.Cells(MP_TOP - 1, MP_COL1 + w).Value = Day(d)
        ws.Cells(MP_TOP - 1, MP_COL1 + w).HorizontalAlignment = xlCenter
        ws.Cells(MP_TOP - 1, MP_COL1 + w).Font.Size = 7
    Next w
    ws.Range(ws.Cells(MP_TOP - 1, 2), ws.Cells(MP_TOP - 1, MP_COL1 + nW - 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(MP_TOP - 1, 2).Value = "항목"
    ws.Cells(MP_TOP - 1, 3).Value = "진척"
End Sub

' 그룹(단계/프로젝트) 요약 행
Private Sub DrawGroupRow(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal t As Variant, _
                         ByVal grp As Collection, ByVal startW As Date, ByVal nW As Long, ByVal col As Long, _
                         ByVal planS As Date, ByVal planE As Date)
    Dim s As Date, e As Date, k As Variant, pct As Double, late As Boolean
    ws.Cells(r, 2).Value = "'" & label
    ws.Cells(r, 2).Font.Bold = True
    If grp.Count > 0 Then
        s = ToDate(t(grp(1), T_START)): e = ToDate(t(grp(1), T_DUE))
        For Each k In grp
            If ToDate(t(k, T_START)) < s Then s = ToDate(t(k, T_START))
            If ToDate(t(k, T_DUE)) > e Then e = ToDate(t(k, T_DUE))
        Next k
        pct = WeightedPct(t, grp)
    End If
    If planS > 0 Then s = planS
    If planE > 0 Then e = planE
    If s = 0 Or e = 0 Then
        ws.Cells(r, 3).Value = "-"
        Exit Sub
    End If
    ws.Cells(r, 3).Value = Format$(pct, "0") & "%"
    ws.Cells(r, 3).Font.Bold = True
    late = (e < Date And pct < 100 And grp.Count > 0)
    FillBar ws, r, s, e, pct, startW, nW, col, late, True
    If late Then ws.Cells(r, 2).Font.Color = RGB(192, 0, 0)
End Sub

Private Sub DrawTaskRows(ByVal ws As Worksheet, ByRef r As Long, ByVal t As Variant, ByVal grp As Collection, _
                         ByVal startW As Date, ByVal nW As Long, ByVal pd As Object)
    Dim k As Variant, s As Date, e As Date, st As String, lbl As String, w As Long, col As Long
    ' 시작일 순
    Set grp = SortByStart(t, grp)
    For Each k In grp
        s = ToDate(t(k, T_START)): e = ToDate(t(k, T_DUE)): st = CStr(t(k, T_STATUS))
        lbl = "   "
        If st = ST_DONE Then
            lbl = lbl & SYM_DONE & " "
        ElseIf IsDelayed(t, k) Then
            lbl = lbl & SYM_DELAY & " "
        End If
        If CStr(Nz(t(k, T_COMMENT))) <> "" Then lbl = lbl & SYM_COMMENT
        lbl = lbl & t(k, T_NAME) & "  (" & Format$(s, "m/d") & "~" & Format$(e, "m/d") & ")"
        ws.Cells(r, 1).Value = CStr(t(k, T_ID))
        ws.Cells(r, 2).Value = "'" & lbl
        ws.Cells(r, 3).Value = NumOr(t(k, T_PCT)) & "%"
        col = ProjectColor(CStr(t(k, T_PID)), pd)
        If st = ST_DONE Then col = RGB(166, 166, 166)
        FillBar ws, r, s, e, NumOr(t(k, T_PCT)), startW, nW, col, IsDelayed(t, k), False
        If UCase$(CStr(Nz(t(k, T_MILE)))) = "Y" Then
            w = (WeekMonday(e) - startW) \ 7
            If w >= 0 And w < nW Then ws.Cells(r, MP_COL1 + w).Value = SYM_MILE
        End If
        r = r + 1
    Next k
End Sub

Private Function SortByStart(ByVal t As Variant, ByVal grp As Collection) As Collection
    Dim a() As Long, i As Long, j As Long, tmp As Long, res As New Collection
    If grp.Count = 0 Then Set SortByStart = grp: Exit Function
    ReDim a(1 To grp.Count)
    For i = 1 To grp.Count: a(i) = grp(i): Next i
    For i = 2 To grp.Count
        tmp = a(i): j = i - 1
        Do While j >= 1
            If ToDate(t(a(j), T_START)) <= ToDate(t(tmp, T_START)) Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = tmp
    Next i
    For i = 1 To grp.Count: res.Add a(i): Next i
    Set SortByStart = res
End Function

' 주 단위 막대: 진척 부분 진하게, 나머지 연하게, 지연이면 나머지 빨강
Private Sub FillBar(ByVal ws As Worksheet, ByVal r As Long, ByVal s As Date, ByVal e As Date, ByVal pct As Double, _
                    ByVal startW As Date, ByVal nW As Long, ByVal col As Long, ByVal late As Boolean, ByVal isGroup As Boolean)
    Dim w1 As Long, w2 As Long, w As Long, nDone As Long, cnt As Long
    w1 = (WeekMonday(s) - startW) \ 7
    w2 = (WeekMonday(e) - startW) \ 7
    If w2 < 0 Or w1 > nW - 1 Then Exit Sub
    cnt = w2 - w1 + 1
    nDone = Int(cnt * pct / 100 + 0.5)
    For w = w1 To w2
        If w >= 0 And w < nW Then
            With ws.Cells(r, MP_COL1 + w).Interior
                If w - w1 < nDone Then
                    .Color = IIf(isGroup, col, LightColor(col, 0.2))
                ElseIf late Then
                    .Color = RGB(244, 176, 176)
                Else
                    .Color = LightColor(col, IIf(isGroup, 0.55, 0.7))
                End If
            End With
        End If
    Next w
End Sub

' MyProject에서 더블클릭 → 캘린더의 그 주로 이동
Public Sub MPDoubleClick(ByVal target As Range, ByRef cancel As Boolean)
    Dim id As String, v As Variant
    id = CStr(Nz(target.Parent.Cells(target.Row, 1).Value))
    If id = "" Or target.Row < MP_TOP Then Exit Sub
    cancel = True
    v = GetTaskArr(id)
    If IsEmpty(v) Then Exit Sub
    GoToDate ToDate(v(T_START))
    ShowDetail id, CAL_TOP + ((ActiveCell.Row - CAL_TOP) \ CAL_BLOCK) * CAL_BLOCK
End Sub

Public Sub GoCalendarButton()
    GoThisWeek
End Sub

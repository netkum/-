Attribute VB_Name = "modCalendar"
'==============================================================
' 주간형 월간 달력 그리기
'  - 한 주 = 9행 블록 (날짜 / 보고 줄 / 막대 5줄 / 메모 2줄)
'  - 열: A 주차 | B~F 월~금 | G 토·일 | H 주간 요약 | J 상세 패널
'  - CalMap 시트의 같은 주소에 막대의 TaskID를 기록 (클릭 판별용)
'==============================================================
Option Explicit

'--- 현재 표시 월 --------------------------------------------
Public Function CurYear() As Long
    CurYear = CLng(NumOr(CfgGet("CurYear", Year(Date)), Year(Date)))
End Function

Public Function CurMonth() As Long
    CurMonth = CLng(NumOr(CfgGet("CurMonth", Month(Date)), Month(Date)))
End Function

Public Function FirstMonday() As Date
    FirstMonday = WeekMonday(DateSerial(CurYear(), CurMonth(), 1))
End Function

Public Function WeeksInMonth() As Long
    Dim lastDay As Date
    lastDay = DateSerial(CurYear(), CurMonth() + 1, 0)
    WeeksInMonth = Int((lastDay - FirstMonday()) / 7) + 1
End Function

Public Function BlockTop(ByVal b As Long) As Long
    BlockTop = CAL_TOP + b * CAL_BLOCK
End Function

Public Function LastCalRow() As Long
    LastCalRow = CAL_TOP + CAL_WEEKS * CAL_BLOCK - 1
End Function

' 셀 위치 판별: kind = "date" / "report" / "lane" / "memo" / "" , d = 해당 날짜
Public Function CellKind(ByVal r As Long, ByVal c As Long, ByRef d As Date, _
                         Optional ByRef lane As Long, Optional ByRef b As Long) As String
    Dim off As Long, slot As Long
    If r < CAL_TOP Or r > LastCalRow() Then Exit Function
    If c < COL_DAY1 Or c > COL_WKND Then Exit Function
    b = (r - CAL_TOP) \ CAL_BLOCK
    If b >= WeeksInMonth() Then Exit Function
    off = (r - CAL_TOP) Mod CAL_BLOCK
    slot = c - COL_DAY1
    d = FirstMonday() + 7 * b + slot       ' 주말 칸은 토요일 날짜
    If off = OFF_DATE Then
        CellKind = "date"
    ElseIf off = OFF_REPORT Then
        CellKind = "report"
    ElseIf off >= OFF_LANE And off < OFF_MEMO Then
        CellKind = "lane": lane = off - OFF_LANE
    Else
        CellKind = "memo": lane = off - OFF_MEMO
    End If
End Function

Private Function SlotOf(ByVal offsetDays As Long) As Long
    If offsetDays < 0 Then offsetDays = 0
    If offsetDays > 5 Then offsetDays = 5
    SlotOf = offsetDays
End Function

'--- 월 이동 버튼 -------------------------------------------
Public Sub PrevMonth()
    MoveMonth -1
End Sub

Public Sub NextMonth()
    MoveMonth 1
End Sub

Public Sub MoveMonth(ByVal delta As Long)
    Dim d As Date
    d = DateSerial(CurYear(), CurMonth() + delta, 1)
    CfgSet "CurYear", Year(d)
    CfgSet "CurMonth", Month(d)
    RenderCalendar
    GetWS(SH_CAL).Range("A" & CAL_TOP).Select
End Sub

Public Sub GoThisWeek()
    GoToDate Date
End Sub

' 날짜가 있는 달로 이동하고 해당 주 블록 선택
Public Sub GoToDate(ByVal d As Date)
    Dim ws As Worksheet, b As Long
    If Year(d) <> CurYear() Or Month(d) <> CurMonth() Then
        CfgSet "CurYear", Year(d)
        CfgSet "CurMonth", Month(d)
        RenderCalendar
    End If
    Set ws = GetWS(SH_CAL)
    ws.Activate
    b = (WeekMonday(d) - FirstMonday()) \ 7
    If b < 0 Then b = 0
    Application.EnableEvents = False
    ws.Cells(BlockTop(b), COL_DAY1 + SlotOf(d - WeekMonday(d))).Select
    Application.EnableEvents = True
    ActiveWindow.ScrollRow = IIf(BlockTop(b) - 1 > CAL_TOP, BlockTop(b) - 1, CAL_TOP)
End Sub

'==============================================================
' 달력 전체 그리기
'==============================================================
Public Sub RenderCalendar()
    Dim ws As Worksheet, mp As Worksheet
    Dim t As Variant, nT As Long, i As Long, b As Long, nW As Long
    Dim fp As String, pd As Object, hol As Object, memo As Object, ev As Variant
    Dim inc() As Boolean

    If Not IsReady() Then Exit Sub
    Set ws = GetWS(SH_CAL): Set mp = GetWS(SH_MAP)
    On Error GoTo EH
    SpeedOn

    ClearCalendar ws, mp
    ws.Range("A1").Value = CurYear() & "년 " & CurMonth() & "월   " & OwnerName() & " 업무일정"

    t = LoadTasks()
    nT = RowCount(t)
    fp = FilterPid(ws.Range(CAL_FILTER).Value)
    If nT > 0 Then
        ReDim inc(1 To nT)
        For i = 1 To nT
            If IsValidTask(t, i) Then
                If CStr(t(i, T_STATUS)) <> ST_CANCEL Then
                    inc(i) = (fp = "" Or CStr(t(i, T_PID)) = fp Or IsAbsence(CStr(t(i, T_PID))))
                End If
            End If
        Next i
    End If

    Set pd = ProjectDict()
    Set hol = HolidayDict()
    Set memo = MemoDict()
    ev = ReadTable(GetLO(TB_EVENT))

    nW = WeeksInMonth()
    For b = 0 To nW - 1
        DrawWeek ws, mp, b, FirstMonday() + 7 * b, t, nT, inc, pd, hol, memo, ev
    Next b
    For b = nW To CAL_WEEKS - 1
        ws.Range(ws.Cells(BlockTop(b), 1), ws.Cells(BlockTop(b) + CAL_BLOCK - 1, 1)).EntireRow.Hidden = True
    Next b
    ClearDetailPanel
    CfgSet "CalDirty", ""
    SpeedOff
    Exit Sub
EH:
    SpeedReset
    Msg "캘린더를 그리는 중 오류가 발생했습니다." & vbLf & Err.Description, vbExclamation
End Sub

Private Sub ClearCalendar(ByVal ws As Worksheet, ByVal mp As Worksheet)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(CAL_TOP, 1), ws.Cells(LastCalRow(), COL_SUM))
    rng.EntireRow.Hidden = False
    rng.UnMerge
    rng.ClearContents
    rng.ClearComments
    rng.Interior.Pattern = xlNone
    rng.Font.ColorIndex = xlAutomatic
    rng.Font.Bold = False
    rng.Font.Italic = False
    rng.Borders.LineStyle = xlNone
    rng.HorizontalAlignment = xlLeft
    rng.VerticalAlignment = xlCenter
    rng.WrapText = False
    mp.Range(mp.Cells(CAL_TOP, 1), mp.Cells(LastCalRow(), COL_SUM)).ClearContents
End Sub

Private Sub DrawWeek(ByVal ws As Worksheet, ByVal mp As Worksheet, ByVal b As Long, ByVal wkStart As Date, _
                     ByVal t As Variant, ByVal nT As Long, inc() As Boolean, ByVal pd As Object, _
                     ByVal hol As Object, ByVal memo As Object, ByVal ev As Variant)
    Dim r0 As Long, wkEnd As Date, slot As Long, c As Long, d As Date, k As Long
    Dim occ(0 To 4, 0 To 5) As Boolean, ovN(0 To 5) As Long, ovTxt(0 To 5) As String
    Dim idx() As Long, n As Long, i As Long, s As Long, e As Long, ln As Long, ok As Boolean
    Dim txt As String, curM As Long

    r0 = BlockTop(b)
    wkEnd = wkStart + 6
    curM = CurMonth()

    '--- 행 높이 / 틀 ---------------------------------------
    ws.Rows(r0 + OFF_DATE).RowHeight = 15
    ws.Rows(r0 + OFF_REPORT).RowHeight = 13
    For k = OFF_LANE To OFF_MEMO - 1: ws.Rows(r0 + k).RowHeight = 14: Next k
    For k = OFF_MEMO To CAL_BLOCK - 1: ws.Rows(r0 + k).RowHeight = 13: Next k

    With ws.Range(ws.Cells(r0, COL_WEEK), ws.Cells(r0 + CAL_BLOCK - 1, COL_WEEK))
        .Merge
        .Cells(1, 1).Value = WeekLabel(wkStart)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    If Date >= wkStart And Date <= wkEnd Then ws.Cells(r0, COL_WEEK).Interior.Color = RGB(255, 242, 204)

    For slot = 0 To 5
        c = COL_DAY1 + slot
        With ws.Range(ws.Cells(r0, c), ws.Cells(r0 + CAL_BLOCK - 1, c))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Color = RGB(191, 191, 191)
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(191, 191, 191)
        End With
        With ws.Range(ws.Cells(r0 + OFF_MEMO, c), ws.Cells(r0 + OFF_MEMO, c)).Borders(xlEdgeTop)
            .LineStyle = xlDot
            .Color = RGB(191, 191, 191)
        End With
    Next slot
    With ws.Range(ws.Cells(r0, COL_WEEK), ws.Cells(r0, COL_SUM)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(128, 128, 128)
    End With

    '--- 날짜 행 -------------------------------------------
    For slot = 0 To 5
        c = COL_DAY1 + slot
        d = wkStart + slot
        With ws.Cells(r0 + OFF_DATE, c)
            If slot < 5 Then
                txt = CStr(Day(d))
                If hol.Exists(CLng(d)) Then txt = txt & " " & hol(CLng(d))
            Else
                txt = Day(d) & "·" & Day(d + 1)
            End If
            .Value = "'" & txt
            .Font.Bold = True
            .Interior.Color = RGB(248, 248, 248)
            If slot = 5 Or hol.Exists(CLng(d)) Then .Font.Color = RGB(192, 0, 0)
            If Month(d) <> curM Then .Font.Color = RGB(166, 166, 166)
            If (slot < 5 And d = Date) Or (slot = 5 And (Date = d Or Date = d + 1)) Then
                .Interior.Color = RGB(255, 230, 153)
                ws.Range(ws.Cells(r0, c), ws.Cells(r0 + CAL_BLOCK - 1, c)).BorderAround xlContinuous, xlMedium, , RGB(237, 125, 49)
            End If
        End With
    Next slot

    '--- 보고 줄: 팀 공통 보고일정 + 내 마일스톤 --------------
    For slot = 0 To 5
        d = wkStart + slot
        txt = EventText(ev, d)
        If slot = 5 Then txt = JoinText(txt, EventText(ev, d + 1))
        For i = 1 To nT
            If inc(i) Then
                If UCase$(CStr(Nz(t(i, T_MILE)))) = "Y" Then
                    k = ToDate(t(i, T_DUE))
                    If k = d Or (slot = 5 And k = d + 1) Then txt = JoinText(txt, SYM_MILE & t(i, T_NAME))
                End If
            End If
        Next i
        If txt <> "" Then
            With ws.Cells(r0 + OFF_REPORT, COL_DAY1 + slot)
                .Value = "'" & txt
                .Font.Color = RGB(192, 0, 0)
                .Font.Bold = True
                .Font.Size = 8
            End With
        End If
    Next slot

    '--- 이번 주에 걸친 Task 모으기 + 정렬 -------------------
    If nT > 0 Then ReDim idx(1 To nT)
    For i = 1 To nT
        If inc(i) Then
            If ToDate(t(i, T_START)) <= wkEnd And ToDate(t(i, T_DUE)) >= wkStart Then
                n = n + 1: idx(n) = i
            End If
        End If
    Next i
    SortIdx t, idx, n

    '--- 레인 배치 -----------------------------------------
    For k = 1 To n
        i = idx(k)
        s = SlotOf(IIf(ToDate(t(i, T_START)) > wkStart, ToDate(t(i, T_START)) - wkStart, 0))
        e = SlotOf(IIf(ToDate(t(i, T_DUE)) < wkEnd, ToDate(t(i, T_DUE)) - wkStart, 6))
        For ln = 0 To CAL_LANES - 1
            ok = True
            For slot = s To e
                If occ(ln, slot) Then ok = False: Exit For
            Next slot
            If ok Then Exit For
        Next ln
        If ok And ln < CAL_LANES Then
            For slot = s To e: occ(ln, slot) = True: Next slot
            DrawBar ws, mp, r0 + OFF_LANE + ln, COL_DAY1 + s, COL_DAY1 + e, t, i, _
                    ToDate(t(i, T_START)) < wkStart, ToDate(t(i, T_DUE)) > wkEnd, pd
        Else
            For slot = s To e
                ovN(slot) = ovN(slot) + 1
                ovTxt(slot) = ovTxt(slot) & vbLf & "[" & t(i, T_PID) & "] " & t(i, T_NAME)
            Next slot
        End If
    Next k

    ' 넘친 업무: 날짜 칸에 +n건, 메모에 목록
    For slot = 0 To 5
        If ovN(slot) > 0 Then
            With ws.Cells(r0 + OFF_DATE, COL_DAY1 + slot)
                .Value = "'" & .Value & "  +" & ovN(slot) & "건"
                .AddComment Left$("표시 못한 업무" & ovTxt(slot), 250)
                .Comment.Shape.TextFrame.AutoSize = True
            End With
        End If
    Next slot

    '--- 메모 줄 복원 --------------------------------------
    For slot = 0 To 5
        d = wkStart + slot
        For k = 1 To MEMO_LINES
            If memo.Exists(MemoKey(d, k)) Then
                ws.Cells(r0 + OFF_MEMO + k - 1, COL_DAY1 + slot).Value = memo(MemoKey(d, k))
            End If
        Next k
        ws.Range(ws.Cells(r0 + OFF_MEMO, COL_DAY1 + slot), ws.Cells(r0 + CAL_BLOCK - 1, COL_DAY1 + slot)).Font.Size = 8
    Next slot

    '--- 주간 요약 -----------------------------------------
    With ws.Range(ws.Cells(r0, COL_SUM), ws.Cells(r0 + CAL_BLOCK - 1, COL_SUM))
        .Merge
        .Cells(1, 1).Value = "'" & WeekSummary(t, nT, inc, wkStart)
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 8
        .Interior.Color = RGB(250, 250, 250)
        .BorderAround xlContinuous, xlThin, , RGB(191, 191, 191)
    End With
End Sub

' 우선순위(상>중>하) → 시작일 → 기간 긴 순
Private Sub SortIdx(ByVal t As Variant, idx() As Long, ByVal n As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = 2 To n
        tmp = idx(i)
        j = i - 1
        Do While j >= 1
            If Not Before(t, tmp, idx(j)) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = tmp
    Next i
End Sub

Private Function Before(ByVal t As Variant, ByVal a As Long, ByVal b As Long) As Boolean
    Dim pa As Long, pb As Long
    pa = PriRank(t(a, T_PRI)): pb = PriRank(t(b, T_PRI))
    If pa <> pb Then Before = (pa < pb): Exit Function
    If ToDate(t(a, T_START)) <> ToDate(t(b, T_START)) Then
        Before = (ToDate(t(a, T_START)) < ToDate(t(b, T_START))): Exit Function
    End If
    Before = (ToDate(t(a, T_DUE)) - ToDate(t(a, T_START))) > (ToDate(t(b, T_DUE)) - ToDate(t(b, T_START)))
End Function

Private Function PriRank(ByVal v As Variant) As Long
    Select Case CStr(Nz(v))
        Case "상": PriRank = 1
        Case "하": PriRank = 3
        Case Else: PriRank = 2
    End Select
End Function

Private Sub DrawBar(ByVal ws As Worksheet, ByVal mp As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, _
                    ByVal t As Variant, ByVal i As Long, ByVal contBefore As Boolean, ByVal contAfter As Boolean, _
                    ByVal pd As Object)
    Dim rng As Range, pid As String, st As String, col As Long, lbl As String, tip As String

    Set rng = ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
    pid = CStr(t(i, T_PID))
    st = CStr(t(i, T_STATUS))
    col = ProjectColor(pid, pd)

    lbl = BarLabel(t, i, pd)
    If contBefore Then lbl = SYM_PREV & lbl

    rng.Interior.Color = LightColor(col, 0.72)
    rng.Font.Size = 8
    rng.Font.Color = RGB(38, 38, 38)
    If st = ST_DONE Then
        rng.Interior.Color = RGB(231, 230, 230)
        rng.Font.Color = RGB(128, 128, 128)
    ElseIf st = ST_PLAN Then
        rng.Font.Color = RGB(89, 89, 89)
    ElseIf st = ST_HOLD Then
        rng.Interior.Pattern = xlLightUp
        rng.Interior.PatternColor = RGB(166, 166, 166)
    End If
    If IsDelayed(t, i) Then rng.BorderAround xlContinuous, xlThin, , RGB(192, 0, 0)

    ws.Cells(r, c1).Value = "'" & lbl
    If contAfter Then
        If c2 > c1 Then
            ws.Cells(r, c2).Value = SYM_NEXT
            ws.Cells(r, c2).HorizontalAlignment = xlRight
        Else
            ws.Cells(r, c1).Value = "'" & lbl & " " & SYM_NEXT
        End If
    End If

    mp.Range(mp.Cells(r, c1), mp.Cells(r, c2)).Value = CStr(t(i, T_ID))

    ' 마우스 올리면 보이는 요약
    tip = CStr(t(i, T_NAME)) & vbLf & Format$(t(i, T_START), "m/d") & "~" & Format$(t(i, T_DUE), "m/d") & _
          "  " & NumOr(t(i, T_PCT)) & "%  " & st
    If CStr(Nz(t(i, T_NOTE))) <> "" Then tip = tip & vbLf & Left$(CStr(t(i, T_NOTE)), 60)
    If CStr(Nz(t(i, T_COMMENT))) <> "" Then tip = tip & vbLf & SYM_COMMENT & "팀장: " & Left$(CStr(t(i, T_COMMENT)), 60)
    ws.Cells(r, c1).AddComment Left$(tip, 250)
    ws.Cells(r, c1).Comment.Shape.TextFrame.AutoSize = True
End Sub

Public Function BarLabel(ByVal t As Variant, ByVal i As Long, ByVal pd As Object) As String
    Dim s As String, st As String
    st = CStr(t(i, T_STATUS))
    If CStr(Nz(t(i, T_COMMENT))) <> "" Then s = SYM_COMMENT
    If st = ST_DONE Then
        s = s & SYM_DONE
    ElseIf IsDelayed(t, i) Then
        s = s & SYM_DELAY
    End If
    s = s & "[" & ProjectLabel(CStr(t(i, T_PID)), pd) & "] " & t(i, T_NAME)
    If st = ST_DOING Then s = s & " " & NumOr(t(i, T_PCT)) & "%"
    If st = ST_HOLD Then s = s & " (보류)"
    BarLabel = s
End Function

Private Function EventText(ByVal ev As Variant, ByVal d As Date) As String
    Dim i As Long, tg As String, s As String, me_ As String
    me_ = OwnerName()
    For i = 1 To RowCount(ev)
        If IsDateVal(ev(i, 1)) Then
            If ToDate(ev(i, 1)) = d Then
                tg = Trim$(CStr(Nz(ev(i, 4))))
                If tg = "" Or tg = FILTER_ALL Or InStr(tg, me_) > 0 Then
                    s = JoinText(s, SYM_MILE & Trim$(CStr(Nz(ev(i, 2))) & " " & CStr(Nz(ev(i, 3)))))
                End If
            End If
        End If
    Next i
    EventText = s
End Function

Private Function JoinText(ByVal a As String, ByVal b As String) As String
    If a = "" Then
        JoinText = b
    ElseIf b = "" Then
        JoinText = a
    Else
        JoinText = a & " / " & b
    End If
End Function

' 주간 요약 문구
Public Function WeekSummary(ByVal t As Variant, ByVal nT As Long, inc() As Boolean, ByVal wkStart As Date) As String
    Dim wkEnd As Date, i As Long, nDone As Long, nDoing As Long, nLate As Long
    Dim sDone As String, sNext As String, sWarn As String, cDone As Long, cNext As Long, cWarn As Long
    Dim st As String, s As Date, e As Date, pid As String

    wkEnd = wkStart + 6
    For i = 1 To nT
        If inc(i) Then
            pid = CStr(t(i, T_PID))
            If Not IsAbsence(pid) Then
                st = CStr(t(i, T_STATUS))
                s = ToDate(t(i, T_START)): e = ToDate(t(i, T_DUE))
                ' 금주 완료
                If st = ST_DONE And IsDateVal(t(i, T_DONE)) Then
                    If ToDate(t(i, T_DONE)) >= wkStart And ToDate(t(i, T_DONE)) <= wkEnd Then
                        nDone = nDone + 1
                        AddLine sDone, cDone, SYM_DONE & t(i, T_NAME)
                    End If
                End If
                ' 금주 진행
                If st = ST_DOING And s <= wkEnd And e >= wkStart Then nDoing = nDoing + 1
                ' 지연 / 임박 (오늘이 속한 주 이전까지만)
                If wkStart <= Date Then
                    If IsDelayed(t, i) And e <= wkEnd And e >= wkStart - 28 And WeekMonday(Date) = wkStart Then
                        nLate = nLate + 1
                        AddLine sWarn, cWarn, SYM_DELAY & t(i, T_NAME) & " " & DDayText(e)
                    ElseIf IsDelayed(t, i) And e >= wkStart And e <= wkEnd Then
                        nLate = nLate + 1
                        AddLine sWarn, cWarn, SYM_DELAY & t(i, T_NAME) & " " & DDayText(e)
                    ElseIf st <> ST_DONE And e >= Date And e <= wkEnd And NumOr(t(i, T_PCT)) < 80 Then
                        If WorkDays(Date, e) <= 3 Then AddLine sWarn, cWarn, "!" & t(i, T_NAME) & " " & DDayText(e)
                    End If
                End If
                ' 차주 예정
                If st <> ST_DONE And s <= wkEnd + 7 And e >= wkEnd + 1 Then
                    If s > wkEnd Then
                        AddLine sNext, cNext, "·" & t(i, T_NAME) & " " & Format$(s, "m/d") & "~"
                    Else
                        AddLine sNext, cNext, "·" & t(i, T_NAME) & " (계속)"
                    End If
                End If
            End If
        End If
    Next i

    WeekSummary = "완료 " & nDone & " │ 진행 " & nDoing & " │ 지연 " & nLate
    If sDone <> "" Then WeekSummary = WeekSummary & vbLf & "─ 금주 완료 ─" & sDone & MoreText(cDone)
    If sNext <> "" Then WeekSummary = WeekSummary & vbLf & "─ 차주 예정 ─" & sNext & MoreText(cNext)
    If sWarn <> "" Then WeekSummary = WeekSummary & vbLf & "─ 지연·임박 ─" & sWarn & MoreText(cWarn)
End Function

Private Sub AddLine(ByRef s As String, ByRef cnt As Long, ByVal lineText As String)
    cnt = cnt + 1
    If cnt <= 3 Then s = s & vbLf & lineText
End Sub

Private Function MoreText(ByVal cnt As Long) As String
    If cnt > 3 Then MoreText = vbLf & "  외 " & (cnt - 3) & "건"
End Function

'==============================================================
' 메모 줄 (tblMemo: 날짜 / 줄 / 내용)
'==============================================================
Public Function MemoKey(ByVal d As Date, ByVal lineNo As Long) As String
    MemoKey = Format$(d, "yyyymmdd") & "|" & lineNo
End Function

Public Function MemoDict() As Object
    Dim m As Object, a As Variant, i As Long
    Set m = CreateObject("Scripting.Dictionary")
    a = ReadTable(GetLO(TB_MEMO))
    For i = 1 To RowCount(a)
        If IsDateVal(a(i, 1)) Then m(MemoKey(ToDate(a(i, 1)), CLng(NumOr(a(i, 2), 1)))) = a(i, 3)
    Next i
    Set MemoDict = m
End Function

' 메모 칸이 바뀌면 tblMemo에 저장
Public Sub SaveMemoCell(ByVal cell As Range)
    Dim d As Date, lineNo As Long, kind As String, lo As ListObject, a As Variant, i As Long
    kind = CellKind(cell.Row, cell.Column, d, lineNo)
    If kind <> "memo" Then Exit Sub
    lineNo = lineNo + 1
    Set lo = GetLO(TB_MEMO)
    a = ReadTable(lo)
    For i = 1 To RowCount(a)
        If IsDateVal(a(i, 1)) Then
            If ToDate(a(i, 1)) = d And CLng(NumOr(a(i, 2))) = lineNo Then
                If CStr(Nz(cell.Value)) = "" Then
                    lo.ListRows(i).Delete
                Else
                    lo.DataBodyRange.Cells(i, 3).Value = cell.Value
                End If
                Exit Sub
            End If
        End If
    Next i
    If CStr(Nz(cell.Value)) <> "" Then AppendRow lo, Array(d, lineNo, cell.Value)
End Sub

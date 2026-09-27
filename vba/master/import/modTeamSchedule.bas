Attribute VB_Name = "modTeamSchedule"
'==============================================================
' TeamSchedule : 팀원(행) × 날짜(열) 월간 일정
'  - 칸 값 = 그날 진행 중인 프로젝트 코드 (여러 개면 쉼표)
'  - 휴가·출장·교육 = 회색 칸
'  - 보고일정 행(◆), 맨 아래 "부재 인원" 행
'==============================================================
Option Explicit

Private Const TS_COL1 As Long = 3     ' C열 = 1일
Private Const TS_TOP As Long = 7      ' 첫 팀원 행

Public Sub TSPrevMonth()
    TSMove -1
End Sub

Public Sub TSNextMonth()
    TSMove 1
End Sub

Private Sub TSMove(ByVal delta As Long)
    Dim ws As Worksheet, d As Date
    Set ws = GetWS(SH_TS)
    d = TSMonth()
    Application.EnableEvents = False
    ws.Range(TS_MONTH).Value = DateSerial(Year(d), Month(d) + delta, 1)
    Application.EnableEvents = True
    RenderTeamSchedule
End Sub

Private Function TSMonth() As Date
    Dim v As Variant
    v = GetWS(SH_TS).Range(TS_MONTH).Value
    If IsDateVal(v) Then
        TSMonth = DateSerial(Year(ToDate(v)), Month(ToDate(v)), 1)
    Else
        TSMonth = DateSerial(Year(Date), Month(Date), 1)
    End If
End Function

Public Sub RenderTeamSchedule()
    Dim ws As Worksheet, a As Variant, nA As Long, m0 As Date, nD As Long, dd As Long, d As Date
    Dim members As Collection, seen As Object, mem As Variant, i As Long, r As Long, c As Long
    Dim txt As String, tip As String, absent As Boolean, absCnt() As Long, hol As Object, ev As Collection
    Dim pc As Object, firstPid As String, e As Variant, evTxt As String

    If Not IsReady() Then Exit Sub
    Set ws = GetWS(SH_TS)
    On Error GoTo EH
    SpeedOn
    m0 = TSMonth()
    ws.Range(TS_MONTH).Value = m0
    ws.Range(TS_MONTH).NumberFormat = "yyyy-mm"
    nD = Day(DateSerial(Year(m0), Month(m0) + 1, 0))

    With ws.Range(ws.Cells(4, 1), ws.Cells(300, TS_COL1 + 31))
        .ClearContents
        .ClearComments
        .ClearFormats
        .Font.Name = "맑은 고딕"
        .Font.Size = 8
        .VerticalAlignment = xlCenter
    End With

    a = ReadTable(GetLO(TB_ALL)): nA = RowCount(a)
    Set hol = HolidayDict()
    Set ev = ExpandEvents()
    Set pc = ProjectColors()

    ' 팀원 목록: tblMember + TaskAll 담당자
    Set members = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    mem = ReadTable(GetLO(TB_MEMBER))
    For i = 1 To RowCount(mem)
        If CStr(Nz(mem(i, M_NAME))) <> "" And Not seen.Exists(CStr(mem(i, M_NAME))) Then
            seen(CStr(mem(i, M_NAME))) = True: members.Add CStr(mem(i, M_NAME))
        End If
    Next i
    For i = 1 To nA
        If CStr(Nz(a(i, T_OWNER))) <> "" And Not seen.Exists(CStr(a(i, T_OWNER))) Then
            seen(CStr(a(i, T_OWNER))) = True: members.Add CStr(a(i, T_OWNER))
        End If
    Next i

    ' 머리글
    ws.Cells(4, 2).Value = "날짜"
    ws.Cells(5, 2).Value = "요일"
    ws.Cells(6, 2).Value = "보고일정"
    For dd = 1 To nD
        d = DateSerial(Year(m0), Month(m0), dd)
        c = TS_COL1 + dd - 1
        ws.Cells(4, c).Value = dd
        ws.Cells(5, c).Value = Format$(d, "aaa")
        If Weekday(d, vbMonday) >= 6 Or hol.Exists(CLng(d)) Then
            ws.Range(ws.Cells(4, c), ws.Cells(5, c)).Font.Color = RGB(192, 0, 0)
            ws.Range(ws.Cells(6, c), ws.Cells(TS_TOP + members.Count, c)).Interior.Color = RGB(242, 242, 242)
            If hol.Exists(CLng(d)) Then ws.Cells(4, c).AddComment hol(CLng(d))
        End If
        If d = Date Then ws.Range(ws.Cells(4, c), ws.Cells(5, c)).Interior.Color = RGB(255, 230, 153)
        evTxt = ""
        For Each e In ev
            If e(0) = d Then evTxt = evTxt & IIf(evTxt = "", "", vbLf) & Trim$(e(1) & " " & e(2)) & IIf(e(3) <> FILTER_ALL, " (" & e(3) & ")", "")
        Next e
        If evTxt <> "" Then
            ws.Cells(6, c).Value = SYM_MILE
            ws.Cells(6, c).Font.Color = RGB(192, 0, 0)
            ws.Cells(6, c).AddComment Left$(evTxt, 250)
            ws.Cells(6, c).Comment.Shape.TextFrame.AutoSize = True
        End If
    Next dd
    With ws.Range(ws.Cells(4, 2), ws.Cells(6, TS_COL1 + nD - 1))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' 팀원 행
    ReDim absCnt(1 To nD)
    r = TS_TOP - 1
    For Each mem In members
        r = r + 1
        ws.Cells(r, 2).Value = mem
        ws.Cells(r, 2).Font.Bold = True
        For dd = 1 To nD
            d = DateSerial(Year(m0), Month(m0), dd)
            c = TS_COL1 + dd - 1
            txt = "": tip = "": absent = False: firstPid = ""
            For i = 1 To nA
                If CStr(Nz(a(i, T_OWNER))) = mem And CStr(Nz(a(i, T_STATUS))) <> ST_CANCEL Then
                    If IsDateVal(a(i, T_START)) And IsDateVal(a(i, T_DUE)) Then
                        If ToDate(a(i, T_START)) <= d And ToDate(a(i, T_DUE)) >= d Then
                            If IsAbsence(CStr(a(i, T_PID))) Then
                                absent = True
                                tip = tip & vbLf & a(i, T_PID) & " " & a(i, T_NAME)
                            Else
                                If InStr("," & txt & ",", "," & a(i, T_PID) & ",") = 0 Then txt = txt & IIf(txt = "", "", ",") & a(i, T_PID)
                                If firstPid = "" Then firstPid = CStr(a(i, T_PID))
                                tip = tip & vbLf & "[" & a(i, T_PID) & "] " & a(i, T_NAME) & " " & NumOr(a(i, T_PCT)) & "%"
                            End If
                        End If
                    End If
                End If
            Next i
            If absent Then
                ws.Cells(r, c).Value = PJ_LEAVE
                If InStr(tip, PJ_TRIP) > 0 Then ws.Cells(r, c).Value = PJ_TRIP
                If InStr(tip, PJ_EDU) > 0 Then ws.Cells(r, c).Value = PJ_EDU
                ws.Cells(r, c).Interior.Color = RGB(191, 191, 191)
                If Weekday(d, vbMonday) < 6 Then absCnt(dd) = absCnt(dd) + 1
            ElseIf txt <> "" Then
                ws.Cells(r, c).Value = "'" & txt
                If pc.Exists(firstPid) Then ws.Cells(r, c).Interior.Color = LightColor(pc(firstPid), 0.7)
            End If
            If tip <> "" Then
                ws.Cells(r, c).AddComment Left$(mem & " " & Format$(d, "m/d") & tip, 250)
                ws.Cells(r, c).Comment.Shape.TextFrame.AutoSize = True
            End If
        Next dd
    Next mem

    ' 부재 인원
    r = r + 1
    ws.Cells(r, 2).Value = "부재 인원"
    ws.Cells(r, 2).Font.Bold = True
    For dd = 1 To nD
        If absCnt(dd) > 0 Then
            ws.Cells(r, TS_COL1 + dd - 1).Value = absCnt(dd)
            ws.Cells(r, TS_COL1 + dd - 1).Font.Color = RGB(192, 0, 0)
        End If
    Next dd
    With ws.Range(ws.Cells(4, 2), ws.Cells(r, TS_COL1 + nD - 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
    End With
    ws.Range(ws.Cells(TS_TOP, TS_COL1), ws.Cells(r, TS_COL1 + nD - 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(TS_TOP, TS_COL1), ws.Cells(r, TS_COL1 + nD - 1)).ShrinkToFit = True
    SpeedOff
    Exit Sub
EH:
    SpeedReset
    Msg "TeamSchedule을 그리는 중 오류: " & Err.Description, vbExclamation
End Sub

' 프로젝트 → 표시색 (ProjectMaster Color 칸의 채우기 색, 없으면 기본색)
Public Function ProjectColors() As Object
    Dim d As Object, lo As ListObject, i As Long, pid As String
    Set d = CreateObject("Scripting.Dictionary")
    Set lo = GetLO(TB_PROJ)
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            pid = CStr(Nz(lo.DataBodyRange.Cells(i, PR_ID).Value))
            If pid <> "" Then d(pid) = ColorOf(lo.DataBodyRange.Cells(i, PR_COLOR), pid)
        Next i
    End If
    Set ProjectColors = d
End Function

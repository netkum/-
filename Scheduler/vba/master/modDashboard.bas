Attribute VB_Name = "modDashboard"
'==============================================================
' Dashboard (첫 화면)
'  KPI → 팀 전체 진척 → 프로젝트별 진척 → 담당자별 업무량 → 지연 업무 → 이번 주 마감 → 승인대기 요청
'  O열(숨김): "P:프로젝트ID" / "T:TaskID" - 더블클릭·코멘트 입력 판별용
'  지연 업무의 M열(노란 칸)에 팀장 코멘트를 쓰면 팀원 파일로 전달됨
'==============================================================
Option Explicit

Public Const DB_KEYCOL As Long = 15     ' O
Public Const DB_CMTCOL As Long = 13     ' M

Public Sub RenderDashboard()
    Dim ws As Worksheet, a As Variant, nA As Long, st As Object, pr As Variant, i As Long, r As Long
    Dim pct As Double, plan As Double, v As Variant, pid As String, pm As String, k As Long
    Dim nTot As Long, nDoing As Long, nDone As Long, nLate As Long, nWeek As Long, nStale As Long, nReq As Long, nPM As Long, nFail As Long
    Dim wkS As Date, wkE As Date, staleDays As Long, mem As Variant, startRow As Long

    If Not IsReady() Then Exit Sub
    Set ws = GetWS(SH_DASH)
    On Error GoTo EH
    SpeedOn
    ClearArea ws

    a = ReadTable(GetLO(TB_ALL)): nA = RowCount(a)
    wkS = WeekMonday(Date): wkE = wkS + 6
    staleDays = CLng(NumOr(CfgGet("StaleDays", 7), 7))

    '--- KPI ------------------------------------------------
    Dim staleMem As Object
    Set staleMem = CreateObject("Scripting.Dictionary")
    For i = 1 To nA
        If IsWorkTask(a, i) Then
            nTot = nTot + 1
            Select Case CStr(a(i, T_STATUS))
                Case ST_DOING: nDoing = nDoing + 1
                Case ST_DONE: nDone = nDone + 1
            End Select
            If IsLateRow(a, i) Then nLate = nLate + 1
            If CStr(a(i, T_STATUS)) <> ST_DONE And ToDate(a(i, T_DUE)) >= wkS And ToDate(a(i, T_DUE)) <= wkE Then nWeek = nWeek + 1
            If CStr(a(i, T_STATUS)) = ST_DOING And NumOr(a(i, A_STALE)) >= staleDays Then staleMem(CStr(a(i, T_OWNER))) = True
        End If
    Next i
    nStale = staleMem.Count
    pr = ReadTable(GetLO(TB_REQ))
    For i = 1 To RowCount(pr)
        If CStr(Nz(pr(i, RQ_CODE))) <> "" And Not IsDateVal(pr(i, RQ_DONE)) Then nReq = nReq + 1
    Next i
    nPM = CountPMNeeded()
    mem = ReadTable(GetLO(TB_MEMBER))
    For i = 1 To RowCount(mem)
        If Left$(CStr(Nz(mem(i, M_RESULT))), 2) = "실패" Then nFail = nFail + 1
    Next i

    ws.Range("B2").Value = "마지막 동기화: " & IIf(IsDateVal(CfgGet("LastSync", "")), Format$(CfgGet("LastSync", ""), "yyyy-mm-dd hh:mm"), "-") & _
                           IIf(nFail > 0, "    " & SYM_DELAY & " 읽기 실패 " & nFail & "명 (Setting의 tblMember 확인)", "")
    If nFail > 0 Then ws.Range("B2").Font.Color = RGB(192, 0, 0)
    Kpi ws, 2, "전체 Task", nTot, 0
    Kpi ws, 3, "진행", nDoing, 0
    Kpi ws, 4, "완료", nDone, 0
    Kpi ws, 5, "지연", nLate, IIf(nLate > 0, RGB(192, 0, 0), 0)
    Kpi ws, 6, "금주 마감", nWeek, IIf(nWeek > 0, RGB(191, 144, 0), 0)
    Kpi ws, 7, "미갱신 팀원", nStale, IIf(nStale > 0, RGB(191, 144, 0), 0)
    Kpi ws, 8, "프로젝트 요청", nReq, IIf(nReq > 0, RGB(47, 117, 181), 0)
    Kpi ws, 9, "PM 지정 필요", nPM, IIf(nPM > 0, RGB(192, 0, 0), 0)

    '--- 팀 전체 진척 ----------------------------------------
    OverallPct a, pct, plan
    ws.Range("B7").Value = "팀 전체 진척률"
    ws.Range("B7").Font.Bold = True
    ws.Range("C7").Value = pct / 100
    ws.Range("C7").NumberFormat = "0%"
    ws.Range("C7").Font.Size = 16
    ws.Range("C7").Font.Bold = True
    ws.Range("D7").Value = "'" & TextBar(pct, 20)
    ws.Range("D7").Font.Color = RGB(47, 117, 181)
    ws.Range("G7").Value = "계획 " & Format$(plan, "0") & "%  (차이 " & Format$(pct - plan, "+0;-0;0") & "%p)"
    If pct - plan < -10 Then ws.Range("G7").Font.Color = RGB(192, 0, 0)

    '--- 프로젝트별 진척 -------------------------------------
    r = 9
    Section ws, r, "프로젝트별 진척  (프로젝트 행 더블클릭 → ProjectView)"
    r = r + 1
    Header ws, r, Array("프로젝트", "PM", "기간", "현재 단계", "단계", "Task(완료/전체)", "계획", "실적", "차이", "상태", "지연", "참여자")
    startRow = r + 1
    Set st = ProjectStats(a)
    pr = ReadTable(GetLO(TB_PROJ))
    For i = 1 To RowCount(pr)
        pid = CStr(Nz(pr(i, PR_ID)))
        If pid <> "" And CStr(Nz(pr(i, PR_STATUS))) <> "완료" And CStr(Nz(pr(i, PR_STATUS))) <> "종료" Then
            r = r + 1
            pm = CStr(Nz(pr(i, PR_PM)))
            If pm = "" Then pm = CStr(Nz(pr(i, PR_AUTOPM)))
            ws.Cells(r, 2).Value = pid & " " & pr(i, PR_NAME)
            ws.Cells(r, 3).Value = pm
            If Left$(pm, 1) = SYM_DELAY Then ws.Cells(r, 3).Font.Color = RGB(192, 0, 0)
            ws.Cells(r, DB_KEYCOL).Value = "P:" & pid
            If st.Exists(pid) Then
                v = st(pid)
                ws.Cells(r, 4).Value = "'" & Format$(IIf(IsDateVal(pr(i, PR_START)), pr(i, PR_START), v(7)), "m/d") & "~" & _
                                       Format$(IIf(IsDateVal(pr(i, PR_END)), pr(i, PR_END), v(8)), "m/d")
                ws.Cells(r, 5).Value = v(9)
                ws.Cells(r, 6).Value = "'" & v(10)
                ws.Cells(r, 7).Value = "'" & v(3) & " / " & v(2)
                ws.Cells(r, 8).Value = v(1) / 100
                ws.Cells(r, 9).Value = v(0) / 100
                ws.Cells(r, 10).Value = (v(0) - v(1)) / 100
                ws.Cells(r, 11).Value = v(12)
                ws.Cells(r, 12).Value = v(5)
                ws.Cells(r, 13).Value = v(6)
                SignalColor ws.Cells(r, 11), CStr(v(12))
                If v(5) > 0 Then ws.Cells(r, 12).Font.Color = RGB(192, 0, 0)
            Else
                ws.Cells(r, 5).Value = "(Task 없음)"
            End If
        End If
    Next i
    ' 승인 전 임시 프로젝트 묶음
    Dim key As Variant, vv As Variant
    For Each key In st.Keys
        If Left$(CStr(key), 2) = "N-" Then
            vv = st(key)
            r = r + 1
            ws.Cells(r, 2).Value = CStr(key) & " (미승인 프로젝트)"
            ws.Cells(r, 2).Font.Color = RGB(47, 117, 181)
            ws.Cells(r, 7).Value = "'" & vv(3) & " / " & vv(2)
            ws.Cells(r, 9).Value = vv(0) / 100
            ws.Cells(r, 13).Value = vv(6)
        End If
    Next key
    If r >= startRow Then
        ws.Range(ws.Cells(startRow, 8), ws.Cells(r, 10)).NumberFormat = "0%"
        AddBar ws.Range(ws.Cells(startRow, 9), ws.Cells(r, 9))
        TableBorder ws, startRow - 1, r, 13
    End If

    '--- 담당자별 업무량 -------------------------------------
    r = r + 2
    Section ws, r, "담당자별 업무량"
    r = r + 1
    Header ws, r, Array("이름", "진행", "예정", "지연", "금주", "합계(미완료)", "최근 갱신", "파일 상태")
    startRow = r + 1
    For k = 1 To RowCount(mem)
        r = r + 1
        MemberRow ws, r, CStr(mem(k, M_NAME)), a, wkS, wkE, staleDays
        ws.Cells(r, 9).Value = mem(k, M_RESULT)
        If Left$(CStr(Nz(mem(k, M_RESULT))), 2) = "실패" Then ws.Cells(r, 9).Font.Color = RGB(192, 0, 0)
    Next k
    If r >= startRow Then
        With ws.Range(ws.Cells(startRow, 7), ws.Cells(r, 7)).FormatConditions.AddColorScale(ColorScaleType:=3)
            .ColorScaleCriteria(1).FormatColor.Color = RGB(198, 239, 206)
            .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 156)
            .ColorScaleCriteria(3).FormatColor.Color = RGB(255, 199, 206)
        End With
        TableBorder ws, startRow - 1, r, 9
    End If

    '--- 지연 업무 -------------------------------------------
    r = r + 2
    Section ws, r, "지연 업무  (M열 노란 칸에 팀장 코멘트 입력 → 팀원 파일에 표시)"
    r = r + 1
    Header ws, r, Array("프로젝트", "담당", "업무", "종료일", "D-day", "진척", "메모(지연 사유)", "", "", "", "", "팀장 코멘트")
    startRow = r + 1
    For i = 1 To nA
        If IsLateRow(a, i) Then
            r = r + 1
            TaskRow ws, r, a, i
            ws.Cells(r, 8).Value = a(i, T_NOTE)
            ws.Cells(r, DB_CMTCOL).Value = a(i, T_COMMENT)
            ws.Cells(r, DB_CMTCOL).Interior.Color = RGB(255, 242, 204)
            ws.Cells(r, 6).Font.Color = RGB(192, 0, 0)
        End If
    Next i
    If r < startRow Then r = r + 1: ws.Cells(r, 2).Value = "지연 업무가 없습니다."
    If r >= startRow Then TableBorder ws, startRow - 1, r, 13

    '--- 이번 주 마감 ----------------------------------------
    r = r + 2
    Section ws, r, "이번 주 마감  (" & Format$(wkS, "m/d") & "~" & Format$(wkE, "m/d") & ")"
    r = r + 1
    Header ws, r, Array("프로젝트", "담당", "업무", "종료일", "D-day", "진척", "상태")
    startRow = r + 1
    For i = 1 To nA
        If IsWorkTask(a, i) Then
            If ToDate(a(i, T_DUE)) >= wkS And ToDate(a(i, T_DUE)) <= wkE Then
                r = r + 1
                TaskRow ws, r, a, i
                ws.Cells(r, 8).Value = a(i, T_STATUS)
            End If
        End If
    Next i
    If r < startRow Then r = r + 1: ws.Cells(r, 2).Value = "이번 주 마감 업무가 없습니다."
    If r >= startRow Then TableBorder ws, startRow - 1, r, 8

    '--- 승인대기 요청 ---------------------------------------
    If nReq > 0 Then
        r = r + 2
        Section ws, r, "승인대기 프로젝트 요청  (ProjectMaster 시트에서 처리 방법 선택 → [요청 처리 적용])"
        r = r + 1
        Header ws, r, Array("요청코드", "요청자", "프로젝트명", "예상 기간", "Task 수", "요청일", "메모")
        startRow = r + 1
        For i = 1 To RowCount(pr)
            If CStr(Nz(pr(i, RQ_CODE))) <> "" And Not IsDateVal(pr(i, RQ_DONE)) Then
                r = r + 1
                ws.Cells(r, 2).Value = pr(i, RQ_CODE)
                ws.Cells(r, 3).Value = pr(i, RQ_OWNER)
                ws.Cells(r, 4).Value = pr(i, RQ_NAME)
                ws.Cells(r, 5).Value = "'" & DateText(pr(i, RQ_START)) & "~" & DateText(pr(i, RQ_END))
                ws.Cells(r, 6).Value = pr(i, RQ_TASKS)
                ws.Cells(r, 7).Value = "'" & DateText(pr(i, RQ_DATE))
                ws.Cells(r, 8).Value = pr(i, RQ_MEMO)
            End If
        Next i
        TableBorder ws, startRow - 1, r, 8
    End If
    SpeedOff
    Exit Sub
EH:
    SpeedReset
    Msg "Dashboard를 그리는 중 오류: " & Err.Description, vbExclamation
End Sub

Private Sub ClearArea(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(2, 2), ws.Cells(2000, DB_KEYCOL))
        .ClearContents
        .ClearFormats
        .FormatConditions.Delete
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub Kpi(ByVal ws As Worksheet, ByVal c As Long, ByVal label As String, ByVal value As Long, ByVal col As Long)
    With ws.Cells(4, c)
        .Value = label
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(89, 89, 89)
        .Interior.Color = RGB(242, 242, 242)
    End With
    With ws.Cells(5, c)
        .Value = value
        .HorizontalAlignment = xlCenter
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = IIf(col = 0, RGB(38, 38, 38), col)
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Range(ws.Cells(4, c), ws.Cells(5, c)).BorderAround xlContinuous, xlThin, , RGB(217, 217, 217)
End Sub

Private Sub Section(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String)
    ws.Cells(r, 2).Value = "■ " & title
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Size = 11
End Sub

Private Sub Header(ByVal ws As Worksheet, ByVal r As Long, ByVal names As Variant)
    Dim n As Long
    n = UBound(names) - LBound(names) + 1
    With ws.Cells(r, 2).Resize(1, n)
        .Value = names
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 84, 106)
    End With
End Sub

Private Sub TableBorder(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal lastCol As Long)
    With ws.Range(ws.Cells(r1, 2), ws.Cells(r2, lastCol)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
    End With
End Sub

Private Sub AddBar(ByVal rng As Range)
    Dim db As Databar
    Set db = rng.FormatConditions.AddDatabar
    db.MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    db.MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
    db.BarColor.Color = RGB(91, 155, 213)
End Sub

Public Sub SignalColor(ByVal cell As Range, ByVal sig As String)
    Select Case sig
        Case "완료": cell.Interior.Color = RGB(217, 217, 217)
        Case "정상": cell.Interior.Color = RGB(198, 239, 206)
        Case "주의": cell.Interior.Color = RGB(255, 235, 156)
        Case "지연": cell.Interior.Color = RGB(255, 199, 206)
    End Select
    cell.HorizontalAlignment = xlCenter
End Sub

Public Function TextBar(ByVal pct As Double, ByVal width As Long) As String
    Dim n As Long
    n = Int(pct / 100 * width + 0.5)
    If n < 0 Then n = 0
    If n > width Then n = width
    TextBar = String$(n, "■") & String$(width - n, "□")
End Function

Private Function DateText(ByVal v As Variant) As String
    If IsDateVal(v) Then DateText = Format$(ToDate(v), "m/d")
End Function

Private Sub TaskRow(ByVal ws As Worksheet, ByVal r As Long, ByVal a As Variant, ByVal i As Long)
    ws.Cells(r, 2).Value = a(i, T_PID) & " " & a(i, T_PNAME)
    ws.Cells(r, 3).Value = a(i, T_OWNER)
    ws.Cells(r, 4).Value = a(i, T_NAME)
    ws.Cells(r, 5).Value = ToDate(a(i, T_DUE))
    ws.Cells(r, 5).NumberFormat = "m/d(aaa)"
    ws.Cells(r, 6).Value = "'" & DDayText(ToDate(a(i, T_DUE)))
    ws.Cells(r, 7).Value = NumOr(a(i, T_PCT)) / 100
    ws.Cells(r, 7).NumberFormat = "0%"
    ws.Cells(r, DB_KEYCOL).Value = "T:" & a(i, T_ID)
End Sub

Private Sub MemberRow(ByVal ws As Worksheet, ByVal r As Long, ByVal nm As String, ByVal a As Variant, _
                      ByVal wkS As Date, ByVal wkE As Date, ByVal staleDays As Long)
    Dim i As Long, nDo As Long, nPl As Long, nLt As Long, nWk As Long, nAll As Long, lastU As Date, st As String
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, T_OWNER))) = nm And IsWorkTask(a, i) Then
            st = CStr(a(i, T_STATUS))
            If st = ST_DOING Then nDo = nDo + 1
            If st = ST_PLAN Then nPl = nPl + 1
            If IsLateRow(a, i) Then nLt = nLt + 1
            If st <> ST_DONE Then
                nAll = nAll + 1
                If ToDate(a(i, T_START)) <= wkE And ToDate(a(i, T_DUE)) >= wkS Then nWk = nWk + 1
            End If
            If IsDateVal(a(i, T_UPDATED)) Then
                If ToDate(a(i, T_UPDATED)) > lastU Then lastU = ToDate(a(i, T_UPDATED))
            End If
        End If
    Next i
    ws.Cells(r, 2).Value = nm
    ws.Cells(r, 3).Value = nDo
    ws.Cells(r, 4).Value = nPl
    ws.Cells(r, 5).Value = nLt
    ws.Cells(r, 6).Value = nWk
    ws.Cells(r, 7).Value = nAll
    If lastU > 0 Then
        ws.Cells(r, 8).Value = lastU
        ws.Cells(r, 8).NumberFormat = "m/d"
        If Date - lastU >= staleDays Then ws.Cells(r, 8).Font.Color = RGB(192, 0, 0)
    End If
    If nLt > 0 Then ws.Cells(r, 5).Font.Color = RGB(192, 0, 0)
End Sub

' Dashboard 더블클릭 → 프로젝트 보기
Public Sub DashDoubleClick(ByVal target As Range, ByRef cancel As Boolean)
    Dim key As String, pid As String, r As Long
    key = CStr(Nz(target.Parent.Cells(target.Row, DB_KEYCOL).Value))
    If key = "" Then Exit Sub
    If target.Column = DB_CMTCOL And Left$(key, 2) = "T:" Then Exit Sub   ' 코멘트 칸은 편집
    cancel = True
    If Left$(key, 2) = "P:" Then
        pid = Mid$(key, 3)
    Else
        r = FindRow(GetLO(TB_ALL), T_ID, Mid$(key, 3))
        If r > 0 Then pid = CStr(GetLO(TB_ALL).DataBodyRange.Cells(r, T_PID).Value)
    End If
    If pid <> "" Then ShowProject pid
End Sub

' Dashboard 코멘트 칸 변경
Public Sub DashChange(ByVal target As Range)
    Dim c As Range, key As String
    For Each c In target.Cells
        If c.Column = DB_CMTCOL Then
            key = CStr(Nz(c.Parent.Cells(c.Row, DB_KEYCOL).Value))
            If Left$(key, 2) = "T:" Then UpsertComment Mid$(key, 3), CStr(Nz(c.Value))
        End If
    Next c
End Sub

Attribute VB_Name = "modCalc"
'==============================================================
' 진척 / 지연 계산
'  Task 가중치 = PlanDays (공휴일 제외 작업일수)
'  프로젝트·단계·전체 진척 = Σ(가중치 × 진척) ÷ Σ(가중치)
'  계획진척 = 경과 작업일 ÷ 전체 작업일 (시작 전 0, 종료 후 100)
'==============================================================
Option Explicit

' 진척 계산 대상 (휴가·출장·교육·공통·취소 제외)
Public Function IsWorkTask(ByVal a As Variant, ByVal i As Long) As Boolean
    Dim pid As String
    pid = CStr(Nz(a(i, T_PID)))
    If pid = "" Or IsSpecialProject(pid) Then Exit Function
    If CStr(Nz(a(i, T_STATUS))) = ST_CANCEL Then Exit Function
    If Not IsDateVal(a(i, T_START)) Or Not IsDateVal(a(i, T_DUE)) Then Exit Function
    IsWorkTask = True
End Function

Public Function IsLateRow(ByVal a As Variant, ByVal i As Long) As Boolean
    Dim st As String
    If Not IsWorkTask(a, i) Then Exit Function
    st = CStr(Nz(a(i, T_STATUS)))
    If st = ST_DONE Then Exit Function
    IsLateRow = (ToDate(a(i, T_DUE)) < Date And NumOr(a(i, T_PCT)) < 100)
End Function

Public Function PlannedPct(ByVal s As Date, ByVal e As Date) As Double
    Dim tot As Long, done As Long
    If Date < s Then PlannedPct = 0: Exit Function
    If Date >= e Then PlannedPct = 100: Exit Function
    tot = WorkDays(s, e)
    done = WorkDays(s, Date)
    PlannedPct = 100 * done / tot
    If PlannedPct > 100 Then PlannedPct = 100
End Function

Public Function Weight(ByVal a As Variant, ByVal i As Long) As Double
    Weight = NumOr(a(i, T_PLAN), 0)
    If Weight <= 0 Then Weight = WorkDays(ToDate(a(i, T_START)), ToDate(a(i, T_DUE)))
End Function

' TaskAll 계산 열 채우기
Public Sub RecalcTaskAll()
    Dim lo As ListObject, a As Variant, i As Long
    Set lo = GetLO(TB_ALL)
    a = ReadTable(lo)
    If IsEmpty(a) Then Exit Sub
    For i = 1 To UBound(a, 1)
        a(i, A_LATE) = IIf(IsLateRow(a, i), "Y", Empty)
        If IsDateVal(a(i, T_START)) And IsDateVal(a(i, T_DUE)) Then
            If NumOr(a(i, T_PLAN)) <= 0 Then a(i, T_PLAN) = WorkDays(ToDate(a(i, T_START)), ToDate(a(i, T_DUE)))
            a(i, A_PLANPCT) = Round(PlannedPct(ToDate(a(i, T_START)), ToDate(a(i, T_DUE))), 0)
            a(i, A_GAP) = Round(NumOr(a(i, T_PCT)) - a(i, A_PLANPCT), 0)
            a(i, A_DDAY) = ToDate(a(i, T_DUE)) - Date
        End If
        If IsDateVal(a(i, T_UPDATED)) Then a(i, A_STALE) = Date - ToDate(a(i, T_UPDATED))
    Next i
    lo.DataBodyRange.Value = a
End Sub

'==============================================================
' 집계 결과: 프로젝트별 통계
'   key = ProjectID → Array(0 실적%, 1 계획%, 2 Task수, 3 완료, 4 진행, 5 지연, 6 참여자,
'                           7 시작, 8 종료, 9 현재단계, 10 단계표시, 11 가중치합, 12 신호)
'==============================================================
Public Function ProjectStats(ByVal a As Variant) As Object
    Dim d As Object, i As Long, pid As String, v As Variant, w As Double, parts As Object
    Dim key As Variant, curT As Variant, strp As Variant
    Set d = CreateObject("Scripting.Dictionary")
    Set parts = CreateObject("Scripting.Dictionary")
    For i = 1 To RowCount(a)
        If IsWorkTask(a, i) Then
            pid = CStr(a(i, T_PID))
            If Not d.Exists(pid) Then
                d(pid) = Array(0#, 0#, 0&, 0&, 0&, 0&, "", ToDate(a(i, T_START)), ToDate(a(i, T_DUE)), "", "", 0#, "")
                parts.Add pid, CreateObject("Scripting.Dictionary")
            End If
            v = d(pid)
            w = Weight(a, i)
            v(0) = v(0) + w * NumOr(a(i, T_PCT))                       ' 임시: 가중 합
            v(1) = v(1) + w * PlannedPct(ToDate(a(i, T_START)), ToDate(a(i, T_DUE)))
            v(11) = v(11) + w
            v(2) = v(2) + 1
            If CStr(a(i, T_STATUS)) = ST_DONE Then v(3) = v(3) + 1
            If CStr(a(i, T_STATUS)) = ST_DOING Then v(4) = v(4) + 1
            If IsLateRow(a, i) Then v(5) = v(5) + 1
            If ToDate(a(i, T_START)) < v(7) Then v(7) = ToDate(a(i, T_START))
            If ToDate(a(i, T_DUE)) > v(8) Then v(8) = ToDate(a(i, T_DUE))
            d(pid) = v
            AddPart parts(pid), CStr(Nz(a(i, T_OWNER)))
        End If
    Next i
    For Each key In d.Keys
        v = d(key)
        If v(11) > 0 Then v(0) = v(0) / v(11): v(1) = v(1) / v(11)
        v(6) = JoinKeys(parts(key))
        curT = "": strp = ""
        PhaseStatus CStr(key), a, curT, strp
        v(9) = curT: v(10) = strp
        v(12) = Signal(v(0), v(1), v(5), v(8))
        d(key) = v
    Next key
    Set ProjectStats = d
End Function

Private Sub AddPart(ByVal dd As Object, ByVal nm As String)
    If nm <> "" Then dd(nm) = True
End Sub

Private Function JoinKeys(ByVal dd As Object) As String
    Dim k As Variant
    For Each k In dd.Keys
        JoinKeys = JoinKeys & IIf(JoinKeys = "", "", ", ") & k
    Next k
End Function

' 신호: 완료 / 정상 / 주의 / 지연
Public Function Signal(ByVal pct As Double, ByVal plan As Double, ByVal nLate As Long, ByVal endD As Date) As String
    If pct >= 99.5 Then
        Signal = "완료"
    ElseIf (endD < Date) Or (pct - plan < -20) Then
        Signal = "지연"
    ElseIf (pct - plan < -10) Or nLate > 0 Then
        Signal = "주의"
    Else
        Signal = "정상"
    End If
End Function

' 팀 전체 진척 (실적, 계획)
Public Sub OverallPct(ByVal a As Variant, ByRef pct As Double, ByRef plan As Double)
    Dim i As Long, w As Double, sw As Double, sp As Double, spl As Double
    For i = 1 To RowCount(a)
        If IsWorkTask(a, i) Then
            w = Weight(a, i)
            sw = sw + w
            sp = sp + w * NumOr(a(i, T_PCT))
            spl = spl + w * PlannedPct(ToDate(a(i, T_START)), ToDate(a(i, T_DUE)))
        End If
    Next i
    If sw > 0 Then pct = sp / sw: plan = spl / sw
End Sub

'==============================================================
' 단계
'==============================================================
' 프로젝트 단계 목록 (순서대로, tblPhase 열) → 2차원 배열 / 없으면 Empty
Public Function GetPhases(ByVal pid As String) As Variant
    Dim a As Variant, i As Long, coll As Collection
    a = ReadTable(GetLO(TB_PHASE))
    Set coll = New Collection
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, PH_PID))) = pid Then coll.Add RowOf(a, i)
    Next i
    If coll.Count = 0 Then Exit Function
    GetPhases = SortPhases(CollTo2D(coll, PH_COLS))
End Function

' 단계 하나의 통계 → Array(실적%, 계획%, Task수, 담당자, 시작, 종료, 지연여부, 진행중여부)
Public Function PhaseStat(ByVal a As Variant, ByVal pid As String, ByVal phaseId As String, _
                          ByVal planS As Variant, ByVal planE As Variant) As Variant
    Dim i As Long, w As Double, sw As Double, sp As Double, spl As Double, n As Long, doing As Boolean
    Dim s As Date, e As Date, owners As Object, pct As Double, plan As Double
    Set owners = CreateObject("Scripting.Dictionary")
    For i = 1 To RowCount(a)
        If IsWorkTask(a, i) Then
            If CStr(a(i, T_PID)) = pid And CStr(Nz(a(i, T_PHASE))) = phaseId Then
                w = Weight(a, i)
                sw = sw + w
                sp = sp + w * NumOr(a(i, T_PCT))
                n = n + 1
                If s = 0 Or ToDate(a(i, T_START)) < s Then s = ToDate(a(i, T_START))
                If ToDate(a(i, T_DUE)) > e Then e = ToDate(a(i, T_DUE))
                If CStr(a(i, T_STATUS)) = ST_DOING Or (NumOr(a(i, T_PCT)) > 0 And NumOr(a(i, T_PCT)) < 100) Then doing = True
                AddPart owners, CStr(Nz(a(i, T_OWNER)))
            End If
        End If
    Next i
    If IsDateVal(planS) Then s = ToDate(planS)
    If IsDateVal(planE) Then e = ToDate(planE)
    If sw > 0 Then pct = sp / sw
    If s > 0 And e > 0 Then plan = PlannedPct(s, e)
    PhaseStat = Array(pct, plan, n, JoinKeys(owners), s, e, (e > 0 And e < Date And pct < 100 And n > 0), doing)
End Function

' 현재 단계 문구와 단계 표시(■완료 ◐진행 □미착수 ▲지연)
Public Sub PhaseStatus(ByVal pid As String, ByVal a As Variant, ByRef curText As Variant, ByRef strip As Variant)
    Dim ph As Variant, p As Long, st As Variant, cur As Long, firstOpen As Long
    ph = GetPhases(pid)
    curText = "": strip = ""
    If IsEmpty(ph) Then Exit Sub
    For p = 1 To UBound(ph, 1)
        st = PhaseStat(a, pid, CStr(ph(p, PH_ID)), ph(p, PH_START), ph(p, PH_END))
        If st(0) >= 99.5 And st(2) > 0 Then
            strip = strip & "■"
        ElseIf st(6) Then
            strip = strip & SYM_DELAY
            If cur = 0 Then cur = p
        ElseIf st(7) Then
            strip = strip & "◐"
            If cur = 0 Then cur = p
        Else
            strip = strip & "□"
            If firstOpen = 0 Then firstOpen = p
        End If
    Next p
    If cur = 0 Then cur = firstOpen
    If cur > 0 Then curText = cur & "/" & UBound(ph, 1) & " " & ph(cur, PH_NAME)
End Sub

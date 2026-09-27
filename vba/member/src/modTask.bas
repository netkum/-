Attribute VB_Name = "modTask"
'==============================================================
' Task 추가 / 수정 / 삭제 / 자동 보정
'==============================================================
Option Explicit

Public Function TaskLO() As ListObject
    Set TaskLO = GetLO(TB_TASK)
End Function

Public Function LoadTasks() As Variant
    LoadTasks = ReadTable(TaskLO())
End Function

' 표시·집계 대상이 되는 정상 Task 여부
Public Function IsValidTask(ByVal t As Variant, ByVal i As Long) As Boolean
    If CStr(Nz(t(i, T_NAME))) = "" Then Exit Function
    If CStr(Nz(t(i, T_PID))) = "" Then Exit Function
    If Not IsDateVal(t(i, T_START)) Then Exit Function
    If Not IsDateVal(t(i, T_DUE)) Then Exit Function
    IsValidTask = True
End Function

Public Function IsDelayed(ByVal t As Variant, ByVal i As Long) As Boolean
    Dim st As String
    st = CStr(Nz(t(i, T_STATUS)))
    If st = ST_DONE Or st = ST_CANCEL Then Exit Function
    If IsAbsence(CStr(t(i, T_PID))) Then Exit Function
    IsDelayed = (ToDate(t(i, T_DUE)) < Date And NumOr(t(i, T_PCT)) < 100)
End Function

Public Function GetTaskArr(ByVal id As String) As Variant
    Dim lo As ListObject, r As Long
    Set lo = TaskLO()
    r = FindRow(lo, T_ID, id)
    If r > 0 Then GetTaskArr = RowValues(lo, r)
End Function

Public Function NewTaskArr() As Variant
    Dim v() As Variant
    ReDim v(1 To T_COLS)
    v(T_PCT) = 0
    v(T_STATUS) = ST_PLAN
    v(T_PRI) = "중"
    NewTaskArr = v
End Function

' 새 TaskID: 이니셜-0001
Public Function NextTaskID() As String
    Dim ini As String, seq As Long, id As String, lo As ListObject
    ini = UCase$(CStr(CfgGet("Initial", "")))
    If ini = "" Then ini = "T"
    Set lo = TaskLO()
    seq = CLng(NumOr(CfgGet("LastSeq", 0)))
    Do
        seq = seq + 1
        id = ini & "-" & Format$(seq, "0000")
    Loop While FindRow(lo, T_ID, id) > 0
    CfgSet "LastSeq", seq
    NextTaskID = id
End Function

' 자동 열 채우기 + 상태/진척 자동 보정 (v: 1..T_COLS)
Public Sub FillAuto(ByRef v As Variant, Optional ByVal pd As Object = Nothing)
    Dim st As String, pct As Double, pid As String

    pid = Trim$(CStr(Nz(v(T_PID))))
    v(T_PID) = pid
    v(T_PNAME) = ProjectNameOf(pid, pd)
    v(T_OWNER) = OwnerName()

    pct = NumOr(v(T_PCT), 0)
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100
    st = CStr(Nz(v(T_STATUS)))
    If st = "" Then st = ST_PLAN

    If st = ST_DONE Then
        pct = 100
    ElseIf pct >= 100 And st <> ST_CANCEL Then
        st = ST_DONE
    ElseIf pct > 0 And st = ST_PLAN Then
        st = ST_DOING
    End If
    If st = ST_DONE Then
        If Not IsDateVal(v(T_DONE)) Then v(T_DONE) = Date
    Else
        v(T_DONE) = Empty
    End If
    v(T_PCT) = pct
    v(T_STATUS) = st
    If CStr(Nz(v(T_PRI))) = "" Then v(T_PRI) = "중"

    If IsDateVal(v(T_START)) Then v(T_START) = ToDate(v(T_START))
    If IsDateVal(v(T_DUE)) Then v(T_DUE) = ToDate(v(T_DUE))
    If IsDateVal(v(T_START)) And IsDateVal(v(T_DUE)) Then
        If v(T_DUE) < v(T_START) Then v(T_DUE) = v(T_START)
        v(T_PLAN) = WorkDays(v(T_START), v(T_DUE))
    Else
        v(T_PLAN) = Empty
    End If
    If UCase$(CStr(Nz(v(T_MILE)))) = "Y" Then v(T_MILE) = "Y" Else v(T_MILE) = Empty
    If Not IsDateVal(v(T_CREATED)) Then v(T_CREATED) = Now
    v(T_UPDATED) = Now
End Sub

' 입력폼에서 저장: ID가 없으면 신규
Public Function SaveTask(ByVal v As Variant) As String
    Dim lo As ListObject, r As Long
    Set lo = TaskLO()
    SpeedOn
    If CStr(Nz(v(T_ID))) = "" Then v(T_ID) = NextTaskID()
    FillAuto v
    r = FindRow(lo, T_ID, CStr(v(T_ID)))
    If r = 0 Then
        AppendRow lo, v
    Else
        SetRowValues lo, r, v
    End If
    SpeedOff
    SaveTask = CStr(v(T_ID))
End Function

Public Sub DeleteTask(ByVal id As String)
    Dim lo As ListObject, r As Long
    Set lo = TaskLO()
    r = FindRow(lo, T_ID, id)
    If r = 0 Then Exit Sub
    SpeedOn
    lo.ListRows(r).Delete
    SpeedOff
End Sub

' TaskDB 시트를 직접 수정했을 때 해당 행들 보정
Public Sub FixEditedRows(ByVal target As Range)
    Dim lo As ListObject, rng As Range, rw As Range, idx As Long
    Dim v As Variant, pd As Object, done As Object
    Set lo = TaskLO()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rng = Intersect(target, lo.DataBodyRange)
    If rng Is Nothing Then Exit Sub
    Set pd = ProjectDict()
    Set done = CreateObject("Scripting.Dictionary")
    SpeedOn
    For Each rw In rng.Rows
        idx = rw.Row - lo.DataBodyRange.Row + 1
        If Not done.Exists(idx) Then
            done(idx) = True
            v = RowValues(lo, idx)
            If CStr(Nz(v(T_NAME))) <> "" Or CStr(Nz(v(T_PID))) <> "" Then
                If CStr(Nz(v(T_ID))) = "" Then v(T_ID) = NextTaskID()
                FillAuto v, pd
                SetRowValues lo, idx, v
            End If
        End If
    Next rw
    SpeedOff
End Sub

' 프로젝트명·계획일수 전체 재계산 (프로젝트/공휴일 갱신 후)
Public Sub RecalcAllTasks()
    Dim lo As ListObject, a As Variant, i As Long, pd As Object
    Set lo = TaskLO()
    a = ReadTable(lo)
    If IsEmpty(a) Then Exit Sub
    Set pd = ProjectDict()
    For i = 1 To UBound(a, 1)
        a(i, T_PNAME) = ProjectNameOf(CStr(Nz(a(i, T_PID))), pd)
        If IsDateVal(a(i, T_START)) And IsDateVal(a(i, T_DUE)) Then
            a(i, T_PLAN) = WorkDays(ToDate(a(i, T_START)), ToDate(a(i, T_DUE)))
        End If
    Next i
    lo.DataBodyRange.Value = a
End Sub

' 특정 열의 코드 일괄 교체 (프로젝트·단계 승인/통합 반영용), 교체 건수 반환
Public Function ReplaceCode(ByVal colIdx As Long, ByVal fromCode As String, ByVal toCode As String) As Long
    Dim lo As ListObject, a As Variant, i As Long, n As Long
    Set lo = TaskLO()
    a = ReadTable(lo)
    If IsEmpty(a) Then Exit Function
    For i = 1 To UBound(a, 1)
        If CStr(Nz(a(i, colIdx))) = fromCode Then a(i, colIdx) = toCode: n = n + 1
    Next i
    If n > 0 Then lo.DataBodyRange.Value = a
    ReplaceCode = n
End Function

' 단계ID 접두어 교체 (N-홍길동-01-02 → P005-02)
Public Sub ReplacePhasePrefix(ByVal fromPid As String, ByVal toPid As String)
    Dim lo As ListObject, a As Variant, i As Long, s As String, changed As Boolean
    Set lo = TaskLO()
    a = ReadTable(lo)
    If IsEmpty(a) Then Exit Sub
    For i = 1 To UBound(a, 1)
        s = CStr(Nz(a(i, T_PHASE)))
        If Left$(s, Len(fromPid) + 1) = fromPid & "-" Then
            If toPid = "" Then
                a(i, T_PHASE) = Empty
            Else
                a(i, T_PHASE) = toPid & Mid$(s, Len(fromPid) + 1)
            End If
            changed = True
        End If
    Next i
    If changed Then lo.DataBodyRange.Value = a
End Sub

'--- 상세 패널 빠른 버튼 ---------------------------------------
Public Sub CompleteTask(ByVal id As String)
    Dim v As Variant
    v = GetTaskArr(id)
    If IsEmpty(v) Then Exit Sub
    v(T_STATUS) = ST_DONE
    v(T_PCT) = 100
    SaveTask v
End Sub

Public Sub ExtendTask(ByVal id As String, Optional ByVal days As Long = 1)
    Dim v As Variant, oldDue As Date, newDue As Date
    v = GetTaskArr(id)
    If IsEmpty(v) Then Exit Sub
    oldDue = ToDate(v(T_DUE))
    newDue = AddWorkDays(oldDue, days)
    v(T_DUE) = newDue
    v(T_NOTE) = Trim$(CStr(Nz(v(T_NOTE))) & " [연장 " & Format$(oldDue, "m/d") & "→" & Format$(newDue, "m/d") & "]")
    SaveTask v
End Sub

' PlanDays 가중 진척률 (idx: 행 번호 컬렉션)
Public Function WeightedPct(ByVal t As Variant, ByVal idx As Collection) As Double
    Dim k As Variant, w As Double, sw As Double, sp As Double
    For Each k In idx
        w = NumOr(t(k, T_PLAN), 0)
        If w <= 0 Then w = WorkDays(ToDate(t(k, T_START)), ToDate(t(k, T_DUE)))
        sw = sw + w
        sp = sp + w * NumOr(t(k, T_PCT), 0)
    Next k
    If sw > 0 Then WeightedPct = sp / sw
End Function

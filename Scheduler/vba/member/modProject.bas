Attribute VB_Name = "modProject"
'==============================================================
' 프로젝트 요청 / PM 여부 / 단계(Phase) 관리
'==============================================================
Option Explicit

' 내가 이 프로젝트의 PM인가? (팀장이 지정했거나 1인 프로젝트 자동 PM)
Public Function IsPM(ByVal pid As String) As Boolean
    Dim pd As Object
    If pid = "" Or IsSpecialProject(pid) Then Exit Function
    Set pd = ProjectDict()
    If Not pd.Exists(pid) Then Exit Function
    IsPM = (pd(pid)(1) = OwnerName())
End Function

'--- 임시 프로젝트 요청 -------------------------------------
' 반환: 임시 코드 N-이름-01
Public Function CreateRequest(ByVal pname As String, ByVal s As Date, ByVal e As Date, ByVal memo As String) As String
    Dim lo As ListObject, k As Long, code As String, pd As Object
    Set lo = GetLO(TB_REQ)
    Set pd = ProjectDict()
    Do
        k = k + 1
        code = "N-" & OwnerName() & "-" & Format$(k, "00")
    Loop While FindRow(lo, RQ_CODE, code) > 0 Or pd.Exists(code)
    SpeedOn
    AppendRow lo, Array(code, pname, IIf(s = 0, Empty, s), IIf(e = 0, Empty, e), memo, Now, Empty, Empty, Empty)
    AppendRow GetLO(TB_PROJ), Array(code, pname, OwnerName(), Empty, PJ_PENDING)
    SpeedOff
    RefreshLists
    CreateRequest = code
End Function

'--- 단계 조회 ----------------------------------------------
Public Function PhaseNameOf(ByVal phaseId As String) As String
    Dim lo As ListObject, r As Long
    Set lo = GetLO(TB_PHASE)
    r = FindRow(lo, PH_ID, phaseId)
    If r > 0 Then
        PhaseNameOf = CStr(lo.DataBodyRange.Cells(r, PH_NAME).Value)
    Else
        PhaseNameOf = phaseId
    End If
End Function

' 프로젝트의 단계 목록 (순서대로) → 2차원 배열 (열 = tblPhase 열), 없으면 Empty
Public Function GetPhases(ByVal pid As String) As Variant
    Dim a As Variant, i As Long, coll As Collection, res As Variant, j As Long, k As Long, tmp As Variant
    a = ReadTable(GetLO(TB_PHASE))
    Set coll = New Collection
    For i = 1 To RowCount(a)
        If CStr(a(i, PH_PID)) = pid Then coll.Add RowOf(a, i)
    Next i
    If coll.Count = 0 Then Exit Function
    res = CollTo2D(coll, PH_COLS)
    ' 순서 정렬 (단순 삽입 정렬)
    For i = 2 To UBound(res, 1)
        For j = i To 2 Step -1
            If NumOr(res(j, PH_ORDER)) < NumOr(res(j - 1, PH_ORDER)) Then
                For k = 1 To PH_COLS
                    tmp = res(j, k): res(j, k) = res(j - 1, k): res(j - 1, k) = tmp
                Next k
            Else
                Exit For
            End If
        Next j
    Next i
    GetPhases = res
End Function

' 새 단계 ID: P001-03 (기존 + 편집 중 목록 중 최댓값 + 1)
Public Function NewPhaseID(ByVal pid As String, Optional ByVal extraIds As Variant) As String
    Dim a As Variant, i As Long, mx As Long, s As String
    a = ReadTable(GetLO(TB_PHASE))
    For i = 1 To RowCount(a)
        mx = MaxSeq(mx, CStr(a(i, PH_ID)), pid)
    Next i
    a = ReadTable(GetLO(TB_MYPHASE))
    For i = 1 To RowCount(a)
        mx = MaxSeq(mx, CStr(a(i, PH_ID)), pid)
    Next i
    If IsArray(extraIds) Then
        For i = LBound(extraIds) To UBound(extraIds)
            mx = MaxSeq(mx, CStr(extraIds(i)), pid)
        Next i
    End If
    NewPhaseID = pid & "-" & Format$(mx + 1, "00")
End Function

Private Function MaxSeq(ByVal mx As Long, ByVal phaseId As String, ByVal pid As String) As Long
    Dim tail As String
    MaxSeq = mx
    If Left$(phaseId, Len(pid) + 1) <> pid & "-" Then Exit Function
    tail = Mid$(phaseId, Len(pid) + 2)
    If IsNumeric(tail) Then
        If CLng(tail) > mx Then MaxSeq = CLng(tail)
    End If
End Function

Public Function CountMyTasksByPhase(ByVal phaseId As String) As Long
    Dim t As Variant, i As Long
    t = LoadTasks()
    For i = 1 To RowCount(t)
        If CStr(Nz(t(i, T_PHASE))) = phaseId Then CountMyTasksByPhase = CountMyTasksByPhase + 1
    Next i
End Function

'--- PM 단계 저장 -------------------------------------------
' rows: 1차원 배열들의 배열 (각: PhaseID, 단계명, 계획시작, 계획종료) / merges: Array(from, to)의 컬렉션
Public Sub SavePhases(ByVal pid As String, ByVal rows As Variant, ByVal n As Long, ByVal merges As Collection)
    Dim i As Long, newRows As Collection, m As Variant, stamp As Date, r As Variant

    stamp = Now
    Set newRows = New Collection
    For i = 1 To n
        r = rows(i)
        newRows.Add Array(r(0), pid, i, r(1), IIf(IsDateVal(r(2)), r(2), Empty), IIf(IsDateVal(r(3)), r(3), Empty), OwnerName(), stamp)
    Next i

    SpeedOn
    ReplaceProjectRows GetLO(TB_MYPHASE), pid, newRows
    ReplaceProjectRows GetLO(TB_PHASE), pid, newRows
    For Each m In merges
        ReplaceCode T_PHASE, CStr(m(0)), CStr(m(1))
        AppendRow GetLO(TB_PHMAP), Array(m(0), m(1), stamp)
    Next m
    SpeedOff
End Sub

' 표에서 ProjectID(2열)가 pid인 행을 newRows로 교체
Public Sub ReplaceProjectRows(ByVal lo As ListObject, ByVal pid As String, ByVal newRows As Collection)
    Dim a As Variant, i As Long, keep As Collection, v As Variant
    Set keep = New Collection
    a = ReadTable(lo)
    For i = 1 To RowCount(a)
        If CStr(a(i, PH_PID)) <> pid Then keep.Add RowOf(a, i)
    Next i
    For Each v In newRows
        keep.Add v
    Next v
    WriteTable lo, CollTo2D(keep, PH_COLS)
End Sub

' MyProject / 입력폼의 [단계 편집] 버튼
Public Sub EditPhasesButton()
    Dim pid As String
    If ActiveSheet.Name = SH_MP Then pid = FilterPid(GetWS(SH_MP).Range(MP_FILTER).Value)
    If pid = "" Then
        Msg "단계를 편집할 프로젝트를 위쪽 프로젝트 목록에서 먼저 선택하세요."
        Exit Sub
    End If
    If OpenPhaseForm(pid) Then
        If ActiveSheet.Name = SH_MP Then RenderMyProject
    End If
End Sub

Public Function OpenPhaseForm(ByVal pid As String) As Boolean
    Dim f As frmPhase
    If Not IsPM(pid) Then
        Msg "단계는 프로젝트 PM만 편집할 수 있습니다." & vbLf & _
            "(참여자가 2명 이상인 프로젝트의 PM은 팀장이 지정합니다)", vbExclamation
        Exit Function
    End If
    Set f = New frmPhase
    f.Init pid
    f.Show
    OpenPhaseForm = f.Saved
    Unload f
End Function

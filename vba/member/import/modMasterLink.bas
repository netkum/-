Attribute VB_Name = "modMasterLink"
'==============================================================
' 팀장 Master가 만든 ProjectList.xlsx 읽기 (파일 열 때 자동)
'   \Master\ProjectList.xlsx 시트 (1행 머리글, 2행부터 데이터)
'   Projects : ProjectID | ProjectName | PM | Color | Status
'   Phases   : PhaseID | ProjectID | Order | PhaseName | PlanStart | PlanEnd | Author | ModifiedAt
'   CodeMap  : FromCode | ToCode | Type(Project/Phase) | Result(승인/통합/반려) | Reason
'   Comments : TaskID | Comment | Date
'   Events   : Date | Time | Title | Target(전체 또는 이름)
'   Holidays : Date | Name            (선택)
'==============================================================
Option Explicit

Public Function ProjectListPath() As String
    ProjectListPath = MasterFolder() & "\ProjectList.xlsx"
End Function

' 새로고침 버튼
Public Sub RefreshAll()
    SyncFromMaster False
    RenderCalendar
End Sub

Public Sub SyncFromMaster(Optional ByVal silent As Boolean = True)
    Dim path As String, wb As Workbook, msgs As String, n As Long
    Dim projA As Variant, phA As Variant, mapA As Variant, cmA As Variant, evA As Variant, holA As Variant

    path = ProjectListPath()
    If Dir(path) = "" Then
        If Not silent Then Msg "팀장 파일(ProjectList.xlsx)을 찾을 수 없습니다." & vbLf & path & vbLf & vbLf & _
                               "Setting 시트의 MasterFolder 값을 확인하세요.", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH
    SpeedOn
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Notify:=False)
    projA = SheetData(wb, "Projects", 5)
    phA = SheetData(wb, "Phases", PH_COLS)
    mapA = SheetData(wb, "CodeMap", 5)
    cmA = SheetData(wb, "Comments", 3)
    evA = SheetData(wb, "Events", 4)
    holA = SheetData(wb, "Holidays", 2)
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If Not IsEmpty(holA) Then WriteTable GetLO(TB_HOL), holA
    ApplyProjects projA
    ApplyPhases phA
    msgs = ApplyCodeMap(mapA)
    n = ApplyComments(cmA)
    WriteTable GetLO(TB_EVENT), evA
    RefreshLists
    RecalcAllTasks
    CfgSet "LastMasterSync", Now
    SpeedOff

    If n > 0 Then msgs = msgs & vbLf & SYM_COMMENT & " 새 팀장 코멘트 " & n & "건이 있습니다. (캘린더에 " & SYM_COMMENT & " 표시)"
    If msgs <> "" Then
        Msg Mid$(msgs, 2)
    ElseIf Not silent Then
        Msg "팀장 파일에서 프로젝트·단계·보고일정을 갱신했습니다."
    End If
    Exit Sub
EH:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    SpeedReset
    If Not silent Then Msg "팀장 파일을 읽는 중 오류: " & Err.Description, vbExclamation
End Sub

' 시트 데이터 (2행부터, nCols 열) → 2차원 배열 / 시트가 없거나 비었으면 Empty
Private Function SheetData(ByVal wb As Workbook, ByVal sheetName As String, ByVal nCols As Long) As Variant
    Dim sh As Worksheet, lastR As Long
    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then Exit Function
    lastR = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function
    SheetData = sh.Range(sh.Cells(2, 1), sh.Cells(lastR, nCols)).Value
End Function

'--- Projects -----------------------------------------------
Private Sub ApplyProjects(ByVal projA As Variant)
    Dim coll As Collection, i As Long, rq As Variant, have As Object
    If IsEmpty(projA) Then Exit Sub
    Set coll = New Collection
    Set have = CreateObject("Scripting.Dictionary")
    For i = 1 To RowCount(projA)
        If CStr(Nz(projA(i, 1))) <> "" Then
            coll.Add RowOf(projA, i)
            have(CStr(projA(i, 1))) = True
        End If
    Next i
    ' 아직 처리 안 된 내 임시 프로젝트는 계속 유지
    rq = ReadTable(GetLO(TB_REQ))
    For i = 1 To RowCount(rq)
        If CStr(Nz(rq(i, RQ_RESULT))) = "" And Not have.Exists(CStr(rq(i, RQ_CODE))) Then
            coll.Add Array(rq(i, RQ_CODE), rq(i, RQ_NAME), OwnerName(), Empty, PJ_PENDING)
        End If
    Next i
    WriteTable GetLO(TB_PROJ), CollTo2D(coll, 5)
End Sub

'--- Phases -------------------------------------------------
' 규칙: 내가 PM인 프로젝트는 "내 단계(tblMyPhase)"와 "Master 단계" 중 수정일시가 최신인 쪽 사용
'       (팀장이 직접 고쳤거나 PM이 바뀐 경우 Master가 최신 → 내 단계를 교체)
Private Sub ApplyPhases(ByVal phA As Variant)
    Dim myA As Variant, pd As Object, pid As Variant, i As Long
    Dim cache As Collection, mine As Collection, mMax As Object, yMax As Object, pmSet As Object

    Set pd = ProjectDict()
    myA = ReadTable(GetLO(TB_MYPHASE))
    Set mMax = CreateObject("Scripting.Dictionary")
    Set yMax = CreateObject("Scripting.Dictionary")
    Set pmSet = CreateObject("Scripting.Dictionary")

    For Each pid In pd.Keys
        If pd(pid)(1) = OwnerName() Then pmSet(CStr(pid)) = True
    Next pid

    For i = 1 To RowCount(phA)
        pid = CStr(Nz(phA(i, PH_PID)))
        If StampOf(phA(i, PH_MOD)) > NumOr(mMax(pid)) Then mMax(pid) = StampOf(phA(i, PH_MOD))
    Next i
    For i = 1 To RowCount(myA)
        pid = CStr(Nz(myA(i, PH_PID)))
        If StampOf(myA(i, PH_MOD)) > NumOr(yMax(pid)) Then yMax(pid) = StampOf(myA(i, PH_MOD))
    Next i

    Set cache = New Collection
    Set mine = New Collection
    ' Master 단계: 내가 PM이 아니거나, Master 쪽이 더 최신이면 Master 사용
    For i = 1 To RowCount(phA)
        pid = CStr(Nz(phA(i, PH_PID)))
        If pid <> "" Then
            If Not pmSet.Exists(pid) Then
                cache.Add RowOf(phA, i)
            ElseIf NumOr(mMax(pid)) >= NumOr(yMax(pid)) Then
                cache.Add RowOf(phA, i)
                mine.Add RowOf(phA, i)
            End If
        End If
    Next i
    ' 내 단계: 내가 PM이고 내 쪽이 더 최신(아직 동기화 전)이거나 Master에 없는 경우
    For i = 1 To RowCount(myA)
        pid = CStr(Nz(myA(i, PH_PID)))
        If pmSet.Exists(pid) Then
            If NumOr(yMax(pid)) > NumOr(mMax(pid)) Then
                cache.Add RowOf(myA, i)
                mine.Add RowOf(myA, i)
            End If
        End If
    Next i
    WriteTable GetLO(TB_PHASE), CollTo2D(cache, PH_COLS)
    WriteTable GetLO(TB_MYPHASE), CollTo2D(mine, PH_COLS)
End Sub

' 날짜+시각 → 숫자 (비교용)
Private Function StampOf(ByVal v As Variant) As Double
    If Not IsDateVal(v) Then Exit Function
    If VarType(v) = vbDate Or IsNumeric(v) Then
        StampOf = CDbl(v)
    Else
        StampOf = CDbl(CDate(v))
    End If
End Function

'--- CodeMap: 임시 프로젝트 승인/통합/반려, 단계 합치기 ---------
Private Function ApplyCodeMap(ByVal mapA As Variant) As String
    Dim i As Long, fromC As String, toC As String, tp As String, res As String, reason As String
    Dim lo As ListObject, r As Long, msgs As String, n As Long

    Set lo = GetLO(TB_REQ)
    For i = 1 To RowCount(mapA)
        fromC = CStr(Nz(mapA(i, 1))): toC = CStr(Nz(mapA(i, 2)))
        tp = CStr(Nz(mapA(i, 3))): res = CStr(Nz(mapA(i, 4))): reason = CStr(Nz(mapA(i, 5)))
        If fromC <> "" Then
            If tp = "Phase" Then
                If toC <> "" Then ReplaceCode T_PHASE, fromC, toC
            Else
                r = FindRow(lo, RQ_CODE, fromC)
                If res = "반려" Then
                    If r > 0 Then
                        If CStr(Nz(lo.DataBodyRange.Cells(r, RQ_RESULT).Value)) = "" Then
                            lo.DataBodyRange.Cells(r, RQ_RESULT).Value = "반려"
                            lo.DataBodyRange.Cells(r, RQ_REASON).Value = reason
                            msgs = msgs & vbLf & "프로젝트 요청 반려: " & fromC & IIf(reason <> "", " (" & reason & ")", "") & _
                                   vbLf & "   → 해당 Task를 다른 프로젝트로 옮겨 주세요."
                        End If
                    End If
                ElseIf toC <> "" Then
                    n = ReplaceCode(T_PID, fromC, toC)
                    ' 승인: 단계ID 접두어 교체 / 통합: 임시 프로젝트 단계는 해제
                    If res = "통합" Then
                        ReplacePhasePrefix fromC, ""
                    Else
                        ReplacePhasePrefix fromC, toC
                    End If
                    RemoveProjectRows GetLO(TB_MYPHASE), fromC
                    If r > 0 Then
                        If CStr(Nz(lo.DataBodyRange.Cells(r, RQ_RESULT).Value)) = "" Then
                            lo.DataBodyRange.Cells(r, RQ_RESULT).Value = IIf(res = "", "승인", res)
                            lo.DataBodyRange.Cells(r, RQ_MAPPED).Value = toC
                            msgs = msgs & vbLf & "프로젝트 요청 " & IIf(res = "통합", "기존 프로젝트로 통합", "승인") & _
                                   ": " & fromC & " → " & toC & " (Task " & n & "건 자동 변경)"
                        End If
                    End If
                End If
            End If
        End If
    Next i
    ApplyCodeMap = msgs
End Function

Private Sub RemoveProjectRows(ByVal lo As ListObject, ByVal pid As String)
    ReplaceProjectRows lo, pid, New Collection
End Sub

'--- Comments -----------------------------------------------
Private Function ApplyComments(ByVal cmA As Variant) As Long
    Dim d As Object, i As Long, lo As ListObject, t As Variant, newC As String, cnt As Long
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To RowCount(cmA)
        If CStr(Nz(cmA(i, 1))) <> "" Then d(CStr(cmA(i, 1))) = CStr(Nz(cmA(i, 2)))
    Next i
    Set lo = GetLO(TB_TASK)
    t = ReadTable(lo)
    If IsEmpty(t) Then Exit Function
    For i = 1 To UBound(t, 1)
        newC = ""
        If d.Exists(CStr(t(i, T_ID))) Then newC = d(CStr(t(i, T_ID)))
        If newC <> CStr(Nz(t(i, T_COMMENT))) Then
            If newC <> "" Then cnt = cnt + 1
            t(i, T_COMMENT) = IIf(newC = "", Empty, newC)
        End If
    Next i
    lo.DataBodyRange.Value = t
    ApplyComments = cnt
End Function

'==============================================================
' 드롭다운 목록 갱신 (Setting 시트 BB~BG 열 + 이름 정의)
'==============================================================
Public Sub RefreshLists()
    Dim ws As Worksheet, a As Variant, i As Long, projs As Collection, filt As Collection

    Set ws = GetWS(SH_SET)
    If ws Is Nothing Then Exit Sub
    SpeedOn
    SetList ws, "BB", "lstStatus", Array(ST_PLAN, ST_DOING, ST_DONE, ST_HOLD, ST_CANCEL)
    SetList ws, "BC", "lstPriority", Array("상", "중", "하")
    SetList ws, "BD", "lstCompletion", Array(0, 10, 25, 50, 75, 90, 100)
    SetList ws, "BG", "lstYN", Array("Y")

    Set projs = New Collection
    Set filt = New Collection
    filt.Add FILTER_ALL
    a = ReadTable(GetLO(TB_PROJ))
    For i = 1 To RowCount(a)
        If IsActiveProject(CStr(Nz(a(i, P_STATUS)))) And CStr(Nz(a(i, P_ID))) <> "" Then
            projs.Add CStr(a(i, P_ID))
            filt.Add CStr(a(i, P_ID)) & " " & CStr(Nz(a(i, P_NAME)))
        End If
    Next i
    projs.Add PJ_COMMON: projs.Add PJ_LEAVE: projs.Add PJ_TRIP: projs.Add PJ_EDU
    SetList ws, "BE", "lstProject", CollToArr(projs)
    SetList ws, "BF", "lstFilter", CollToArr(filt)
    SpeedOff
End Sub

Public Function IsActiveProject(ByVal st As String) As Boolean
    IsActiveProject = Not (st = "완료" Or st = "종료" Or st = "취소" Or st = "보류")
End Function

Private Function CollToArr(ByVal c As Collection) As Variant
    Dim a() As Variant, i As Long
    ReDim a(0 To c.Count - 1)
    For i = 1 To c.Count: a(i - 1) = c(i): Next i
    CollToArr = a
End Function

Private Sub SetList(ByVal ws As Worksheet, ByVal colL As String, ByVal nm As String, ByVal items As Variant)
    Dim i As Long, n As Long
    ws.Range(colL & "4:" & colL & "500").ClearContents
    ws.Range(colL & "3").Value = nm
    n = UBound(items) - LBound(items) + 1
    For i = 0 To n - 1
        ws.Range(colL & (4 + i)).Value = items(LBound(items) + i)
    Next i
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & ws.Name & "'!$" & colL & "$4:$" & colL & "$" & (3 + n)
End Sub

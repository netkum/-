Attribute VB_Name = "modSync"
'==============================================================
' [전체 동기화] : 팀원 파일 수집 → 통합 → 계산 → 배포 → 화면 갱신
'  1. Member 폴더의 *_업무일정.xlsm 검색
'  2. 읽기 전용으로 열어 tblTask / tblMyRequest / tblMyPhase / tblMyPhaseMap 읽기
'  3. 실패한 파일은 이전 데이터 유지
'  4. 신규/수정/삭제 비교 → TaskAll 교체
'  5. 프로젝트 요청·참여자·자동 PM·PM 단계·단계 합치기 반영
'  6. ProjectList.xlsx 배포 → 화면 갱신 → 결과창
'==============================================================
Option Explicit

Private mLog As Collection

Public Sub SyncAll()
    Dim folder As String, files As Collection, f As Variant, owner As String, path As String
    Dim prev As Variant, prevIdx As Object, newRows As Collection, seenID As Object
    Dim tasks As Variant, reqs As Variant, phs As Variant, phMap As Variant, ver As String, errMsg As String
    Dim i As Long, j As Long, row() As Variant, nNew As Long, nMod As Long, nDel As Long, nDup As Long
    Dim reqAll As Collection, phByOwner As Object, mapAll As Collection, result As String, okCnt As Long, ngCnt As Long
    Dim stamp As Date, lastUpd As Date, k As Variant

    If Not IsReady() Then Msg "먼저 SetupMaster를 실행하세요.", vbExclamation: Exit Sub
    folder = MemberFolder()
    If Dir(folder, vbDirectory) = "" Then
        Msg "팀원 폴더를 찾을 수 없습니다: " & folder & vbLf & "Setting 시트의 MemberFolder를 확인하세요.", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH
    Application.StatusBar = "팀원 파일 수집 중..."
    SpeedOn
    stamp = Now
    Set mLog = New Collection

    ' 파일 목록 먼저 수집 (Workbooks.Open 중 Dir 상태가 바뀌지 않도록)
    Set files = New Collection
    f = Dir(folder & "\" & MEMBER_PATTERN)
    Do While f <> ""
        files.Add CStr(f)
        f = Dir()
    Loop

    prev = ReadTable(GetLO(TB_ALL))
    Set prevIdx = CreateObject("Scripting.Dictionary")
    For i = 1 To RowCount(prev)
        prevIdx(CStr(prev(i, T_ID))) = i
    Next i

    Set newRows = New Collection
    Set seenID = CreateObject("Scripting.Dictionary")
    Set reqAll = New Collection
    Set mapAll = New Collection
    Set phByOwner = CreateObject("Scripting.Dictionary")

    For Each f In files
        owner = Split(CStr(f), "_")(0)
        path = folder & "\" & f
        Application.StatusBar = "수집 중: " & f
        errMsg = ReadMemberFile(path, tasks, reqs, phs, phMap, ver)
        lastUpd = 0
        If errMsg = "" Then
            j = 0
            For i = 1 To RowCount(tasks)
                If CStr(Nz(tasks(i, T_ID))) <> "" Then
                    ReDim row(1 To A_COLS)
                    For k = 1 To T_COLS: row(k) = tasks(i, k): Next k
                    If CStr(Nz(row(T_OWNER))) = "" Then row(T_OWNER) = owner
                    row(A_SRC) = owner
                    row(A_SYNC) = stamp
                    If seenID.Exists(CStr(row(T_ID))) Then nDup = nDup + 1
                    seenID(CStr(row(T_ID))) = True
                    If StampOf(row(T_UPDATED)) > CDbl(lastUpd) Then lastUpd = CDate(StampOf(row(T_UPDATED)))
                    newRows.Add row
                    j = j + 1
                End If
            Next i
            For i = 1 To RowCount(reqs)
                reqAll.Add Array(owner, RowOf(reqs, i))
            Next i
            If Not IsEmpty(phs) Then phByOwner(owner) = phs
            For i = 1 To RowCount(phMap)
                mapAll.Add RowOf(phMap, i)
            Next i
            result = "정상"
            okCnt = okCnt + 1
            UpdateMember owner, CStr(f), stamp, result, j, lastUpd, ver
            AddLog stamp, CStr(f), owner, result, j, ""
        Else
            ' 실패: 이전 데이터 유지
            j = 0
            For i = 1 To RowCount(prev)
                If CStr(Nz(prev(i, A_SRC))) = owner Then
                    newRows.Add RowOf(prev, i)
                    seenID(CStr(prev(i, T_ID))) = True
                    j = j + 1
                End If
            Next i
            result = "실패(이전 데이터 유지)"
            ngCnt = ngCnt + 1
            UpdateMember owner, CStr(f), Empty, result & ": " & errMsg, j, Empty, ""
            AddLog stamp, CStr(f), owner, "실패", j, errMsg
        End If
    Next f

    ' 신규 / 수정 / 삭제
    For Each k In newRows
        If Not prevIdx.Exists(CStr(k(T_ID))) Then
            nNew = nNew + 1
        ElseIf CStr(Nz(prev(prevIdx(CStr(k(T_ID))), T_UPDATED))) <> CStr(Nz(k(T_UPDATED))) Then
            nMod = nMod + 1
        End If
    Next k
    For i = 1 To RowCount(prev)
        If Not seenID.Exists(CStr(prev(i, T_ID))) Then nDel = nDel + 1
    Next i

    WriteTable GetLO(TB_ALL), CollTo2D(newRows, A_COLS)

    ' 요청·참여자·PM·단계
    Dim nReq As Long, nPhase As Long, nPM As Long
    nReq = CollectRequests(reqAll)
    UpdateParticipants
    nPhase = ApplyPMPhases(phByOwner)
    ApplyPhaseMerges mapAll
    FillComments
    RecalcTaskAll
    nPM = CountPMNeeded()
    CfgSet "LastSync", stamp
    TrimLog
    SpeedOff

    ExportProjectList True
    RefreshViews
    Application.StatusBar = False

    Msg "업무일정 동기화 완료   " & Format$(stamp, "yyyy-mm-dd hh:mm") & vbLf & vbLf & _
        MemberSummaryText() & vbLf & _
        "총 " & newRows.Count & "건 │ 신규 " & nNew & " │ 수정 " & nMod & " │ 삭제 " & nDel & " │ 지연 " & CountLate() & vbLf & _
        IIf(nReq > 0, "새 프로젝트 요청 " & nReq & "건 → ProjectMaster 승인대기 확인" & vbLf, "") & _
        IIf(nPhase > 0, "단계 변경 " & nPhase & "개 프로젝트 반영" & vbLf, "") & _
        IIf(nPM > 0, SYM_DELAY & " PM 지정 필요 " & nPM & "건 (참여자 2명 이상)" & vbLf, "") & _
        IIf(nDup > 0, SYM_DELAY & " 중복 TaskID " & nDup & "건 (같은 이니셜 사용 여부 확인)" & vbLf, "") & _
        IIf(ngCnt > 0, SYM_DELAY & " 읽기 실패 " & ngCnt & "명 (이전 데이터 유지)", "")
    Exit Sub
EH:
    SpeedReset
    Application.StatusBar = False
    Msg "동기화 중 오류: " & Err.Description, vbCritical
End Sub

' 팀원 파일 읽기 → 오류 메시지 (정상이면 "")
Private Function ReadMemberFile(ByVal path As String, ByRef tasks As Variant, ByRef reqs As Variant, _
                                ByRef phs As Variant, ByRef phMap As Variant, ByRef ver As String) As String
    Dim wb As Workbook, a As Variant
    tasks = Empty: reqs = Empty: phs = Empty: phMap = Empty: ver = ""
    On Error GoTo EH
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Notify:=False)
    If FindTable(wb, "tblTask") Is Nothing Then
        ReadMemberFile = "tblTask 표 없음"
    Else
        tasks = TableData(wb, "tblTask", T_COLS)
        reqs = TableData(wb, "tblMyRequest", 9)
        phs = TableData(wb, "tblMyPhase", PH_COLS)
        phMap = TableData(wb, "tblMyPhaseMap", 3)
        a = TableData(wb, "tblSetting", 2)
        ver = SettingOf(a, "Version")
    End If
    wb.Close SaveChanges:=False
    Exit Function
EH:
    ReadMemberFile = Err.Description
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

Private Function FindTable(ByVal wb As Workbook, ByVal nm As String) As ListObject
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        On Error Resume Next
        Set FindTable = sh.ListObjects(nm)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next sh
End Function

Private Function TableData(ByVal wb As Workbook, ByVal nm As String, ByVal nCols As Long) As Variant
    Dim lo As ListObject
    Set lo = FindTable(wb, nm)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If lo.ListColumns.Count < nCols Then nCols = lo.ListColumns.Count
    TableData = lo.DataBodyRange.Resize(, nCols).Value
End Function

Private Function SettingOf(ByVal a As Variant, ByVal key As String) As String
    Dim i As Long
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, 1))) = key Then SettingOf = CStr(Nz(a(i, 2))): Exit Function
    Next i
End Function

'--- 팀원 목록 / 로그 ----------------------------------------
Private Sub UpdateMember(ByVal owner As String, ByVal fileName As String, ByVal syncAt As Variant, _
                         ByVal result As String, ByVal rows As Long, ByVal lastUpd As Variant, ByVal ver As String)
    Dim lo As ListObject, r As Long
    Set lo = GetLO(TB_MEMBER)
    r = FindRow(lo, M_NAME, owner)
    If r = 0 Then r = AppendRow(lo, Array(owner))
    With lo.DataBodyRange
        .Cells(r, M_FILE).Value = fileName
        If Not IsEmpty(syncAt) Then .Cells(r, M_SYNC).Value = syncAt
        .Cells(r, M_RESULT).Value = result
        .Cells(r, M_ROWS).Value = rows
        If Not IsEmpty(lastUpd) Then
            If CDbl(lastUpd) > 0 Then .Cells(r, M_LASTUPD).Value = lastUpd
        End If
        If ver <> "" Then .Cells(r, M_VER).Value = ver
    End With
End Sub

Private Sub AddLog(ByVal t As Date, ByVal fileName As String, ByVal owner As String, ByVal result As String, _
                   ByVal rows As Long, ByVal msgText As String)
    AppendRow GetLO(TB_LOG), Array(t, fileName, owner, result, rows, msgText)
End Sub

' 로그는 최근 500줄만 유지
Private Sub TrimLog()
    Dim lo As ListObject, n As Long
    Set lo = GetLO(TB_LOG)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    n = lo.ListRows.Count
    If n > 500 Then lo.DataBodyRange.Resize(n - 500).Delete
End Sub

Private Function MemberSummaryText() As String
    Dim a As Variant, i As Long, s As String
    a = ReadTable(GetLO(TB_MEMBER))
    For i = 1 To RowCount(a)
        If Left$(CStr(Nz(a(i, M_RESULT))), 2) = "정상" Then
            s = s & a(i, M_NAME) & " √ " & a(i, M_ROWS) & "건    "
        Else
            s = s & vbLf & a(i, M_NAME) & " ! " & a(i, M_RESULT) & vbLf
        End If
    Next i
    MemberSummaryText = s
End Function

Private Function CountLate() As Long
    Dim a As Variant, i As Long
    a = ReadTable(GetLO(TB_ALL))
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, A_LATE))) = "Y" Then CountLate = CountLate + 1
    Next i
End Function

'--- 프로젝트 요청 수집 --------------------------------------
' reqAll: Array(owner, row(1..9: ReqCode, ProjectName, PlanStart, PlanEnd, Memo, ReqDate, Result, MappedTo, Reason))
Private Function CollectRequests(ByVal reqAll As Collection) As Long
    Dim lo As ListObject, it As Variant, rq As Variant, r As Long, code As String
    Set lo = GetLO(TB_REQ)
    For Each it In reqAll
        rq = it(1)
        code = CStr(Nz(rq(1)))
        If code <> "" And CStr(Nz(rq(7))) = "" Then     ' 팀원 쪽에서 아직 처리 결과가 없는 요청
            r = FindRow(lo, RQ_CODE, code)
            If r = 0 Then
                AppendRow lo, Array(code, it(0), rq(2), rq(3), rq(4), rq(5), CountTasksOf(code), rq(6), Empty, Empty, Empty, Empty)
                CollectRequests = CollectRequests + 1
            Else
                lo.DataBodyRange.Cells(r, RQ_TASKS).Value = CountTasksOf(code)
            End If
        End If
    Next it
End Function

Public Function CountTasksOf(ByVal pid As String) As Long
    Dim a As Variant, i As Long
    a = ReadTable(GetLO(TB_ALL))
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, T_PID))) = pid And CStr(Nz(a(i, T_STATUS))) <> ST_CANCEL Then CountTasksOf = CountTasksOf + 1
    Next i
End Function

'--- 참여자 / 자동 PM ----------------------------------------
Public Sub UpdateParticipants()
    Dim lo As ListObject, a As Variant, p As Variant, i As Long, j As Long, parts As Object, pid As String
    Dim names As String, cnt As Long, k As Variant, dd As Object
    Set lo = GetLO(TB_PROJ)
    p = ReadTable(lo)
    If IsEmpty(p) Then Exit Sub
    a = ReadTable(GetLO(TB_ALL))
    Set parts = CreateObject("Scripting.Dictionary")
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, T_STATUS))) <> ST_CANCEL Then
            pid = CStr(Nz(a(i, T_PID)))
            If Not parts.Exists(pid) Then parts.Add pid, CreateObject("Scripting.Dictionary")
            Set dd = parts(pid)
            dd(CStr(Nz(a(i, T_OWNER)))) = True
        End If
    Next i
    For i = 1 To UBound(p, 1)
        pid = CStr(Nz(p(i, PR_ID)))
        names = "": cnt = 0
        If parts.Exists(pid) Then
            For Each k In parts(pid).Keys
                names = names & IIf(names = "", "", ", ") & k
                cnt = cnt + 1
            Next k
        End If
        p(i, PR_PART) = names
        If CStr(Nz(p(i, PR_PM))) <> "" Then
            p(i, PR_AUTOPM) = p(i, PR_PM)
        ElseIf cnt = 1 Then
            p(i, PR_AUTOPM) = names
        ElseIf cnt >= 2 Then
            p(i, PR_AUTOPM) = SYM_DELAY & " 지정 필요"
        Else
            p(i, PR_AUTOPM) = Empty
        End If
    Next i
    lo.DataBodyRange.Resize(, PR_COLS).Value = p
End Sub

' 실제 PM (지정 PM > 1인 자동 PM > 임시 프로젝트 요청자)
Public Function EffectivePM(ByVal pid As String) As String
    Dim lo As ListObject, r As Long, v As String
    Set lo = GetLO(TB_PROJ)
    r = FindRow(lo, PR_ID, pid)
    If r > 0 Then
        v = CStr(Nz(lo.DataBodyRange.Cells(r, PR_PM).Value))
        If v = "" Then v = CStr(Nz(lo.DataBodyRange.Cells(r, PR_AUTOPM).Value))
        If Left$(v, 1) = SYM_DELAY Then v = ""
        EffectivePM = v
        Exit Function
    End If
    r = FindRow(GetLO(TB_REQ), RQ_CODE, pid)
    If r > 0 Then EffectivePM = CStr(Nz(GetLO(TB_REQ).DataBodyRange.Cells(r, RQ_OWNER).Value))
End Function

Public Function CountPMNeeded() As Long
    Dim p As Variant, i As Long
    p = ReadTable(GetLO(TB_PROJ))
    For i = 1 To RowCount(p)
        If Left$(CStr(Nz(p(i, PR_AUTOPM))), 1) = SYM_DELAY Then CountPMNeeded = CountPMNeeded + 1
    Next i
End Function

'--- PM 파일의 단계 반영 --------------------------------------
' 프로젝트마다 PM 파일의 tblMyPhase 수정일시가 Master보다 최신이면 교체. 반영한 프로젝트 수 반환
Private Function ApplyPMPhases(ByVal phByOwner As Object) As Long
    Dim lo As ListObject, cur As Variant, owner As Variant, phs As Variant, i As Long, pid As String
    Dim pmRows As Object, mMax As Object, yMax As Object, keep As Collection, key As Variant, it As Variant

    Set lo = GetLO(TB_PHASE)
    cur = ReadTable(lo)
    Set mMax = CreateObject("Scripting.Dictionary")
    For i = 1 To RowCount(cur)
        pid = CStr(Nz(cur(i, PH_PID)))
        If StampOf(cur(i, PH_MOD)) > NumOr(mMax(pid)) Then mMax(pid) = StampOf(cur(i, PH_MOD))
    Next i

    ' PM 본인 파일의 단계만 채택
    Set pmRows = CreateObject("Scripting.Dictionary")
    Set yMax = CreateObject("Scripting.Dictionary")
    For Each owner In phByOwner.Keys
        phs = phByOwner(owner)
        For i = 1 To RowCount(phs)
            pid = CStr(Nz(phs(i, PH_PID)))
            If pid <> "" Then
                If EffectivePM(pid) = owner Then
                    If Not pmRows.Exists(pid) Then pmRows.Add pid, New Collection
                    pmRows(pid).Add RowOf(phs, i)
                    If StampOf(phs(i, PH_MOD)) > NumOr(yMax(pid)) Then yMax(pid) = StampOf(phs(i, PH_MOD))
                End If
            End If
        Next i
    Next owner

    Set keep = New Collection
    For i = 1 To RowCount(cur)
        pid = CStr(Nz(cur(i, PH_PID)))
        If pmRows.Exists(pid) Then
            If NumOr(yMax(pid)) > NumOr(mMax(pid)) Then GoTo NextRow    ' PM 쪽이 최신 → 교체 대상
        End If
        keep.Add RowOf(cur, i)
NextRow:
    Next i
    For Each key In pmRows.Keys
        If NumOr(yMax(key)) > NumOr(mMax(key)) Then
            For Each it In pmRows(key)
                keep.Add it
            Next it
            ApplyPMPhases = ApplyPMPhases + 1
        End If
    Next key
    If ApplyPMPhases > 0 Then WriteTable lo, SortPhases(CollTo2D(keep, PH_COLS))
End Function

Public Function StampOf(ByVal v As Variant) As Double
    If Not IsDateVal(v) Then Exit Function
    If VarType(v) = vbDate Or IsNumeric(v) Then
        StampOf = CDbl(v)
    Else
        StampOf = CDbl(CDate(v))
    End If
End Function

' 프로젝트 → 순서 정렬
Public Function SortPhases(ByVal a As Variant) As Variant
    Dim i As Long, j As Long, k As Long, tmp As Variant
    If IsEmpty(a) Then Exit Function
    For i = 2 To UBound(a, 1)
        For j = i To 2 Step -1
            If PhaseKey(a, j) < PhaseKey(a, j - 1) Then
                For k = 1 To UBound(a, 2)
                    tmp = a(j, k): a(j, k) = a(j - 1, k): a(j - 1, k) = tmp
                Next k
            Else
                Exit For
            End If
        Next j
    Next i
    SortPhases = a
End Function

Private Function PhaseKey(ByVal a As Variant, ByVal i As Long) As String
    PhaseKey = CStr(Nz(a(i, PH_PID))) & "|" & Format$(NumOr(a(i, PH_ORDER)), "0000")
End Function

'--- PM의 단계 합치기 → CodeMap(Type=Phase) ------------------
Private Sub ApplyPhaseMerges(ByVal mapAll As Collection)
    Dim lo As ListObject, it As Variant, fromC As String, toC As String, a As Variant, i As Long, exists As Boolean
    Set lo = GetLO(TB_MAP)
    For Each it In mapAll
        fromC = CStr(Nz(it(1))): toC = CStr(Nz(it(2)))
        If fromC <> "" And toC <> "" Then
            exists = False
            a = ReadTable(lo)
            For i = 1 To RowCount(a)
                If CStr(a(i, 1)) = fromC And CStr(a(i, 3)) = "Phase" Then exists = True: Exit For
            Next i
            If Not exists Then AppendRow lo, Array(fromC, toC, "Phase", "통합", "PM 단계 합치기", Now)
            ReplaceInAll T_PHASE, fromC, toC
        End If
    Next it
End Sub

Public Sub ReplaceInAll(ByVal colIdx As Long, ByVal fromC As String, ByVal toC As String)
    Dim lo As ListObject, a As Variant, i As Long, changed As Boolean
    Set lo = GetLO(TB_ALL)
    a = ReadTable(lo)
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, colIdx))) = fromC Then a(i, colIdx) = toC: changed = True
    Next i
    If changed Then lo.DataBodyRange.Value = a
End Sub

'--- 팀장 코멘트 → TaskAll -----------------------------------
Public Sub FillComments()
    Dim lo As ListObject, a As Variant, c As Variant, d As Object, i As Long
    Set lo = GetLO(TB_ALL)
    a = ReadTable(lo)
    If IsEmpty(a) Then Exit Sub
    c = ReadTable(GetLO(TB_CMT))
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To RowCount(c)
        d(CStr(Nz(c(i, CO_ID)))) = c(i, CO_TEXT)
    Next i
    For i = 1 To UBound(a, 1)
        If d.Exists(CStr(a(i, T_ID))) Then a(i, T_COMMENT) = d(CStr(a(i, T_ID))) Else a(i, T_COMMENT) = Empty
    Next i
    lo.DataBodyRange.Value = a
End Sub

Public Sub RefreshViews()
    RenderDashboard
    RenderProjectView
    RenderTeamSchedule
    RefreshLists
End Sub

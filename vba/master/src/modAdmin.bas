Attribute VB_Name = "modAdmin"
'==============================================================
' 팀장 관리 기능
'  - 프로젝트 요청 처리 (승인 / 기존 통합 / 반려)
'  - 팀장 코멘트 저장
'  - ProjectList.xlsx 배포 (팀원 파일이 열릴 때 읽음)
'  - 팀 공통 보고일정 펼치기
'  - 드롭다운 목록
'==============================================================
Option Explicit

'--- [요청 처리 적용] 버튼 -----------------------------------
Public Sub ProcessRequests()
    Dim lo As ListObject, a As Variant, i As Long, code As String, act As String, target As String
    Dim reason As String, n As Long, errs As String, owner As String, pr As ListObject

    Set lo = GetLO(TB_REQ)
    Set pr = GetLO(TB_PROJ)
    a = ReadTable(lo)
    If IsEmpty(a) Then Msg "처리할 요청이 없습니다.": Exit Sub

    SpeedOn
    For i = 1 To UBound(a, 1)
        code = CStr(Nz(a(i, RQ_CODE)))
        act = Trim$(CStr(Nz(a(i, RQ_ACTION))))
        target = Trim$(CStr(Nz(a(i, RQ_TARGET))))
        reason = CStr(Nz(a(i, RQ_REASON)))
        owner = CStr(Nz(a(i, RQ_OWNER)))
        If code <> "" And act <> "" And Not IsDateVal(a(i, RQ_DONE)) Then
            Select Case act
                Case "승인"
                    If target = "" Then target = NextProjectID()
                    If FindRow(pr, PR_ID, target) > 0 Then
                        errs = errs & vbLf & code & ": 정식 코드 " & target & " 는 이미 있습니다. (기존 통합을 선택하거나 다른 코드 입력)"
                        GoTo NextReq
                    End If
                    AppendRow pr, Array(target, a(i, RQ_NAME), owner, owner, owner, a(i, RQ_START), a(i, RQ_END), Empty, "진행")
                    AddCodeMap code, target, "Project", "승인", ""
                    RenamePhasePrefix code, target
                    ReplaceInAll T_PID, code, target
                Case "기존 통합"
                    If FindRow(pr, PR_ID, target) = 0 Then
                        errs = errs & vbLf & code & ": 통합 대상 프로젝트 코드(" & target & ")를 정확히 입력하세요."
                        GoTo NextReq
                    End If
                    AddCodeMap code, target, "Project", "통합", ""
                    RemovePhasesOf code
                    ReplaceInAll T_PID, code, target
                Case "반려"
                    AddCodeMap code, "", "Project", "반려", reason
                    RemovePhasesOf code
                    target = ""
                Case Else
                    errs = errs & vbLf & code & ": 처리 방법은 승인 / 기존 통합 / 반려 중 하나여야 합니다."
                    GoTo NextReq
            End Select
            lo.DataBodyRange.Cells(i, RQ_TARGET).Value = target
            lo.DataBodyRange.Cells(i, RQ_DONE).Value = Now
            n = n + 1
        End If
NextReq:
    Next i
    SpeedOff

    If n > 0 Then
        UpdateParticipants
        RecalcTaskAll
        ExportProjectList True
        RefreshViews
    End If
    Msg "요청 " & n & "건을 처리했습니다." & IIf(n > 0, vbLf & "팀원이 파일을 열면 코드가 자동으로 바뀝니다.", "") & _
        IIf(errs <> "", vbLf & vbLf & "처리하지 못한 요청:" & errs, "")
End Sub

' 다음 정식 코드: P + 3자리 (Setting의 ProjectPrefix)
Public Function NextProjectID() As String
    Dim a As Variant, i As Long, pre As String, mx As Long, tail As String
    pre = CStr(CfgGet("ProjectPrefix", "P"))
    a = ReadTable(GetLO(TB_PROJ))
    For i = 1 To RowCount(a)
        If Left$(CStr(Nz(a(i, PR_ID))), Len(pre)) = pre Then
            tail = Mid$(CStr(a(i, PR_ID)), Len(pre) + 1)
            If IsNumeric(tail) Then mx = IIf(CLng(tail) > mx, CLng(tail), mx)
        End If
    Next i
    NextProjectID = pre & Format$(mx + 1, "000")
End Function

Private Sub AddCodeMap(ByVal fromC As String, ByVal toC As String, ByVal tp As String, ByVal res As String, ByVal reason As String)
    AppendRow GetLO(TB_MAP), Array(fromC, toC, tp, res, reason, Now)
End Sub

' 임시 프로젝트 단계 → 정식 코드로 이름 변경
Private Sub RenamePhasePrefix(ByVal fromPid As String, ByVal toPid As String)
    Dim lo As ListObject, a As Variant, i As Long, id As String, changed As Boolean
    Set lo = GetLO(TB_PHASE)
    a = ReadTable(lo)
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, PH_PID))) = fromPid Then
            id = CStr(Nz(a(i, PH_ID)))
            If Left$(id, Len(fromPid) + 1) = fromPid & "-" Then a(i, PH_ID) = toPid & Mid$(id, Len(fromPid) + 1)
            a(i, PH_PID) = toPid
            changed = True
        End If
    Next i
    If changed Then lo.DataBodyRange.Value = a
End Sub

Private Sub RemovePhasesOf(ByVal pid As String)
    Dim lo As ListObject, a As Variant, i As Long, keep As Collection
    Set lo = GetLO(TB_PHASE)
    a = ReadTable(lo)
    If IsEmpty(a) Then Exit Sub
    Set keep = New Collection
    For i = 1 To UBound(a, 1)
        If CStr(Nz(a(i, PH_PID))) <> pid Then keep.Add RowOf(a, i)
    Next i
    If keep.Count < UBound(a, 1) Then WriteTable lo, CollTo2D(keep, PH_COLS)
End Sub

'--- 팀장 코멘트 --------------------------------------------
Public Sub UpsertComment(ByVal taskId As String, ByVal text As String)
    Dim lo As ListObject, r As Long, a As Variant, ra As Long
    If taskId = "" Then Exit Sub
    Set lo = GetLO(TB_CMT)
    r = FindRow(lo, CO_ID, taskId)
    If Trim$(text) = "" Then
        If r > 0 Then lo.ListRows(r).Delete
    Else
        ra = FindRow(GetLO(TB_ALL), T_ID, taskId)
        If r = 0 Then r = AppendRow(lo, Array(taskId))
        lo.DataBodyRange.Cells(r, CO_TEXT).Value = text
        lo.DataBodyRange.Cells(r, CO_DATE).Value = Now
        If ra > 0 Then
            lo.DataBodyRange.Cells(r, CO_OWNER).Value = GetLO(TB_ALL).DataBodyRange.Cells(ra, T_OWNER).Value
            lo.DataBodyRange.Cells(r, CO_TASK).Value = GetLO(TB_ALL).DataBodyRange.Cells(ra, T_NAME).Value
        End If
    End If
    ra = FindRow(GetLO(TB_ALL), T_ID, taskId)
    If ra > 0 Then GetLO(TB_ALL).DataBodyRange.Cells(ra, T_COMMENT).Value = IIf(Trim$(text) = "", Empty, text)
    CfgSet "CommentDirty", "Y"
End Sub

' [코멘트 보내기] 버튼 : 동기화 없이 ProjectList만 다시 배포
Public Sub PublishButton()
    ExportProjectList False
End Sub

'==============================================================
' ProjectList.xlsx 배포
'==============================================================
Public Sub ExportProjectList(Optional ByVal silent As Boolean = True)
    Dim wb As Workbook, path As String, p As Variant, i As Long, coll As Collection, lo As ListObject
    Dim pm As String, col As Long, cell As Range

    path = ExportFolder() & "\ProjectList.xlsx"
    On Error GoTo EH
    SpeedOn
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)

    ' Projects
    Set lo = GetLO(TB_PROJ)
    Set coll = New Collection
    p = ReadTable(lo)
    For i = 1 To RowCount(p)
        If CStr(Nz(p(i, PR_ID))) <> "" Then
            pm = CStr(Nz(p(i, PR_PM)))
            If pm = "" Then pm = CStr(Nz(p(i, PR_AUTOPM)))
            If Left$(pm, 1) = SYM_DELAY Then pm = ""
            Set cell = lo.DataBodyRange.Cells(i, PR_COLOR)
            If cell.Interior.ColorIndex = xlNone Then col = PaletteColor(CStr(p(i, PR_ID))) Else col = cell.Interior.Color
            coll.Add Array(p(i, PR_ID), p(i, PR_NAME), pm, col, IIf(CStr(Nz(p(i, PR_STATUS))) = "", "진행", p(i, PR_STATUS)))
        End If
    Next i
    PutSheet wb, "Projects", Array("ProjectID", "ProjectName", "PM", "Color", "Status"), CollTo2D(coll, 5), True
    PutSheet wb, "Phases", Array("PhaseID", "ProjectID", "Order", "PhaseName", "PlanStart", "PlanEnd", "Author", "ModifiedAt"), _
             ReadTable(GetLO(TB_PHASE)), False
    PutSheet wb, "CodeMap", Array("FromCode", "ToCode", "Type", "Result", "Reason", "Date"), ReadTable(GetLO(TB_MAP)), False
    PutSheet wb, "Comments", Array("TaskID", "Comment", "Date"), FirstCols(ReadTable(GetLO(TB_CMT)), 3), False
    PutSheet wb, "Events", Array("Date", "Time", "Title", "Target"), CollTo2D(ExpandEvents(), 4), False
    PutSheet wb, "Holidays", Array("Date", "Name"), ReadTable(GetLO(TB_HOL)), False

    wb.SaveAs Filename:=path, FileFormat:=51
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.DisplayAlerts = True
    CfgSet "LastExport", Now
    CfgSet "CommentDirty", ""
    SpeedOff
    If Not silent Then Msg "ProjectList.xlsx 를 배포했습니다." & vbLf & path & vbLf & "팀원이 파일을 열면 반영됩니다."
    Exit Sub
EH:
    Application.DisplayAlerts = True
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    SpeedReset
    Msg "ProjectList.xlsx 저장 실패: " & Err.Description & vbLf & path & vbLf & _
        "(누군가 파일을 열어 두었는지 확인 후 [ProjectList 배포]를 다시 누르세요)", vbExclamation
End Sub

Private Sub PutSheet(ByVal wb As Workbook, ByVal nm As String, ByVal headers As Variant, ByVal data As Variant, ByVal first As Boolean)
    Dim ws As Worksheet, n As Long
    If first Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    End If
    ws.Name = nm
    n = UBound(headers) - LBound(headers) + 1
    ws.Range("A1").Resize(1, n).Value = headers
    If Not IsEmpty(data) Then
        ws.Range("A2").Resize(UBound(data, 1), n).Value = FirstCols(data, n)
    End If
End Sub

Private Function FirstCols(ByVal data As Variant, ByVal n As Long) As Variant
    Dim r() As Variant, i As Long, j As Long
    If IsEmpty(data) Then Exit Function
    If UBound(data, 2) = n Then FirstCols = data: Exit Function
    ReDim r(1 To UBound(data, 1), 1 To n)
    For i = 1 To UBound(data, 1)
        For j = 1 To n
            If j <= UBound(data, 2) Then r(i, j) = data(i, j)
        Next j
    Next i
    FirstCols = r
End Function

'==============================================================
' 팀 공통 보고일정 펼치기 (오늘 기준 -60일 ~ +180일)
'  tblTeamEvent: Title | Time | Target | Repeat(한번/매주/매월/매월마지막) | Date | Weekday(월~일) | DayOfMonth | From | To
'==============================================================
Public Function ExpandEvents() As Collection
    Dim a As Variant, i As Long, res As New Collection, d As Date, d0 As Date, d1 As Date
    Dim rep As String, wd As Long, dom As Long, fromD As Date, toD As Date
    a = ReadTable(GetLO(TB_EVENT))
    d0 = Date - 60: d1 = Date + 180
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, 1))) <> "" Then
            rep = CStr(Nz(a(i, 4)))
            fromD = IIf(IsDateVal(a(i, 8)), ToDate(a(i, 8)), d0)
            toD = IIf(IsDateVal(a(i, 9)), ToDate(a(i, 9)), d1)
            If fromD < d0 Then fromD = d0
            If toD > d1 Then toD = d1
            wd = WeekdayNo(CStr(Nz(a(i, 6))))
            dom = CLng(NumOr(a(i, 7), 0))
            Select Case rep
                Case "매주"
                    If wd > 0 Then
                        For d = fromD To toD
                            If Weekday(d, vbMonday) = wd Then AddEvent res, d, a, i
                        Next d
                    End If
                Case "매월"
                    If dom > 0 Then
                        For d = fromD To toD
                            If Day(d) = dom Then AddEvent res, d, a, i
                        Next d
                    End If
                Case "매월마지막"
                    If wd > 0 Then
                        For d = fromD To toD
                            If Weekday(d, vbMonday) = wd And Month(d + 7) <> Month(d) Then AddEvent res, d, a, i
                        Next d
                    End If
                Case Else   ' 한번
                    If IsDateVal(a(i, 5)) Then AddEvent res, ToDate(a(i, 5)), a, i
            End Select
        End If
    Next i
    Set ExpandEvents = res
End Function

Private Sub AddEvent(ByVal res As Collection, ByVal d As Date, ByVal a As Variant, ByVal i As Long)
    res.Add Array(d, CStr(Nz(a(i, 2))), CStr(Nz(a(i, 1))), IIf(CStr(Nz(a(i, 3))) = "", FILTER_ALL, a(i, 3)))
End Sub

Private Function WeekdayNo(ByVal s As String) As Long
    WeekdayNo = InStr("월화수목금토일", Left$(Trim$(s), 1))
End Function

'==============================================================
' 드롭다운 목록 (Setting 시트 AH~AM 열 + 이름 정의)
'==============================================================
Public Sub RefreshLists()
    Dim ws As Worksheet, a As Variant, i As Long, projs As Collection, ids As Collection, mem As Collection
    Set ws = GetWS(SH_SET)
    If ws Is Nothing Then Exit Sub
    SpeedOn
    SetList ws, "AH", "lstRepeat", ToColl(Array("한번", "매주", "매월", "매월마지막"))
    SetList ws, "AI", "lstWeekday", ToColl(Array("월", "화", "수", "목", "금", "토", "일"))
    SetList ws, "AJ", "lstAction", ToColl(Array("승인", "기존 통합", "반려"))

    Set projs = New Collection: Set ids = New Collection: Set mem = New Collection
    a = ReadTable(GetLO(TB_PROJ))
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, PR_ID))) <> "" Then
            projs.Add CStr(a(i, PR_ID)) & " " & CStr(Nz(a(i, PR_NAME)))
            ids.Add CStr(a(i, PR_ID))
        End If
    Next i
    If projs.Count = 0 Then projs.Add "(프로젝트 없음)": ids.Add "-"
    a = ReadTable(GetLO(TB_MEMBER))
    For i = 1 To RowCount(a)
        If CStr(Nz(a(i, M_NAME))) <> "" Then mem.Add CStr(a(i, M_NAME))
    Next i
    If mem.Count = 0 Then mem.Add "-"
    SetList ws, "AK", "lstProjectAll", projs
    SetList ws, "AL", "lstProjectID", ids
    SetList ws, "AM", "lstMember", mem
    SpeedOff
End Sub

Private Function ToColl(ByVal arr As Variant) As Collection
    Dim c As New Collection, v As Variant
    For Each v In arr: c.Add v: Next v
    Set ToColl = c
End Function

Private Sub SetList(ByVal ws As Worksheet, ByVal colL As String, ByVal nm As String, ByVal items As Collection)
    Dim i As Long
    ws.Range(colL & "4:" & colL & "1000").ClearContents
    ws.Range(colL & "3").Value = nm
    For i = 1 To items.Count
        ws.Range(colL & (3 + i)).Value = items(i)
    Next i
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & ws.Name & "'!$" & colL & "$4:$" & colL & "$" & (3 + items.Count)
End Sub

' ProjectMaster의 tblPhase를 팀장이 직접 고치면 수정일시·작성자 기록 (팀장 수정이 PM 파일보다 우선)
Public Sub OnPhaseEdited(ByVal target As Range)
    Dim lo As ListObject, rng As Range, rw As Range, idx As Long
    Set lo = GetLO(TB_PHASE)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rng = Intersect(target, lo.DataBodyRange)
    If rng Is Nothing Then Exit Sub
    Application.EnableEvents = False
    For Each rw In rng.Rows
        idx = rw.Row - lo.DataBodyRange.Row + 1
        If CStr(Nz(lo.DataBodyRange.Cells(idx, PH_PID).Value)) <> "" Then
            If CStr(Nz(lo.DataBodyRange.Cells(idx, PH_ID).Value)) = "" Then
                lo.DataBodyRange.Cells(idx, PH_ID).Value = NewPhaseID(CStr(lo.DataBodyRange.Cells(idx, PH_PID).Value))
            End If
            lo.DataBodyRange.Cells(idx, PH_AUTHOR).Value = "팀장"
            lo.DataBodyRange.Cells(idx, PH_MOD).Value = Now
        End If
    Next rw
    Application.EnableEvents = True
End Sub

Private Function NewPhaseID(ByVal pid As String) As String
    Dim a As Variant, i As Long, mx As Long, id As String, tail As String
    a = ReadTable(GetLO(TB_PHASE))
    For i = 1 To RowCount(a)
        id = CStr(Nz(a(i, PH_ID)))
        If Left$(id, Len(pid) + 1) = pid & "-" Then
            tail = Mid$(id, Len(pid) + 2)
            If IsNumeric(tail) Then mx = IIf(CLng(tail) > mx, CLng(tail), mx)
        End If
    Next i
    NewPhaseID = pid & "-" & Format$(mx + 1, "00")
End Function

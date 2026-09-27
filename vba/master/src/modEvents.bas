Attribute VB_Name = "modEvents"
'==============================================================
' ThisWorkbook 이벤트에서 호출하는 처리
'==============================================================
Option Explicit

Public Sub OnOpen()
    If Not IsReady() Then Exit Sub
    RefreshLists
    GetWS(SH_DASH).Activate
End Sub

' 저장할 때 코멘트가 바뀌었으면 ProjectList 자동 배포
Public Sub OnBeforeSave()
    If Not IsReady() Then Exit Sub
    If CStr(CfgGet("CommentDirty", "")) = "Y" Then ExportProjectList True
End Sub

Public Sub OnSheetChange(ByVal sh As Object, ByVal target As Range)
    Dim lo As ListObject, c As Range, r As Long
    If Not IsReady() Then Exit Sub
    On Error GoTo EH
    Application.EnableEvents = False
    Select Case sh.Name
        Case SH_ALL
            Set lo = GetLO(TB_ALL)
            If Not lo.DataBodyRange Is Nothing Then
                For Each c In target.Cells
                    If c.Column = T_COMMENT And c.Row > 1 Then
                        r = c.Row - lo.DataBodyRange.Row + 1
                        If r >= 1 And r <= lo.ListRows.Count Then UpsertComment CStr(lo.DataBodyRange.Cells(r, T_ID).Value), CStr(Nz(c.Value))
                    End If
                Next c
            End If
        Case SH_DASH
            DashChange target
        Case SH_PM
            OnPhaseEdited target
            Set lo = GetLO(TB_PROJ)
            If Not lo.DataBodyRange Is Nothing Then
                If Not Intersect(target, lo.DataBodyRange) Is Nothing Then
                    UpdateParticipants
                    RefreshLists
                End If
            End If
        Case SH_PV
            If Not Intersect(target, sh.Range(PV_FILTER)) Is Nothing Then RenderProjectView
        Case SH_TS
            If Not Intersect(target, sh.Range(TS_MONTH)) Is Nothing Then RenderTeamSchedule
    End Select
    Application.EnableEvents = True
    Exit Sub
EH:
    SpeedReset
End Sub

Public Sub OnDoubleClick(ByVal sh As Object, ByVal target As Range, ByRef cancel As Boolean)
    If Not IsReady() Then Exit Sub
    If sh.Name = SH_DASH Then DashDoubleClick target, cancel
End Sub

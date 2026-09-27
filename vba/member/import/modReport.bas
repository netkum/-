Attribute VB_Name = "modReport"
'==============================================================
' 주간보고 초안 (WeeklyReport 시트)
'  - A열: 텍스트 형식 (복사해서 메일/보고서에 붙여넣기)
'  - C~F열: 표 형식 (프로젝트 | 금주 실적 | 차주 계획 | 이슈)
'==============================================================
Option Explicit

Public Sub MakeWeeklyReport()
    Dim s As String, base As Date
    s = InputBox("보고 기준 주의 날짜를 입력하세요 (그 주 월~일 기준).", APP_TITLE, Format$(Date, "yyyy-mm-dd"))
    If s = "" Then Exit Sub
    base = ParseDate(s)
    If base = 0 Then Msg "날짜 형식이 올바르지 않습니다.", vbExclamation: Exit Sub
    BuildWeeklyReport WeekMonday(base)
    GetWS(SH_RPT).Activate
End Sub

Public Sub BuildWeeklyReport(ByVal wkStart As Date)
    Dim ws As Worksheet, t As Variant, nT As Long, i As Long, pd As Object
    Dim wkEnd As Date, nxS As Date, nxE As Date, pids As Collection, seen As Object
    Dim pid As Variant, thisW As String, nextW As String, issue As String, absence As String
    Dim r As Long, tr As Long, st As String, s As Date, e As Date, lines As Collection, ln As Variant

    Set ws = GetWS(SH_RPT)
    wkEnd = wkStart + 6: nxS = wkStart + 7: nxE = wkStart + 13
    t = LoadTasks(): nT = RowCount(t)
    Set pd = ProjectDict()

    ' 이번 주·다음 주에 걸친 프로젝트 목록 (등장 순서)
    Set pids = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To nT
        If IsValidTask(t, i) Then
            If CStr(t(i, T_STATUS)) <> ST_CANCEL And ToDate(t(i, T_START)) <= nxE And ToDate(t(i, T_DUE)) >= wkStart Then
                If IsAbsence(CStr(t(i, T_PID))) Then
                    If ToDate(t(i, T_START)) <= wkEnd Then
                        absence = absence & ", " & t(i, T_PID) & " " & Format$(t(i, T_START), "m/d") & _
                                  IIf(ToDate(t(i, T_DUE)) > ToDate(t(i, T_START)), "~" & Format$(t(i, T_DUE), "m/d"), "")
                    End If
                ElseIf Not seen.Exists(CStr(t(i, T_PID))) Then
                    seen(CStr(t(i, T_PID))) = True
                    pids.Add CStr(t(i, T_PID))
                End If
            End If
        End If
    Next i

    SpeedOn
    ws.Range("A3:F1000").Clear
    Set lines = New Collection
    lines.Add "■ " & Replace(WeekLabel(wkStart), vbLf, " ") & " 주간보고 (" & OwnerName() & ", " & _
              Format$(wkStart, "m/d") & "~" & Format$(wkStart + 4, "m/d") & ")"

    ' 표 머리글
    tr = 4
    ws.Range("C3").Value = "표 형식"
    ws.Range("C4:F4").Value = Array("프로젝트", "금주 실적", "차주 계획", "이슈")
    With ws.Range("C4:F4")
        .Font.Bold = True
        .Interior.Color = RGB(68, 84, 106)
        .Font.Color = RGB(255, 255, 255)
    End With

    For Each pid In pids
        thisW = "": nextW = "": issue = ""
        For i = 1 To nT
            If IsValidTask(t, i) Then
                If CStr(t(i, T_PID)) = pid And CStr(t(i, T_STATUS)) <> ST_CANCEL Then
                    st = CStr(t(i, T_STATUS))
                    s = ToDate(t(i, T_START)): e = ToDate(t(i, T_DUE))
                    ' 금주 실적
                    If s <= wkEnd And e >= wkStart Then
                        If st = ST_DONE Then
                            thisW = AddItem(thisW, t(i, T_NAME) & " 완료")
                        ElseIf s <= Date Or wkEnd < Date Then
                            thisW = AddItem(thisW, t(i, T_NAME) & " 진행(" & NumOr(t(i, T_PCT)) & "%)")
                        End If
                    ElseIf st = ST_DONE And IsDateVal(t(i, T_DONE)) Then
                        If ToDate(t(i, T_DONE)) >= wkStart And ToDate(t(i, T_DONE)) <= wkEnd Then thisW = AddItem(thisW, t(i, T_NAME) & " 완료")
                    End If
                    ' 차주 계획
                    If st <> ST_DONE And s <= nxE And e >= nxS Then
                        If e <= nxE Then
                            nextW = AddItem(nextW, t(i, T_NAME) & " 완료 예정(" & Format$(e, "m/d") & ")")
                        ElseIf s >= nxS Then
                            nextW = AddItem(nextW, t(i, T_NAME) & " 착수(" & Format$(s, "m/d") & "~)")
                        Else
                            nextW = AddItem(nextW, t(i, T_NAME) & " 계속 진행")
                        End If
                    End If
                    ' 이슈: 지연 또는 메모
                    If s <= wkEnd And e >= wkStart - 28 And st <> ST_DONE Then
                        If IsDelayed(t, i) Then
                            issue = AddItem(issue, t(i, T_NAME) & " " & DDayText(e) & " 지연" & _
                                    IIf(CStr(Nz(t(i, T_NOTE))) <> "", " - " & t(i, T_NOTE), ""))
                        ElseIf CStr(Nz(t(i, T_NOTE))) <> "" And s <= wkEnd And e >= wkStart Then
                            issue = AddItem(issue, t(i, T_NAME) & ": " & t(i, T_NOTE))
                        End If
                    End If
                End If
            End If
        Next i
        If thisW <> "" Or nextW <> "" Or issue <> "" Then
            lines.Add "[" & pid & " " & ProjectNameOf(CStr(pid), pd) & "]"
            lines.Add "  - 금주: " & IIf(thisW = "", "-", thisW)
            lines.Add "  - 차주: " & IIf(nextW = "", "-", nextW)
            If issue <> "" Then lines.Add "  - 이슈: " & issue
            tr = tr + 1
            ws.Cells(tr, 3).Value = pid & " " & ProjectNameOf(CStr(pid), pd)
            ws.Cells(tr, 4).Value = Replace(thisW, ", ", vbLf)
            ws.Cells(tr, 5).Value = Replace(nextW, ", ", vbLf)
            ws.Cells(tr, 6).Value = Replace(issue, ", ", vbLf)
        End If
    Next pid
    If absence <> "" Then lines.Add "[근태] " & Mid$(absence, 3)
    If pids.Count = 0 And absence = "" Then lines.Add "(해당 주에 등록된 업무가 없습니다)"

    ws.Range("A3").Value = "텍스트 형식 (A열 선택 → 복사)"
    ws.Range("A3").Font.Bold = True
    ws.Range("C3").Font.Bold = True
    r = 4
    For Each ln In lines
        ws.Cells(r, 1).Value = "'" & ln
        r = r + 1
    Next ln
    ws.Cells(4, 1).Font.Bold = True
    If tr > 4 Then
        With ws.Range(ws.Cells(4, 3), ws.Cells(tr, 6))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(191, 191, 191)
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    End If
    SpeedOff
End Sub

Private Function AddItem(ByVal s As String, ByVal item As String) As String
    If s = "" Then AddItem = item Else AddItem = s & ", " & item
End Function

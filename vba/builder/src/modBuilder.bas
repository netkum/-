Attribute VB_Name = "modBuilder"
'==============================================================
' 팀 스케줄러 만들기 (build.bat 대신 엑셀 안에서 실행)
'  1) 새 통합 문서 → Alt+F11 → 파일 > 파일 가져오기 → Scheduler\Builder.bas
'  2) Alt+F8 → BuildAll 실행 → Scheduler 폴더 경로 입력
'  필요: 보안 센터 > 매크로 설정 > "VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음"
'==============================================================
Option Explicit

Public Sub BuildAll()
    Dim root As String
    root = InputBox("Scheduler 폴더 경로를 입력하세요." & vbLf & "(vba, Master, Member, Template 폴더가 있는 곳)", _
                    "팀 스케줄러 만들기", "F:\claude\Scheduler")
    root = Trim$(root)
    If root = "" Then Exit Sub
    If Right$(root, 1) = "\" Then root = Left$(root, Len(root) - 1)
    If Dir(root & "\vba\member\modConfig.bas") = "" Or Dir(root & "\vba\master\modConfig.bas") = "" Then
        MsgBox "코드 파일을 찾을 수 없습니다:" & vbLf & root & "\vba\member, \vba\master" & vbLf & _
               "압축을 푼 폴더 경로를 확인하세요.", vbExclamation
        Exit Sub
    End If
    If Not CanAccessVBA() Then
        MsgBox "VBA 프로젝트 접근이 꺼져 있습니다." & vbLf & vbLf & _
               "파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정 >" & vbLf & _
               "'VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음' 체크 후 다시 실행하세요.", vbExclamation
        Exit Sub
    End If
    MakeDir root & "\Master"
    MakeDir root & "\Member"
    MakeDir root & "\Template"

    On Error GoTo EH
    BuildBook root, "member", Array("frmTask", "frmRequest", "frmPhase"), "SetupTemplateSilent", _
              root & "\Template\업무일정_Template.xlsm"
    BuildBook root, "master", Array(), "SetupMasterSilent", root & "\Master\연구팀_Master.xlsm"
    MsgBox "완료! 두 파일을 만들었습니다." & vbLf & vbLf & _
           root & "\Template\업무일정_Template.xlsm" & vbLf & _
           root & "\Master\연구팀_Master.xlsm" & vbLf & vbLf & _
           "템플릿을 Member 폴더에 '이름_업무일정.xlsm'으로 복사해서 사용하세요.", vbInformation
    Exit Sub
EH:
    Application.DisplayAlerts = True
    MsgBox "만드는 중 오류: " & Err.Description, vbCritical
End Sub

Private Function CanAccessVBA() As Boolean
    Dim n As Long
    On Error Resume Next
    n = ThisWorkbook.VBProject.VBComponents.Count
    CanAccessVBA = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub MakeDir(ByVal p As String)
    If Dir(p, vbDirectory) = "" Then MkDir p
End Sub

Private Sub BuildBook(ByVal root As String, ByVal pkg As String, ByVal forms As Variant, _
                      ByVal macro As String, ByVal savePath As String)
    Dim wb As Workbook, vbp As Object, c As Object, f As Variant, folder As String, fn As String
    Dim files As Collection

    folder = root & "\vba\" & pkg & "\"
    Set wb = Workbooks.Add
    Set vbp = wb.VBProject

    ' 1) 사용자 정의 폼 (Microsoft Forms 참조가 먼저 추가되도록)
    For Each f In forms
        Set c = vbp.VBComponents.Add(3)
        c.Name = CStr(f)
        SetCode c, folder & f & ".txt"
    Next f

    ' 2) 클래스 → 모듈
    Set files = New Collection
    fn = Dir(folder & "*.cls")
    Do While fn <> ""
        files.Add fn
        fn = Dir()
    Loop
    fn = Dir(folder & "*.bas")
    Do While fn <> ""
        files.Add fn
        fn = Dir()
    Loop
    For Each f In files
        vbp.VBComponents.Import folder & f
    Next f

    ' 3) ThisWorkbook
    SetCode WorkbookModule(wb), folder & "ThisWorkbook.txt"

    ' 4) 시트·표·버튼 구성
    Application.Run "'" & wb.Name & "'!" & macro

    ' 5) 저장 (52 = 매크로 사용 통합 문서)
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=savePath, FileFormat:=52
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
End Sub

Private Function WorkbookModule(ByVal wb As Workbook) As Object
    Dim c As Object, v As Variant
    If wb.CodeName <> "" Then
        Set WorkbookModule = wb.VBProject.VBComponents(wb.CodeName)
        Exit Function
    End If
    For Each c In wb.VBProject.VBComponents
        If c.Type = 100 Then
            On Error Resume Next
            v = c.Properties("IsAddin").Value
            If Err.Number = 0 Then
                On Error GoTo 0
                Set WorkbookModule = c
                Exit Function
            End If
            On Error GoTo 0
        End If
    Next c
End Function

Private Sub SetCode(ByVal comp As Object, ByVal path As String)
    Dim cm As Object, text As String
    Set cm = comp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    text = ReadUtf8(path)
    cm.AddFromString text
End Sub

' UTF-8 텍스트 파일 읽기 (폼 코드 .txt)
Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
    If Left$(ReadUtf8, 1) = ChrW(&HFEFF) Then ReadUtf8 = Mid$(ReadUtf8, 2)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
' 입력폼에서 코드로 만든 컨트롤의 이벤트를 폼으로 전달
'   폼에 Public Sub OnCtlClick(name) / OnCtlChange(name) 필요
'==============================================================
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public WithEvents Cbo As MSForms.ComboBox
Attribute Cbo.VB_VarHelpID = -1
Public WithEvents Lst As MSForms.ListBox
Attribute Lst.VB_VarHelpID = -1
Public Host As Object
Public CtlName As String

Private Sub Btn_Click()
    CallByName Host, "OnCtlClick", VbMethod, CtlName
End Sub

Private Sub Cbo_Change()
    CallByName Host, "OnCtlChange", VbMethod, CtlName
End Sub

Private Sub Lst_Click()
    CallByName Host, "OnCtlChange", VbMethod, CtlName
End Sub

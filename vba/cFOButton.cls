VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cFOButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Button As MSForms.CommandButton
Public Owner As Object
Public ActionName As String
Private Sub Button_Click()
    Owner.Act ActionName
End Sub
Public Sub DetachButton()
    Set Button = Nothing
    Set Owner = Nothing
End Sub

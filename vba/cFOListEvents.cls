VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cFOListEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents List As MSForms.ListBox
Public Owner As Object

Private Sub List_Change()
    Owner.SelectionChanged
End Sub

Public Sub DetachList()
    Set List = Nothing
    Set Owner = Nothing
End Sub

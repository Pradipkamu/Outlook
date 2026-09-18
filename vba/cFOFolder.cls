VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cFOFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Items As Outlook.Items
Private Busy As Boolean
Public Sub Attach(ByVal folder As Outlook.Folder)
    Set Items = folder.Items
End Sub
Public Sub Detach()
    Set Items = Nothing
End Sub
Private Sub Items_ItemAdd(ByVal Item As Object)
    Dim result As Object
    If Busy Then Exit Sub
    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub
    Busy = True
    On Error GoTo Done
    Set result = FOCall("setup", FOSnapshot(Item))
Done:
    Busy = False
End Sub

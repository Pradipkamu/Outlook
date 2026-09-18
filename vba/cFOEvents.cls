VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cFOEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents App As Outlook.Application
Private InSend As Boolean
Private WithEvents PulseItems As Outlook.Items
Private InPulse As Boolean

Public Sub Attach(ByVal applicationObject As Outlook.Application)
    Set App = applicationObject
    Set PulseItems = App.Session.GetDefaultFolder(olFolderTasks).Items
End Sub
Public Sub Detach()
    Set PulseItems = Nothing
    Set App = Nothing
End Sub

Private Sub App_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim result As Object, choice As VbMsgBoxResult, form As Organizer
    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub
    If InSend Then Cancel = True: Exit Sub
    InSend = True
    On Error GoTo Failed
    If FOProp(Item, "FOGenerated") = "1" Then
        Set result = FOCall("permit", FOXML("id", FOProp(Item, "FOID")) & FOXML("occ", FOProp(Item, "FOOccurrence")))
        Cancel = (FOValue(result, "allowed") <> "1")
        GoTo Cleanup
    End If
    If FOProp(Item, "FOID") <> "" Then
        Cancel = (MsgBox("This email is already linked to a follow-up. Send and retain its current schedule/status?", vbYesNo + vbQuestion) <> vbYes)
        GoTo Cleanup
    End If
    choice = MsgBox("Add this email to follow-up?" & vbCrLf & "Yes: configure follow-up" & vbCrLf & "No: send without follow-up" & vbCrLf & "Cancel: keep email open", vbYesNoCancel + vbQuestion)
    If choice = vbCancel Then Cancel = True: GoTo Cleanup
    If choice = vbNo Then GoTo Cleanup
    Cancel = True
    Item.Save
    Set result = FOCall("find", FOSnapshot(Item))
    If FOValue(result, "id") <> "" Then
        If MsgBox("This conversation already has an open follow-up. Send and retain that schedule? You can change it from the dashboard.", vbYesNo + vbQuestion) = vbYes Then
            FOPut Item, "FOID", FOValue(result, "id")
            Cancel = False
        End If
        GoTo Cleanup
    End If
    Set FOEditingMail = Item
    FOPending = True
    FOSavedID = ""
    Set result = FOCall("ui_busy", FOXML("busy", "1"))
    Set form = New Organizer
    form.EditMail
    form.Show vbModal
    Unload form
    If FOSavedID <> "" Then
        FOPut Item, "FOID", FOSavedID
        Cancel = False
    End If
Cleanup:
    On Error Resume Next
    If FOPending Then Set result = FOCall("ui_busy", FOXML("busy", "0"))
    FOPending = False
    InSend = False
    Exit Sub
Failed:
    Cancel = True
    MsgBox "Send stopped: " & Err.Description & vbCrLf & "Check organizer, then try again.", vbExclamation
    Resume Cleanup
End Sub

Private Sub App_NewMailEx(ByVal EntryIDCollection As String)
    Dim ids As Variant, id As Variant, item As Object, result As Object, count As Long
    Dim started As Single
    On Error GoTo Done
    started = Timer
    ids = Split(EntryIDCollection, ",")
    For Each id In ids
        count = count + 1
        If count > 50 Then Exit For
        If Timer < started Or Timer - started > 2 Then Exit For
        Set item = App.Session.GetItemFromID(CStr(id))
        If TypeOf item Is Outlook.MailItem Then
            Set result = FOCall("event", FOSnapshot(item))
        End If
    Next id
    FOPoll
Done:
    ' The independent bounded scan catches events missed here.
End Sub

Private Sub PulseItems_ItemChange(ByVal Item As Object)
    HandlePulse Item
End Sub

Private Sub PulseItems_ItemAdd(ByVal Item As Object)
    HandlePulse Item
End Sub

Private Sub HandlePulse(ByVal Item As Object)
    If FOProp(Item, "FOBackgroundPulse") <> "1" Then Exit Sub
    If InPulse Then Exit Sub
    InPulse = True
    On Error GoTo Done
    ' No modal message boxes or task writes inside this COM-triggered callback.
    FOPoll True
Done:
    InPulse = False
End Sub

Attribute VB_Name = "FollowupOrganizer"
Option Explicit
Public FOEvents As cFOEvents
Public FOWindow As Organizer
Public FOEditingMail As Outlook.MailItem
Public FOPending As Boolean
Public FOSavedID As String
Public FOLastAlert As String
Private FOBusy As Boolean
Private FOFolderWatch As cFOFolder
Private FODailyMarker As String

Public Function FOHome() As String
    FOHome = "D:\FollowupOrganizer\data"
End Function

Public Function FOValue(ByVal node As Object, ByVal key As String) As String
    Dim n As Object
    Set n = node.SelectSingleNode(key)
    If Not n Is Nothing Then FOValue = n.Text
End Function

Private Function FOUTCToLocal(ByVal value As String) As Date
    FOUTCToLocal = DateSerial(CLng(Left$(value, 4)), CLng(Mid$(value, 6, 2)), CLng(Mid$(value, 9, 2))) + TimeSerial(CLng(Mid$(value, 12, 2)), CLng(Mid$(value, 15, 2)), 0) + TimeSerial(5, 30, 0)
End Function

Public Function FOXML(ByVal key As String, ByVal value As String) As String
    Dim doc As Object, n As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set n = doc.createElement(key)
    n.Text = value
    FOXML = n.XML
End Function

Public Function FOCall(ByVal action As String, Optional ByVal fields As String = "") As Object
    Dim req As Object, doc As Object, fso As Object, stream As Object, token As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.OpenTextFile(FOHome() & "\token.txt", 1)
    token = Trim$(stream.ReadAll)
    stream.Close
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.SetTimeouts 1000, 1000, 2000, 3000
    req.Open "POST", "http://127.0.0.1:8765/" & action, False
    req.SetRequestHeader "Content-Type", "application/xml; charset=utf-8"
    req.SetRequestHeader "X-FO-Token", token
    req.Send "<request>" & fields & "</request>"
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.setProperty "ProhibitDTD", True
    If Not doc.LoadXML(req.ResponseText) Then Err.Raise vbObjectError + 2100, , "Invalid organizer response."
    If req.Status <> 200 Then Err.Raise vbObjectError + 2101, , FOValue(doc.DocumentElement, "error")
    Set FOCall = doc.DocumentElement
End Function

Public Function FOProp(ByVal item As Object, ByVal key As String) As String
    Dim p As Outlook.UserProperty
    On Error GoTo Missing
    Set p = item.UserProperties.Find(key)
    If Not p Is Nothing Then FOProp = CStr(p.Value)
Missing:
End Function

Public Sub FOPut(ByVal item As Object, ByVal key As String, ByVal value As String)
    Dim p As Outlook.UserProperty
    Set p = item.UserProperties.Find(key)
    If p Is Nothing Then Set p = item.UserProperties.Add(key, olText, False)
    p.Value = value
End Sub

Public Function FOSMTP(ByVal mail As Outlook.MailItem) As String
    Dim ex As Outlook.ExchangeUser
    On Error GoTo Fallback
    If mail.SenderEmailType = "EX" Then
        Set ex = mail.Sender.GetExchangeUser
        If Not ex Is Nothing Then FOSMTP = ex.PrimarySmtpAddress
    Else
        FOSMTP = mail.SenderEmailAddress
    End If
    Exit Function
Fallback:
    FOSMTP = ""
End Function

Public Function FORecipientList(ByVal mail As Outlook.MailItem) As String
    Dim recipient As Outlook.Recipient, address As String, ex As Outlook.ExchangeUser, entry As Outlook.AddressEntry
    On Error GoTo Done
    For Each recipient In mail.Recipients
        If recipient.Type = olTo Then
            address = ""
            Set entry = recipient.AddressEntry
            If entry.Type = "EX" Then
                Set ex = entry.GetExchangeUser
                If Not ex Is Nothing Then address = ex.PrimarySmtpAddress
            Else
                address = entry.Address
            End If
            If address <> "" Then
                If FORecipientList <> "" Then FORecipientList = FORecipientList & ";"
                FORecipientList = FORecipientList & address
            End If
        End If
    Next recipient
Done:
End Function

Public Function FOIsOwn(ByVal mail As Outlook.MailItem) As Boolean
    Dim a As Outlook.Account, sender As String
    sender = LCase$(FOSMTP(mail))
    If sender = "" Then Exit Function
    For Each a In Application.Session.Accounts
        If LCase$(a.SmtpAddress) = sender Then FOIsOwn = True: Exit Function
    Next a
End Function

Private Function FOMapi(ByVal mail As Outlook.MailItem, ByVal tag As String) As String
    On Error GoTo Missing
    FOMapi = CStr(mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/" & tag))
Missing:
End Function

Public Function FOSnapshot(ByVal mail As Outlook.MailItem, Optional ByVal prefix As String = "") As String
    Dim result As String, sent As String
    sent = "0"
    If FOIsOwn(mail) Then sent = "1"
    result = FOXML(prefix & "entry", mail.EntryID)
    If mail.EntryID <> "" Then result = result & FOXML(prefix & "store", mail.Parent.StoreID)
    result = result & FOXML(prefix & "conv", mail.ConversationID)
    result = result & FOXML(prefix & "subject", mail.Subject)
    result = result & FOXML(prefix & "mid", FOMapi(mail, "0x1035001F"))
    result = result & FOXML(prefix & "refs", FOMapi(mail, "0x1039001F"))
    result = result & FOXML(prefix & "inreply", FOMapi(mail, "0x1042001F"))
    result = result & FOXML(prefix & "fid", FOProp(mail, "FOID"))
    result = result & FOXML(prefix & "occ", FOProp(mail, "FOOccurrence"))
    result = result & FOXML(prefix & "generated", FOProp(mail, "FOGenerated"))
    result = result & FOXML(prefix & "sent", sent)
    ' Windows local timestamps are interpreted as IST in this pilot.
    result = result & FOXML(prefix & "received", Format$(mail.ReceivedTime, "yyyy-mm-dd\Thh:nn:ss"))
    FOSnapshot = result
End Function

Public Function FOInboxSnapshot(ByVal mail As Outlook.MailItem, Optional ByVal prefix As String = "") As String
    ' Inbox candidates are received mail. Avoid sender-address inspection here so
    ' the approval loop does not create one Outlook security prompt per message.
    Dim result As String
    result = FOXML(prefix & "entry", mail.EntryID)
    If mail.EntryID <> "" Then result = result & FOXML(prefix & "store", mail.Parent.StoreID)
    result = result & FOXML(prefix & "conv", mail.ConversationID)
    result = result & FOXML(prefix & "subject", mail.Subject)
    result = result & FOXML(prefix & "mid", FOMapi(mail, "0x1035001F"))
    result = result & FOXML(prefix & "refs", FOMapi(mail, "0x1039001F"))
    result = result & FOXML(prefix & "inreply", FOMapi(mail, "0x1042001F"))
    result = result & FOXML(prefix & "fid", FOProp(mail, "FOID"))
    result = result & FOXML(prefix & "occ", FOProp(mail, "FOOccurrence"))
    result = result & FOXML(prefix & "generated", FOProp(mail, "FOGenerated"))
    result = result & FOXML(prefix & "sent", "0")
    result = result & FOXML(prefix & "received", Format$(mail.ReceivedTime, "yyyy-mm-dd\Thh:nn:ss"))
    FOInboxSnapshot = result
End Function

Public Function FOImportantSnapshot(ByVal mail As Outlook.MailItem, Optional ByVal prefix As String = "") As String
    Dim result As String, folderPath As String
    result = FOInboxSnapshot(mail, prefix)
    On Error Resume Next
    folderPath = mail.Parent.FolderPath
    On Error GoTo 0
    FOImportantSnapshot = result & FOXML(prefix & "folder", folderPath)
End Function

Public Function FOSelected() As Outlook.MailItem
    Dim item As Object
    If Not Application.ActiveInspector Is Nothing Then
        Set item = Application.ActiveInspector.CurrentItem
    ElseIf Not Application.ActiveExplorer Is Nothing Then
        If Application.ActiveExplorer.Selection.Count = 1 Then Set item = Application.ActiveExplorer.Selection.Item(1)
    End If
    If TypeOf item Is Outlook.MailItem Then Set FOSelected = item
End Function

Public Sub FollowupStart()
    On Error GoTo Failed
    If Not FOEvents Is Nothing Then FOEvents.Detach
    Set FOEvents = New cFOEvents
    FOEvents.Attach Application
    Dim reply As Object
    FODisableLegacyMonitor
    Set reply = FOCall("pulse")
    FOSetupFolder
    Exit Sub
Failed:
    ' The independent worker triggers another background event after connecting.
    Debug.Print "Follow-up service not ready: " & Err.Description
End Sub

Private Sub FOSetupFolder()
    Dim inbox As Outlook.Folder, folder As Outlook.Folder, child As Outlook.Folder, result As Object
    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    For Each child In inbox.Folders
        If child.Name = "Follow-up Mail" Then Set folder = child: Exit For
    Next child
    If folder Is Nothing Then Set folder = inbox.Folders.Add("Follow-up Mail", olFolderInbox)
    Set FOFolderWatch = New cFOFolder
    FOFolderWatch.Attach folder
    Set result = FOCall("folder", FOXML("entry", folder.EntryID) & FOXML("store", folder.StoreID))
End Sub

Public Sub FollowupDashboard()
    On Error GoTo Failed
    If FOEvents Is Nothing Then FollowupStart
    If FOWindow Is Nothing Then Set FOWindow = New Organizer
    FOWindow.ShowDashboard "Today"
    If Not FOWindow.Visible Then FOWindow.Show vbModeless
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub AddFollowup()
    On Error GoTo Failed
    Set FOEditingMail = FOSelected()
    If FOEditingMail Is Nothing Then MsgBox "Select one received or sent email.", vbInformation: Exit Sub
    FOPending = False
    FOEditingMail.Save
    If FOWindow Is Nothing Then Set FOWindow = New Organizer
    FOWindow.EditMail
    If Not FOWindow.Visible Then FOWindow.Show vbModeless
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub StopFollowupAutomation()
    Dim fso As Object, stream As Object
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FOHome()) Then fso.CreateFolder FOHome()
    Set stream = fso.CreateTextFile(FOHome() & "\STOP", True, False)
    stream.Write "User pressed Stop Automation. Review before resuming."
    stream.Close
    MsgBox "Automatic processing blocked. Already submitted emails cannot be recalled.", vbInformation
    Exit Sub
Failed:
    MsgBox "Could not write emergency stop: " & Err.Description, vbCritical
End Sub

Public Sub MonitorFollowupFolder()
    Dim folder As Outlook.Folder, result As Object
    On Error GoTo Failed
    Set folder = Application.Session.PickFolder
    If folder Is Nothing Then Exit Sub
    Set result = FOCall("folder", FOXML("entry", folder.EntryID) & FOXML("store", folder.StoreID))
    MsgBox "Folder added to reply reconciliation. Add rule destination folders individually; child folders are separate.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub FOPoll(Optional ByVal Background As Boolean = False)
    Dim result As Object, row As Object, signature As String, count As Long
    Dim preferences As Object, marker As String, moment As String, reviewMoment As String
    If FOBusy Then Exit Sub
    FOBusy = True
    On Error GoTo Failed
    Set result = FOCall("pulse")
    If FOFolderWatch Is Nothing Then FOSetupFolder
    Set result = FOCall("list")
    For Each row In result.SelectNodes("row")
        If FOValue(row, "status") = "Review" Then
            count = count + 1
            signature = signature & FOValue(row, "id") & FOValue(row, "version")
            reviewMoment = FOValue(row, "review_after")
            If reviewMoment <> "" Then
                If Now >= FOUTCToLocal(reviewMoment) Then signature = signature & "overdue" & Format$(Now, "yyyymmddhh")
            End If
        End If
    Next row
    Set preferences = FOCall("settings")
    moment = Format$(Now, "hh:nn")
    If FODailyMarker = "" Then
        Set result = FOCall("review_marker")
        FODailyMarker = FOValue(result, "marker")
    End If
    If moment >= FOValue(preferences, "morning") Then marker = Format$(Date, "yyyy-mm-dd") & " morning"
    If moment >= FOValue(preferences, "evening") Then marker = Format$(Date, "yyyy-mm-dd") & " evening"
    If InStr(1, FOValue(preferences, "weekdays"), CStr(Weekday(Date, vbMonday) - 1), vbBinaryCompare) = 0 Then marker = ""
    If InStr(1, FOValue(preferences, "holidays"), Format$(Date, "yyyy-mm-dd"), vbBinaryCompare) > 0 Then marker = ""
    If marker <> "" And marker <> FODailyMarker Then
        FODailyMarker = marker
        Set result = FOCall("review_marker", FOXML("marker", marker))
        If Background Then
            FOBackgroundAttention "Today"
        Else
            If MsgBox("Time to review today's follow-ups, deadlines and sending problems. Open Today?", vbYesNo + vbInformation) = vbYes Then FollowupDashboard
        End If
    End If
    If count > 0 And signature <> FOLastAlert Then
        FOLastAlert = signature
        If Background Then
            FOBackgroundAttention "Review Replies"
        ElseIf MsgBox(CStr(count) & " follow-up(s) have replies. Reminders are paused. View now?", vbYesNo + vbInformation) = vbYes Then
            FollowupDashboard
            FOWindow.ShowDashboard "Review Replies"
        End If
    End If
    FOBusy = False
    Exit Sub
Failed:
    FOBusy = False
    ' No recursive retries or repeated modal errors on a stopped service.
End Sub

Private Sub FOBackgroundAttention(ByVal view As String)
    On Error GoTo Failed
    If FOWindow Is Nothing Then Set FOWindow = New Organizer
    FOWindow.ShowDashboard view
    If Not FOWindow.Visible Then FOWindow.Show vbModeless
    Exit Sub
Failed:
    Debug.Print "Background follow-up display: " & Err.Description
End Sub

Public Sub FODisableLegacyMonitor()
    Dim tasks As Outlook.Items, item As Object, i As Long
    Dim reminder As Outlook.Reminder
    Set tasks = Application.Session.GetDefaultFolder(olFolderTasks).Items
    Set tasks = tasks.Restrict("[Subject] = 'Follow-up Organizer monitor'")
    If tasks.Count > 100 Then Err.Raise vbObjectError + 2140, , "Too many legacy monitor tasks. Review Tasks before continuing."
    For i = tasks.Count To 1 Step -1
        Set item = tasks.Item(i)
        If TypeOf item Is Outlook.TaskItem Then
            If FOProp(item, "FOClock") = "1" Then
                item.ReminderSet = False
                item.Save
            End If
        End If
    Next i
    ' Dismiss only active reminders positively identified as this old monitor.
    For i = Application.Reminders.Count To 1 Step -1
        Set reminder = Application.Reminders.Item(i)
        If FOProp(reminder.Item, "FOClock") = "1" Then reminder.Dismiss
    Next i
End Sub

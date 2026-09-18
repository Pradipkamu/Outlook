Attribute VB_Name = "FOExcel"
Option Explicit
Private Const HEADERS As String = "id|version|subject|to|cc|account|priority|next_at|unit|every|maximum|mode|message|later_message|deadline|stop_date|note|status|count|review_after|error|anchor_entry|anchor_store"

Public Sub ExportFollowupsToExcel()
    Dim excel As Object, book As Object, sheet As Object, r As Object, row As Object
    Dim headers As Variant, i As Long, n As Long, path As String
    On Error GoTo Failed
    Set r = FOCall("list")
    Set excel = CreateObject("Excel.Application")
    Set book = excel.Workbooks.Add
    Set sheet = book.Worksheets(1)
    sheet.Name = "Reminders"
    headers = Split(HEADERS, "|")
    For i = 0 To UBound(headers): sheet.Cells(1, i + 1).Value = headers(i): Next i
    sheet.Cells.NumberFormat = "@"
    n = 2
    For Each row In r.SelectNodes("row")
        If Left$(FOValue(row, "id"), 6) <> "setup:" Then
            For i = 0 To UBound(headers)
                sheet.Cells(n, i + 1).Value2 = FOValue(row, CStr(headers(i)))
            Next i
            n = n + 1
        End If
    Next row
    sheet.Rows(1).Font.Bold = True
    sheet.Rows(1).Interior.Color = RGB(15, 108, 189)
    sheet.Rows(1).Font.Color = RGB(255, 255, 255)
    sheet.Columns.ColumnWidth = 22
    sheet.Columns(3).ColumnWidth = 45
    sheet.Columns(13).ColumnWidth = 60
    sheet.Columns(14).ColumnWidth = 60
    sheet.Range("A1:W" & CStr(n)).AutoFilter
    sheet.Activate
    excel.ActiveWindow.SplitRow = 1
    excel.ActiveWindow.FreezePanes = True
    Set sheet = book.Worksheets.Add(, book.Worksheets(book.Worksheets.Count))
    sheet.Name = "Instructions"
    sheet.Cells(1, 1).Value = "Follow-up Organizer configuration export"
    sheet.Cells(2, 1).Value = "Pause automation before editing. Close Excel before importing. Import validates all rows before applying any."
    sheet.Cells(3, 1).Value = "Keep id/version unchanged. Existing rows only: add new follow-ups from Outlook so messages are linked."
    sheet.Cells(4, 1).Value = "Set status to Closed to remove from reminders, or Paused to stop temporarily. Deleting a row does NOT close a follow-up."
    sheet.Cells(5, 1).Value = "Other statuses are preserved. next_at is ISO date/time with offset, e.g. 2026-09-15T09:30:00+05:30."
    sheet.Cells(6, 1).Value = "Review replies through Outlook. Import never resumes paused/review/closed work; use Edit / Continue in Outlook."
    sheet.Cells(7, 1).Value = "Edit messages, recipients, frequency, maximum, priority and future schedule. No formulas are accepted."
    sheet.Columns(1).ColumnWidth = 110
    sheet.UsedRange.WrapText = True
    sheet.UsedRange.Rows.AutoFit
    path = FOHome() & "\Followups-" & Format$(Now, "yyyymmdd-hhnnss") & ".xlsx"
    book.SaveAs path, 51
    excel.Visible = True
    MsgBox "Excel export saved: " & path, vbInformation
    Set book = Nothing: Set excel = Nothing
    Exit Sub
Failed:
    MsgBox "Excel export failed: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If Not excel Is Nothing Then excel.Quit
End Sub

Public Sub ImportFollowupsFromExcel()
    Dim excel As Object, book As Object, sheet As Object, picker As Object, path As String
    Dim headers As Variant, i As Long, n As Long, last As Long, batch As String, fields As String, result As Object
    On Error GoTo Failed
    Set excel = CreateObject("Excel.Application")
    excel.AutomationSecurity = 3
    path = CStr(excel.GetOpenFilename("Excel workbooks (*.xlsx),*.xlsx", , "Import organizer configuration"))
    If path = "False" Then excel.Quit: Exit Sub
    Set book = excel.Workbooks.Open(path, 0, True)
    Set sheet = book.Worksheets("Reminders")
    headers = Split(HEADERS, "|")
    For i = 0 To UBound(headers)
        If CStr(sheet.Cells(1, i + 1).Value2) <> headers(i) Then Err.Raise vbObjectError + 2200, , "Columns changed. Export a fresh workbook."
    Next i
    last = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
    If last > 1001 Then Err.Raise vbObjectError + 2201, , "Import is limited to 1,000 follow-ups per workbook."
    For n = 2 To last
        fields = ""
        If CStr(sheet.Cells(n, 1).Value2) <> "" Then
            For i = 0 To UBound(headers)
                If sheet.Cells(n, i + 1).HasFormula Then Err.Raise vbObjectError + 2202, , "Remove formulas before import."
                fields = fields & FOXML(CStr(headers(i)), CStr(sheet.Cells(n, i + 1).Value2))
            Next i
            batch = batch & "<row>" & fields & "</row>"
        End If
    Next n
    book.Close False: excel.Quit
    Set book = Nothing: Set excel = Nothing
    Set result = FOCall("import", FOXML("batch", "<rows>" & batch & "</rows>"))
    MsgBox "Imported " & FOValue(result, "count") & " rows. Automation remains stopped; review before resuming.", vbInformation
    Exit Sub
Failed:
    MsgBox "Import not completed: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If Not excel Is Nothing Then excel.Quit
End Sub

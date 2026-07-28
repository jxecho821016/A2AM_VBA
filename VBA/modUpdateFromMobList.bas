Option Explicit

' Pile Records <- Project Mob List import.
'
' Lives in Pile_record_MUST_COMPLETE.xlsm. Assign UpdateFromMobList to
' the "Update" button on the Summary sheet. Clicking it:
'   1. Reads the OneDrive/SharePoint link to the Project Mob List
'      workbook from Summary!E7 (the cell's hyperlink if it has one,
'      otherwise the cell text).
'   2. Opens that workbook read-only in the background.
'   3. Copies every Schedule row with a Pile ID - columns A:M from
'      row 11 down - into Summary starting at C10 (so Schedule A lands
'      in C, B in D, ... M in O). Values only; Summary formatting is
'      kept and formulas in the source (e.g. Toe) arrive as numbers.
'   4. Clears the previous C10:O import area first, so deleted
'      schedule rows do not linger, then closes the source workbook
'      without saving.
'
' Columns right of O (Pile Cage, Casing, Concrete Time) are never
' touched - they hold data entered in this workbook.

Private Const SUMMARY_SHEET As String = "Summary"
Private Const LINK_CELL As String = "E7"
Private Const SOURCE_SHEET As String = "Schedule"
Private Const SOURCE_FIRST_DATA_ROW As Long = 11
Private Const SOURCE_COLUMNS As Long = 13       ' A:M
Private Const DEST_FIRST_CELL As String = "C10"
Private Const CLEAR_LAST_ROW As Long = 5000

Public Sub UpdateFromMobList()
    Dim summarySheet As Worksheet
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim linkAddress As String
    Dim lastRow As Long
    Dim rowCount As Long
    Dim previousScreenUpdating As Boolean
    Dim previousAlerts As Boolean
    Dim errorMessage As String

    On Error GoTo CleanFail

    previousScreenUpdating = Application.ScreenUpdating
    previousAlerts = Application.DisplayAlerts

    Set summarySheet = FindSheet(ThisWorkbook, SUMMARY_SHEET)
    If summarySheet Is Nothing Then
        MsgBox "This workbook has no sheet called """ & SUMMARY_SHEET & _
               """." & vbCrLf & "Sheets found: " & _
               SheetNameList(ThisWorkbook) & vbCrLf & vbCrLf & _
               "Rename the tab, or change the SUMMARY_SHEET constant " & _
               "at the top of modUpdateFromMobList.", vbExclamation, _
               "Update from mob list"
        Exit Sub
    End If

    linkAddress = MobListLink(summarySheet)
    If Len(linkAddress) = 0 Then
        MsgBox "Please put the OneDrive link to the Project Mob List " & _
               "workbook in cell " & LINK_CELL & " of the " & _
               SUMMARY_SHEET & " sheet first.", vbExclamation, _
               "Update from mob list"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set sourceBook = OpenMobList(linkAddress, sourceSheet)
    If sourceBook Is Nothing Then
        Application.DisplayAlerts = previousAlerts
        Application.ScreenUpdating = previousScreenUpdating
        MsgBox "Excel could not open the Project Mob List from the " & _
               "link in " & LINK_CELL & ":" & vbCrLf & linkAddress & _
               vbCrLf & vbCrLf & _
               "A SharePoint *share* link (the "":x:/s/..."" kind from " & _
               """Copy link"") opens a redirect page, not the file. " & _
               "Use the file's direct path instead:" & vbCrLf & _
               "1. Open the mob list workbook in Excel." & vbCrLf & _
               "2. Go to File > Info and click ""Copy path""." & vbCrLf & _
               "3. Paste that into " & LINK_CELL & " (the macro removes " & _
               "the ?web=1 ending automatically) and click Update again.", _
               vbExclamation, "Update from mob list"
        Exit Sub
    End If

    lastRow = sourceSheet.Cells( _
        sourceSheet.Rows.Count, 1).End(xlUp).Row

    ' Old import cleared even when the schedule is now empty.
    summarySheet.Range(DEST_FIRST_CELL) _
        .Resize(CLEAR_LAST_ROW, SOURCE_COLUMNS).ClearContents

    If lastRow >= SOURCE_FIRST_DATA_ROW Then
        rowCount = lastRow - SOURCE_FIRST_DATA_ROW + 1
        summarySheet.Range(DEST_FIRST_CELL) _
            .Resize(rowCount, SOURCE_COLUMNS).Value = _
            sourceSheet.Cells(SOURCE_FIRST_DATA_ROW, 1) _
                .Resize(rowCount, SOURCE_COLUMNS).Value
    End If

    sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing

    Application.DisplayAlerts = previousAlerts
    Application.ScreenUpdating = previousScreenUpdating

    If rowCount = 0 Then
        MsgBox "The mob list Schedule tab has no pile rows yet - the " & _
               "import area was cleared.", vbInformation, _
               "Update from mob list"
    Else
        MsgBox rowCount & " schedule row(s) copied from the Project " & _
               "Mob List.", vbInformation, "Update from mob list"
    End If
    Exit Sub

CleanFail:
    errorMessage = Err.Description
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    Application.DisplayAlerts = previousAlerts
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not update from the Project Mob List." & vbCrLf & vbCrLf & _
           errorMessage & vbCrLf & vbCrLf & _
           "Check that the link in " & LINK_CELL & " points to the mob " & _
           "list workbook and that you are signed in to OneDrive.", _
           vbExclamation, "Update from mob list"
End Sub

' Opens the mob list workbook and hands back the Schedule sheet via
' scheduleSheet. A SharePoint share link served to Workbooks.Open
' returns a junk single-sheet workbook named after the link token, so
' every candidate is validated by looking for the Schedule tab; for
' share links a "?download=1" variant is tried first, which makes
' SharePoint serve the real file. Returns Nothing when no candidate
' yields a workbook containing the Schedule sheet.
Private Function OpenMobList(ByVal linkAddress As String, _
                             ByRef scheduleSheet As Worksheet) As Workbook
    Dim candidates As Variant
    Dim i As Long
    Dim book As Workbook

    If InStr(1, linkAddress, "/:", vbTextCompare) > 0 Then
        candidates = Array(linkAddress & "?download=1", linkAddress)
    Else
        candidates = Array(linkAddress)
    End If

    For i = LBound(candidates) To UBound(candidates)
        Set book = Nothing
        On Error Resume Next
        Set book = Workbooks.Open( _
            Filename:=CStr(candidates(i)), ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo 0

        If Not book Is Nothing Then
            Set scheduleSheet = FindSheet(book, SOURCE_SHEET)
            If Not scheduleSheet Is Nothing Then
                Set OpenMobList = book
                Exit Function
            End If
            ' Wrong content (e.g. the share-link redirect page opened
            ' as a workbook) - discard and try the next candidate.
            book.Close SaveChanges:=False
        End If
    Next i
End Function

' Sheet lookup that ignores capitalisation and stray spaces in tab
' names, so "Summary " or "schedule" still resolve.
Private Function FindSheet(ByVal book As Workbook, _
                           ByVal wantedName As String) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In book.Worksheets
        If LCase$(Trim$(candidate.Name)) = LCase$(Trim$(wantedName)) Then
            Set FindSheet = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Function SheetNameList(ByVal book As Workbook) As String
    Dim candidate As Worksheet
    Dim names As String

    For Each candidate In book.Worksheets
        If Len(names) > 0 Then names = names & ", "
        names = names & candidate.Name
    Next candidate

    SheetNameList = names
End Function

' The link may be stored as a real hyperlink or as plain text; a share
' link's "?e=..." suffix is stripped because Workbooks.Open on some
' builds rejects it.
Private Function MobListLink(ByVal summarySheet As Worksheet) As String
    Dim linkCell As Range
    Dim address As String
    Dim queryPos As Long

    Set linkCell = summarySheet.Range(LINK_CELL)

    If linkCell.Hyperlinks.Count > 0 Then
        address = linkCell.Hyperlinks(1).address
    Else
        address = Trim$(CStr(linkCell.Value))
    End If

    queryPos = InStr(1, address, "?")
    If queryPos > 0 Then address = Left$(address, queryPos - 1)

    MobListLink = address
End Function

Option Explicit

' Pile record sheet creation from the Schedule tab.
'
' Assign CreatePileRecords to a button. For every Schedule row (from
' SCHEDULE_FIRST_DATA_ROW down) whose column M reads "Y", the
' "Record_Template" sheet is copied and named after the Pile ID in
' column A. The copy is then filled from the same Schedule row:
'   B8  = Pile ID              (Schedule column A)
'   B10 = Installation Date    (Schedule column B, planned install date)
'   B20 = How many lines have been installed (Schedule column I, No. lines)
'   B24 = Line 1 "Line end"    (Schedule column H, cable length m)
'
' A row whose Pile ID already exists as a sheet is skipped, never
' overwritten - the workbook already ships with a "P-1" equipment sheet,
' and re-clicking the button must not clobber records filled in on site.
' Values are written as literals at creation time; later edits to the
' Schedule row do not flow into an already-created record sheet.
'
' The new copy is captured by position (last sheet), never via
' ActiveSheet - with ScreenUpdating off, ActiveSheet is unreliable.

Private Const RECORD_TEMPLATE_SHEET As String = "Record_Template"
Private Const SCHEDULE_SHEET As String = "Schedule"
Private Const SCHEDULE_FIRST_DATA_ROW As Long = 11

Private Const COL_PILE_ID As Long = 1       ' A - Pile ID
Private Const COL_INSTALL_DATE As Long = 2  ' B - Planned install date
Private Const COL_CABLE_LENGTH As Long = 8  ' H - Cable length m
Private Const COL_NO_LINES As Long = 9      ' I - No. lines
Private Const COL_RECORD_FLAG As Long = 13  ' M - Pile Records (Y/N)

Public Sub CreatePileRecords()
    Dim scheduleSheet As Worksheet
    Dim templateSheet As Worksheet
    Dim newSheet As Worksheet
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim pileId As String
    Dim sheetName As String
    Dim createdCount As Long
    Dim skippedExisting As String
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim errorMessage As String

    On Error GoTo CleanFail

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Resolved one at a time, tolerating case and stray spaces in the
    ' tab names, so a missing sheet reports which one it was instead of
    ' a bare "subscript out of range".
    Set scheduleSheet = FindSheet(SCHEDULE_SHEET)
    Set templateSheet = FindSheet(RECORD_TEMPLATE_SHEET)

    If scheduleSheet Is Nothing Or templateSheet Is Nothing Then
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating
        MsgBox MissingSheetMessage(scheduleSheet, templateSheet), _
               vbExclamation, "Pile records"
        Exit Sub
    End If

    lastRow = scheduleSheet.Cells(scheduleSheet.Rows.Count, COL_PILE_ID) _
        .End(xlUp).Row

    For rowNumber = SCHEDULE_FIRST_DATA_ROW To lastRow
        If UCase$(Trim$(CStr(scheduleSheet.Cells(rowNumber, COL_RECORD_FLAG).Value))) = "Y" Then
            pileId = Trim$(CStr(scheduleSheet.Cells(rowNumber, COL_PILE_ID).Value))
            If Len(pileId) > 0 Then
                sheetName = SafeSheetName(pileId)
                If SheetExists(sheetName) Then
                    skippedExisting = skippedExisting & vbCrLf & "  " & sheetName
                Else
                    templateSheet.Copy _
                        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)

                    ' Capture the copy by position, not ActiveSheet.
                    Set newSheet = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
                    newSheet.Visible = xlSheetVisible
                    newSheet.Name = sheetName

                    newSheet.Range("B8").Value = pileId
                    newSheet.Range("B10").Value = _
                        scheduleSheet.Cells(rowNumber, COL_INSTALL_DATE).Value
                    newSheet.Range("B20").Value = _
                        scheduleSheet.Cells(rowNumber, COL_NO_LINES).Value
                    newSheet.Range("B24").Value = _
                        scheduleSheet.Cells(rowNumber, COL_CABLE_LENGTH).Value

                    createdCount = createdCount + 1
                End If
            End If
        End If
    Next rowNumber

    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating

    If createdCount = 0 And Len(skippedExisting) = 0 Then
        MsgBox "No Schedule rows are marked ""Y"" in column M.", _
               vbInformation, "Pile records"
    ElseIf Len(skippedExisting) = 0 Then
        MsgBox createdCount & " pile record sheet(s) created.", _
               vbInformation, "Pile records"
    Else
        MsgBox createdCount & " pile record sheet(s) created." & vbCrLf & vbCrLf & _
               "Skipped - a sheet with this name already exists:" & _
               skippedExisting, vbInformation, "Pile records"
    End If
    Exit Sub

CleanFail:
    errorMessage = Err.Description
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not create the pile record sheets." & vbCrLf & vbCrLf & _
           errorMessage, vbExclamation, "Pile records"
End Sub

' Sheet lookup that ignores capitalisation and stray spaces in tab
' names, so "schedule" or "Record_Template " still resolve.
Private Function FindSheet(ByVal wantedName As String) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In ThisWorkbook.Worksheets
        If LCase$(Trim$(candidate.Name)) = LCase$(Trim$(wantedName)) Then
            Set FindSheet = candidate
            Exit Function
        End If
    Next candidate
End Function

' Names the sheet that is missing and lists what this workbook does
' contain - the usual cause is the module sitting in the wrong
' workbook, since Schedule and Record_Template live in the Project Mob
' List file.
Private Function MissingSheetMessage( _
    ByVal scheduleSheet As Worksheet, _
    ByVal templateSheet As Worksheet _
) As String
    Dim missingNames As String
    Dim candidate As Worksheet
    Dim sheetList As String

    If scheduleSheet Is Nothing Then
        missingNames = """" & SCHEDULE_SHEET & """"
    End If
    If templateSheet Is Nothing Then
        If Len(missingNames) > 0 Then missingNames = missingNames & " and "
        missingNames = missingNames & """" & RECORD_TEMPLATE_SHEET & """"
    End If

    For Each candidate In ThisWorkbook.Worksheets
        If Len(sheetList) > 0 Then sheetList = sheetList & ", "
        sheetList = sheetList & candidate.Name
    Next candidate

    MissingSheetMessage = _
        ThisWorkbook.Name & " has no sheet called " & missingNames & "." & _
        vbCrLf & vbCrLf & _
        "Sheets in this workbook: " & sheetList & vbCrLf & vbCrLf & _
        "This macro belongs in the Project Mob List workbook - the one " & _
        "holding the Schedule and Record_Template tabs. If it was " & _
        "imported into another workbook (or into PERSONAL.XLSB), move " & _
        "it to that project, or rename the tabs to match."
End Function

' Excel sheet names are capped at 31 characters and reject : \ / ? * [ ]
Private Function SafeSheetName(ByVal proposedName As String) As String
    Dim cleaned As String
    Dim badChars As Variant
    Dim i As Long

    cleaned = proposedName
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(badChars) To UBound(badChars)
        cleaned = Replace$(cleaned, badChars(i), "-")
    Next i

    If Len(cleaned) > 31 Then cleaned = Left$(cleaned, 31)
    SafeSheetName = cleaned
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim candidate As Worksheet

    On Error Resume Next
    Set candidate = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not candidate Is Nothing
End Function

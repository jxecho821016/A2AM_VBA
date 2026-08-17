Option Explicit

' A2AM Master Tracker - response import (replaces the Power Automate
' flow). Assign UpdateResponses to the UPDATE button.
'
' It appends new rows from the two Forms response workbooks into the
' tracker's own response sheets:
'   A2AM Cable Inventory.xlsx                        -> CableInventoryResponse
'   A2AM Project Mobilisation _ Demobilisation ...   -> ProjectMobDemobResponse
' Everything downstream (Cable Stock formulas, charts) then updates
' itself as it already does.
'
' The response sheets hold Excel Tables, and the tracker's formulas
' address them with structured references such as
' CableInventoryResponse[15m Ordered Cable]. New rows are therefore
' added THROUGH the table (the table is resized to include them);
' writing to the cells underneath would leave the rows outside the
' table where no structured-reference formula can see them.
'
' Columns are matched by HEADER NAME, not by position, so the source
' and destination column orders may differ and extra calculated
' columns in the tracker are left untouched.
'
' A row is imported only if it is not already present. The comparison
' uses the ID column when the response sheet has one (see KEY_HEADERS),
' otherwise a fingerprint of the whole row. Existing rows are never
' modified or reordered - new rows are only ever appended below.
'
' Works in Excel for Mac and Windows. Nothing is saved automatically;
' review the result, then save the tracker yourself.

Private Const HEADER_ROW As Long = 1

Private Const CABLE_SOURCE_FILE As String = "A2AM Cable Inventory.xlsx"
Private Const CABLE_DEST_SHEET As String = "CableInventoryResponse"

Private Const MOB_SOURCE_FILE As String = _
    "A2AM Project Mobilisation _ Demobilisation Tracker.xlsx"
Private Const MOB_DEST_SHEET As String = "ProjectMobDemobResponse"

' Header names treated as a unique row ID, in order of preference.
Private Const KEY_HEADERS As String = _
    "id,responseid,submissionid,starttime"

Public Sub UpdateResponses()
    Dim report As String
    Dim cableReport As String
    Dim mobReport As String
    Dim previousScreenUpdating As Boolean
    Dim previousAlerts As Boolean
    Dim errorMessage As String

    On Error GoTo CleanFail

    previousScreenUpdating = Application.ScreenUpdating
    previousAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    cableReport = ImportResponses(CABLE_SOURCE_FILE, CABLE_DEST_SHEET)
    mobReport = ImportResponses(MOB_SOURCE_FILE, MOB_DEST_SHEET)

    ' Structured-reference formulas only pick the new rows up once the
    ' resized tables have been recalculated.
    Application.CalculateFull

    Application.DisplayAlerts = previousAlerts
    Application.ScreenUpdating = previousScreenUpdating

    report = "Cable Inventory:" & vbCrLf & "  " & cableReport & vbCrLf & _
             vbCrLf & _
             "Project Mob / Demob:" & vbCrLf & "  " & mobReport & vbCrLf & _
             vbCrLf & "Save the tracker to keep the imported rows."

    MsgBox report, vbInformation, "Update responses"
    Exit Sub

CleanFail:
    errorMessage = Err.Description
    Application.DisplayAlerts = previousAlerts
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "The response import did not finish." & vbCrLf & vbCrLf & _
           errorMessage, vbExclamation, "Update responses"
End Sub

' Imports one source workbook into one destination sheet and returns a
' one-line summary for the report.
Private Function ImportResponses( _
    ByVal sourceFileName As String, _
    ByVal destSheetName As String _
) As String
    Dim destSheet As Worksheet
    Dim destTable As ListObject
    Dim destHeaderRange As Range
    Dim destDataRange As Range
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceTable As ListObject
    Dim sourceHeaderRange As Range
    Dim sourceDataRange As Range
    Dim sourcePath As String
    Dim sourceValues As Variant
    Dim destValues As Variant
    Dim destHeaders As Variant
    Dim sourceHeaders As Variant
    Dim columnMap As Variant
    Dim existingKeys As Collection
    Dim newRows As Collection
    Dim rowValues As Variant
    Dim outputBlock As Variant
    Dim rowKey As String
    Dim keyColumn As Long
    Dim sourceRow As Long
    Dim sourceCol As Long
    Dim destCol As Long
    Dim mappedCount As Long
    Dim addedCount As Long
    Dim skippedCount As Long
    Dim blankCount As Long
    Dim i As Long

    Set destSheet = FindSheet(ThisWorkbook, destSheetName)
    If destSheet Is Nothing Then
        ImportResponses = "skipped - this workbook has no """ & _
                          destSheetName & """ sheet."
        Exit Function
    End If

    GetSheetLayout destSheet, destTable, destHeaderRange, destDataRange
    If destHeaderRange Is Nothing Then
        ImportResponses = "skipped - no header row on """ & _
                          destSheetName & """."
        Exit Function
    End If

    sourcePath = ResolveSourcePath(sourceFileName)
    If Len(sourcePath) = 0 Then
        ImportResponses = "skipped - " & sourceFileName & " was not found."
        Exit Function
    End If

    Set sourceBook = OpenSourceReadOnly(sourcePath)
    If sourceBook Is Nothing Then
        ImportResponses = "skipped - could not open " & sourceFileName & "."
        Exit Function
    End If

    On Error GoTo CleanFail

    ' Forms response workbooks keep the responses on the first sheet.
    Set sourceSheet = sourceBook.Worksheets(1)
    GetSheetLayout sourceSheet, sourceTable, sourceHeaderRange, sourceDataRange

    If sourceHeaderRange Is Nothing Then
        sourceBook.Close SaveChanges:=False
        On Error GoTo 0
        ImportResponses = "skipped - no header row in " & sourceFileName & "."
        Exit Function
    End If

    destHeaders = HeaderKeys(destHeaderRange)
    sourceHeaders = HeaderKeys(sourceHeaderRange)

    ' Map each source column to the destination column with the same
    ' header. Unmatched source columns are ignored.
    ReDim columnMap(1 To UBound(sourceHeaders))
    For sourceCol = 1 To UBound(sourceHeaders)
        columnMap(sourceCol) = 0
        If Len(sourceHeaders(sourceCol)) > 0 Then
            For destCol = 1 To UBound(destHeaders)
                If destHeaders(destCol) = sourceHeaders(sourceCol) Then
                    columnMap(sourceCol) = destCol
                    mappedCount = mappedCount + 1
                    Exit For
                End If
            Next destCol
        End If
    Next sourceCol

    If mappedCount = 0 Then
        sourceBook.Close SaveChanges:=False
        On Error GoTo 0
        ImportResponses = "skipped - no column headers matched """ & _
                          destSheetName & """."
        Exit Function
    End If

    keyColumn = PreferredKeyColumn(sourceHeaders)

    ' Existing rows, so nothing is imported twice - including rows the
    ' Power Automate flow already copied.
    destValues = RangeValues(destDataRange)
    Set existingKeys = ExistingRowKeys( _
        destValues, UBound(sourceHeaders), columnMap, keyColumn)

    sourceValues = RangeValues(sourceDataRange)
    Set newRows = New Collection

    If Not IsEmpty(sourceValues) Then
        For sourceRow = 1 To UBound(sourceValues, 1)
            rowKey = RowKey(sourceValues, sourceRow, columnMap, keyColumn)

            If Len(rowKey) = 0 Then
                blankCount = blankCount + 1
            ElseIf KeyExists(existingKeys, rowKey) Then
                skippedCount = skippedCount + 1
            Else
                ReDim rowValues(1 To UBound(destHeaders))
                For sourceCol = 1 To UBound(columnMap)
                    destCol = columnMap(sourceCol)
                    If destCol > 0 Then
                        rowValues(destCol) = sourceValues(sourceRow, sourceCol)
                    End If
                Next sourceCol
                newRows.Add rowValues
                existingKeys.Add True, rowKey  ' guard duplicates in source
                addedCount = addedCount + 1
            End If
        Next sourceRow
    End If

    If addedCount > 0 Then
        ReDim outputBlock(1 To addedCount, 1 To UBound(destHeaders))
        For i = 1 To addedCount
            rowValues = newRows(i)
            For destCol = 1 To UBound(destHeaders)
                outputBlock(i, destCol) = rowValues(destCol)
            Next destCol
        Next i

        AppendRows destSheet, destTable, destHeaderRange, destDataRange, _
                   outputBlock, addedCount, UBound(destHeaders)
    End If

    sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing
    On Error GoTo 0

    ImportResponses = addedCount & " new row(s) added, " & _
                      skippedCount & " already present" & _
                      IIf(blankCount > 0, ", " & blankCount & " blank", _
                          vbNullString) & "."
    Exit Function

CleanFail:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    ImportResponses = "failed - " & Err.Description
End Function

' Header and data ranges of a sheet. When the sheet holds an Excel
' Table the table defines them, so structured-reference formulas and
' this import agree on where the data is.
Private Sub GetSheetLayout( _
    ByVal sheet As Worksheet, _
    ByRef sheetTable As ListObject, _
    ByRef headerRange As Range, _
    ByRef dataRange As Range _
)
    Dim lastColumn As Long
    Dim lastRow As Long

    Set sheetTable = Nothing
    Set headerRange = Nothing
    Set dataRange = Nothing

    If sheet.ListObjects.Count > 0 Then
        Set sheetTable = sheet.ListObjects(1)
        Set headerRange = sheetTable.HeaderRowRange
        If Not sheetTable.DataBodyRange Is Nothing Then
            Set dataRange = sheetTable.DataBodyRange
        End If
        Exit Sub
    End If

    lastColumn = sheet.Cells(HEADER_ROW, sheet.Columns.Count) _
        .End(xlToLeft).Column
    If lastColumn < 1 Then Exit Sub
    If Len(Trim$(CStr(sheet.Cells(HEADER_ROW, 1).value))) = 0 And _
       lastColumn = 1 Then Exit Sub

    Set headerRange = sheet.Cells(HEADER_ROW, 1).Resize(1, lastColumn)

    lastRow = LastDataRow(sheet, lastColumn)
    If lastRow > HEADER_ROW Then
        Set dataRange = sheet.Cells(HEADER_ROW + 1, 1) _
            .Resize(lastRow - HEADER_ROW, lastColumn)
    End If
End Sub

' Writes the new rows and, for a table, resizes it to take them in.
' Structured references only see rows inside the table's range.
Private Sub AppendRows( _
    ByVal destSheet As Worksheet, _
    ByVal destTable As ListObject, _
    ByVal destHeaderRange As Range, _
    ByVal destDataRange As Range, _
    ByVal outputBlock As Variant, _
    ByVal rowCount As Long, _
    ByVal columnCount As Long _
)
    Dim firstColumn As Long
    Dim writeStart As Range
    Dim lastCell As Range

    firstColumn = destHeaderRange.Column

    If destDataRange Is Nothing Then
        ' Header only: the first data row sits directly beneath it.
        Set writeStart = destSheet.Cells(destHeaderRange.Row + 1, firstColumn)
    ElseIf destDataRange.Rows.Count = 1 And IsBlankRange(destDataRange) Then
        ' A new table keeps one empty placeholder row - fill that first.
        Set writeStart = destSheet.Cells(destDataRange.Row, firstColumn)
    Else
        Set writeStart = destSheet.Cells( _
            destDataRange.Row + destDataRange.Rows.Count, firstColumn)
    End If

    writeStart.Resize(rowCount, columnCount).value = outputBlock

    If destTable Is Nothing Then Exit Sub

    Set lastCell = writeStart.Offset(rowCount - 1, columnCount - 1)
    destTable.Resize destSheet.Range(destHeaderRange.Cells(1, 1), lastCell)
End Sub

' Header row as normalised keys, 1-based within the header range.
Private Function HeaderKeys(ByVal headerRange As Range) As Variant
    Dim keys() As String
    Dim col As Long

    ReDim keys(1 To headerRange.Columns.Count)
    For col = 1 To headerRange.Columns.Count
        keys(col) = NormaliseKey(headerRange.Cells(1, col).value)
    Next col

    HeaderKeys = keys
End Function

' Range contents as a 2-D array, or Empty when there is no data.
Private Function RangeValues(ByVal target As Range) As Variant
    Dim block As Variant
    Dim single_(1 To 1, 1 To 1) As Variant

    If target Is Nothing Then Exit Function

    block = target.value
    If IsArray(block) Then
        RangeValues = block
    Else
        single_(1, 1) = block
        RangeValues = single_
    End If
End Function

Private Function IsBlankRange(ByVal target As Range) As Boolean
    IsBlankRange = (Application.WorksheetFunction.CountA(target) = 0)
End Function

' Last row holding data across the given columns.
Private Function LastDataRow( _
    ByVal sheet As Worksheet, _
    ByVal columnCount As Long _
) As Long
    Dim col As Long
    Dim candidate As Long

    LastDataRow = HEADER_ROW
    For col = 1 To columnCount
        candidate = sheet.Cells(sheet.Rows.Count, col).End(xlUp).Row
        If candidate > LastDataRow Then LastDataRow = candidate
    Next col
End Function

' Column holding the response ID, or 0 to fingerprint the whole row.
Private Function PreferredKeyColumn(ByVal sourceHeaders As Variant) As Long
    Dim wanted As Variant
    Dim i As Long
    Dim col As Long

    wanted = Split(KEY_HEADERS, ",")
    For i = LBound(wanted) To UBound(wanted)
        For col = 1 To UBound(sourceHeaders)
            If sourceHeaders(col) = NormaliseKey(wanted(i)) Then
                PreferredKeyColumn = col
                Exit Function
            End If
        Next col
    Next i
End Function

' Keys of the rows already in the destination table, built the same
' way as the source keys so the two are comparable.
Private Function ExistingRowKeys( _
    ByVal destValues As Variant, _
    ByVal sourceColumnCount As Long, _
    ByVal columnMap As Variant, _
    ByVal keyColumn As Long _
) As Collection
    Dim keys As Collection
    Dim mirrored As Variant
    Dim destRow As Long
    Dim sourceCol As Long
    Dim destCol As Long
    Dim rowKey As String

    Set keys = New Collection
    Set ExistingRowKeys = keys

    If IsEmpty(destValues) Then Exit Function

    ' Re-shape each destination row into source-column order so the
    ' same RowKey routine produces a comparable key.
    For destRow = 1 To UBound(destValues, 1)
        ReDim mirrored(1 To 1, 1 To sourceColumnCount)
        For sourceCol = 1 To UBound(columnMap)
            destCol = columnMap(sourceCol)
            If destCol > 0 Then
                mirrored(1, sourceCol) = destValues(destRow, destCol)
            End If
        Next sourceCol

        rowKey = RowKey(mirrored, 1, columnMap, keyColumn)
        If Len(rowKey) > 0 Then
            On Error Resume Next
            keys.Add True, rowKey
            On Error GoTo 0
        End If
    Next destRow
End Function

' Identity of one row: the ID cell, or every mapped value joined.
Private Function RowKey( _
    ByVal values As Variant, _
    ByVal rowIndex As Long, _
    ByVal columnMap As Variant, _
    ByVal keyColumn As Long _
) As String
    Dim parts As String
    Dim sourceCol As Long

    If keyColumn > 0 Then
        RowKey = NormaliseKey(values(rowIndex, keyColumn))
        Exit Function
    End If

    For sourceCol = 1 To UBound(columnMap)
        If columnMap(sourceCol) > 0 Then
            parts = parts & NormaliseKey(values(rowIndex, sourceCol)) & Chr$(1)
        End If
    Next sourceCol

    ' A row of nothing but empty cells has no identity.
    If Len(Replace(parts, Chr$(1), vbNullString)) = 0 Then Exit Function
    RowKey = parts
End Function

Private Function KeyExists( _
    ByVal keys As Collection, _
    ByVal keyText As String _
) As Boolean
    Dim probe As Variant

    On Error Resume Next
    probe = keys.Item(keyText)
    KeyExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' Dates and numbers must key identically whether they arrive as text
' or as typed values, so both are rendered in a fixed form.
Private Function NormaliseKey(ByVal rawValue As Variant) As String
    Dim text As String

    If IsError(rawValue) Then
        NormaliseKey = "#err"
        Exit Function
    End If
    If IsEmpty(rawValue) Then Exit Function

    If VarType(rawValue) = vbDate Then
        text = Format$(CDate(rawValue), "yyyy-mm-dd hh:nn:ss")
    ElseIf IsNumeric(rawValue) And VarType(rawValue) <> vbString Then
        text = Format$(CDbl(rawValue), "0.##########")
    Else
        text = CStr(rawValue)
    End If

    NormaliseKey = LCase$(Trim$(text))
End Function

Private Function FindSheet( _
    ByVal book As Workbook, _
    ByVal wantedName As String _
) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In book.Worksheets
        If LCase$(Trim$(candidate.Name)) = LCase$(Trim$(wantedName)) Then
            Set FindSheet = candidate
            Exit Function
        End If
    Next candidate
End Function

' Looks for the source workbook beside the tracker, then in the usual
' sub-folders, and finally asks the user to point at it.
Private Function ResolveSourcePath(ByVal fileName As String) As String
    Dim separator As String
    Dim folders As Variant
    Dim candidate As String
    Dim i As Long

    separator = Application.PathSeparator

    folders = Array( _
        ThisWorkbook.Path, _
        ThisWorkbook.Path & separator & "Responses", _
        ThisWorkbook.Path & separator & "Forms", _
        ThisWorkbook.Path & separator & "Apps")

    For i = LBound(folders) To UBound(folders)
        If Len(CStr(folders(i))) > 0 Then
            candidate = CStr(folders(i)) & separator & fileName
            If FileExists(candidate) Then
                ResolveSourcePath = candidate
                Exit Function
            End If
        End If
    Next i

    ResolveSourcePath = AskForSourceFile(fileName)
End Function

Private Function AskForSourceFile(ByVal fileName As String) As String
    Dim chosen As Variant

    chosen = Application.GetOpenFilename( _
        FileFilter:="Excel files (*.xlsx; *.xlsm), *.xlsx; *.xlsm", _
        Title:="Locate " & fileName)

    If VarType(chosen) = vbBoolean Then Exit Function  ' cancelled
    AskForSourceFile = CStr(chosen)
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    Dim result As String

    On Error Resume Next
    result = Dir$(fullPath)
    On Error GoTo 0

    FileExists = (Len(result) > 0)
End Function

' Reuses the workbook if it is already open, otherwise opens it
' read-only. On Mac, sandboxed Excel must be granted access to the
' file before it can be opened from a path.
Private Function OpenSourceReadOnly(ByVal fullPath As String) As Workbook
    Dim book As Workbook
    Dim fileName As String
    Dim separator As String

    separator = Application.PathSeparator
    fileName = Mid$(fullPath, InStrRev(fullPath, separator) + 1)

    For Each book In Application.Workbooks
        If StrComp(book.Name, fileName, vbTextCompare) = 0 Then
            Set OpenSourceReadOnly = book
            Exit Function
        End If
    Next book

    #If Mac Then
        Dim grantPaths(0) As Variant
        grantPaths(0) = fullPath
        If Val(Application.Version) >= 15 Then
            If Not GrantAccessToMultipleFiles(grantPaths) Then Exit Function
        End If
    #End If

    On Error Resume Next
    Set OpenSourceReadOnly = Workbooks.Open( _
        fileName:=fullPath, _
        UpdateLinks:=False, _
        ReadOnly:=True, _
        AddToMru:=False)
    On Error GoTo 0
End Function

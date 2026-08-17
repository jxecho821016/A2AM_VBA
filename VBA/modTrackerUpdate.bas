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
    "id,responseid,submissionid,startTime,starttime"

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
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourcePath As String
    Dim sourceValues As Variant
    Dim destHeaders As Variant
    Dim sourceHeaders As Variant
    Dim columnMap As Variant
    Dim existingKeys As Collection
    Dim newRows As Collection
    Dim rowValues As Variant
    Dim outputBlock As Variant
    Dim rowKey As String
    Dim keyColumn As Long
    Dim lastDestRow As Long
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

    destHeaders = HeaderKeys(destSheet)
    sourceHeaders = HeaderKeys(sourceSheet)

    If UBound(sourceHeaders) < 1 Or UBound(destHeaders) < 1 Then
        sourceBook.Close SaveChanges:=False
        On Error GoTo 0
        ImportResponses = "skipped - no header row found."
        Exit Function
    End If

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
    Set existingKeys = ExistingRowKeys( _
        destSheet, destHeaders, sourceHeaders, columnMap, keyColumn)

    sourceValues = ResponseValues(sourceSheet, UBound(sourceHeaders))
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

        lastDestRow = LastDataRow(destSheet, UBound(destHeaders))
        destSheet.Cells(lastDestRow + 1, 1) _
            .Resize(addedCount, UBound(destHeaders)).value = outputBlock
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

' Header row of a sheet as normalised keys, 1-based by column.
Private Function HeaderKeys(ByVal sheet As Worksheet) As Variant
    Dim lastColumn As Long
    Dim keys() As String
    Dim col As Long

    lastColumn = sheet.Cells(HEADER_ROW, sheet.Columns.Count) _
        .End(xlToLeft).Column
    If lastColumn < 1 Then lastColumn = 1

    ReDim keys(1 To lastColumn)
    For col = 1 To lastColumn
        keys(col) = NormaliseKey(sheet.Cells(HEADER_ROW, col).value)
    Next col

    HeaderKeys = keys
End Function

' All response rows below the header as a 2-D array (or Empty).
Private Function ResponseValues( _
    ByVal sheet As Worksheet, _
    ByVal columnCount As Long _
) As Variant
    Dim lastRow As Long
    Dim block As Variant

    lastRow = LastDataRow(sheet, columnCount)
    If lastRow <= HEADER_ROW Then Exit Function

    block = sheet.Cells(HEADER_ROW + 1, 1) _
        .Resize(lastRow - HEADER_ROW, columnCount).value

    ' A single row comes back as a plain value when only one column is
    ' involved; force a 2-D shape for the caller.
    If Not IsArray(block) Then
        Dim single_(1 To 1, 1 To 1) As Variant
        single_(1, 1) = block
        ResponseValues = single_
    Else
        ResponseValues = block
    End If
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

' Keys of the rows already in the destination sheet, built the same
' way as the source keys so the two are comparable.
Private Function ExistingRowKeys( _
    ByVal destSheet As Worksheet, _
    ByVal destHeaders As Variant, _
    ByVal sourceHeaders As Variant, _
    ByVal columnMap As Variant, _
    ByVal keyColumn As Long _
) As Collection
    Dim keys As Collection
    Dim destValues As Variant
    Dim mirrored As Variant
    Dim lastRow As Long
    Dim destRow As Long
    Dim sourceCol As Long
    Dim destCol As Long
    Dim rowKey As String

    Set keys = New Collection
    Set ExistingRowKeys = keys

    lastRow = LastDataRow(destSheet, UBound(destHeaders))
    If lastRow <= HEADER_ROW Then Exit Function

    destValues = destSheet.Cells(HEADER_ROW + 1, 1) _
        .Resize(lastRow - HEADER_ROW, UBound(destHeaders)).value
    If Not IsArray(destValues) Then Exit Function

    ' Re-shape each destination row into source-column order so the
    ' same RowKey routine produces a comparable key.
    For destRow = 1 To UBound(destValues, 1)
        ReDim mirrored(1 To 1, 1 To UBound(sourceHeaders))
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

    If IsDate(rawValue) And Not VarType(rawValue) = vbString Then
        text = Format$(CDate(rawValue), "yyyy-mm-dd hh:nn:ss")
    ElseIf IsNumeric(rawValue) And Not VarType(rawValue) = vbString Then
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

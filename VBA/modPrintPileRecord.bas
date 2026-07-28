Option Explicit

' Pile record "print" to image.
'
' Assign PrintPileRecord to a button (put the button on Record_Template
' before creating copies and every generated pile record sheet carries
' it). Clicking it exports the active record sheet's print area
' (A1:B57 on the template) as a JPG:
'   1. A folder picker asks where to save.
'   2. The file is named <Pile ID>_Pile_Record.jpg, e.g.
'      P-1_Pile_Record.jpg. The pile ID comes from B8, falling back to
'      the sheet name if B8 is empty.
'   3. An existing file with that name prompts before overwriting.
'
' The export copies the range as a BITMAP (an exact snapshot of the
' filled-in cells - metafile copies render blank on export in several
' Excel builds), pastes it into a temporary chart parked in the visible
' window area so Excel actually draws it, exports the chart as JPG,
' then deletes it.

Private Const RECORD_TITLE As String = "TIP Testing Site Record Sheet"
Private Const FILE_SUFFIX As String = "_Pile_Record.jpg"

Public Sub PrintPileRecord()
    Dim recordSheet As Worksheet
    Dim exportRange As Range
    Dim chartObj As ChartObject
    Dim pileId As String
    Dim folderPath As String
    Dim filePath As String
    Dim previousScreenUpdating As Boolean
    Dim errorMessage As String

    On Error GoTo CleanFail

    previousScreenUpdating = Application.ScreenUpdating
    Set recordSheet = ActiveSheet

    ' Only run on a pile record sheet (a copy of Record_Template).
    If CStr(recordSheet.Range("A1").Value) <> RECORD_TITLE Then
        MsgBox "Please run this from a pile record sheet.", _
               vbExclamation, "Print pile record"
        Exit Sub
    End If
    If recordSheet.Name = "Record_Template" Then
        MsgBox "This is the template. Open the pile's own record sheet " & _
               "and print from there.", vbExclamation, "Print pile record"
        Exit Sub
    End If

    pileId = Trim$(CStr(recordSheet.Range("B8").Value))
    If Len(pileId) = 0 Then pileId = recordSheet.Name

    ' Ask for the save location.
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Choose where to save " & pileId & FILE_SUFFIX
        If Len(ThisWorkbook.Path) > 0 Then
            .InitialFileName = ThisWorkbook.Path & Application.PathSeparator
        End If
        If .Show = 0 Then Exit Sub  ' user cancelled
        folderPath = .SelectedItems(1)
    End With

    filePath = folderPath & Application.PathSeparator & _
               SafeFileName(pileId) & FILE_SUFFIX

    If Len(Dir$(filePath)) > 0 Then
        If MsgBox(filePath & vbCrLf & vbCrLf & _
                  "This file already exists. Overwrite it?", _
                  vbYesNo + vbQuestion, "Print pile record") = vbNo Then
            Exit Sub
        End If
    End If

    ' Export the sheet's print area; fall back to the used range if no
    ' print area is set on this copy.
    On Error Resume Next
    Set exportRange = recordSheet.Range(recordSheet.PageSetup.PrintArea)
    On Error GoTo CleanFail
    If exportRange Is Nothing Then Set exportRange = recordSheet.UsedRange

    ' Screen updating stays ON here: the chart must genuinely render
    ' on screen, otherwise Excel exports it blank or skips the paste.
    Application.ScreenUpdating = True
    recordSheet.Activate

    ' Temporary chart sized exactly to the range, parked at the top of
    ' the visible window so Excel draws it before the export.
    Set chartObj = recordSheet.ChartObjects.Add( _
        Left:=ActiveWindow.VisibleRange.Left, _
        Top:=ActiveWindow.VisibleRange.Top, _
        Width:=exportRange.Width, _
        Height:=exportRange.Height)
    chartObj.Chart.ChartArea.Format.Line.Visible = msoFalse
    chartObj.Activate

    ' Bitmap copy = pixel-exact snapshot of the filled-in cells. The
    ' metafile (xlPicture) alternative exports blank in several Excel
    ' builds, so it is only the fallback.
    exportRange.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    chartObj.Chart.Paste
    DoEvents
    If chartObj.Chart.Shapes.Count = 0 Then
        exportRange.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        chartObj.Chart.Paste
        DoEvents
    End If
    If chartObj.Chart.Shapes.Count = 0 Then
        Err.Raise vbObjectError + 513, , _
            "Excel did not paste the sheet picture into the export " & _
            "chart. Please try again."
    End If

    With chartObj.Chart
        With .Shapes(1)
            .LockAspectRatio = msoFalse
            .Left = 0
            .Top = 0
            .Width = chartObj.Width
            .Height = chartObj.Height
        End With
        .Refresh
        DoEvents  ' let the paste render before exporting
        .Export Filename:=filePath, FilterName:="JPG"
    End With

    chartObj.Delete
    Set chartObj = Nothing
    Application.CutCopyMode = False
    recordSheet.Range("A1").Select  ' drop the chart selection
    Application.ScreenUpdating = previousScreenUpdating

    MsgBox "Saved:" & vbCrLf & filePath, vbInformation, "Print pile record"
    Exit Sub

CleanFail:
    errorMessage = Err.Description
    On Error Resume Next
    If Not chartObj Is Nothing Then chartObj.Delete
    On Error GoTo 0
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "Could not save the pile record image." & vbCrLf & vbCrLf & _
           errorMessage, vbExclamation, "Print pile record"
End Sub

' Windows filenames reject \ / : * ? " < > |
Private Function SafeFileName(ByVal proposedName As String) As String
    Dim cleaned As String
    Dim badChars As Variant
    Dim i As Long

    cleaned = proposedName
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(badChars) To UBound(badChars)
        cleaned = Replace$(cleaned, badChars(i), "-")
    Next i

    SafeFileName = cleaned
End Function

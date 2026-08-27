Option Explicit

' Pile record "print" to image.
'
' Assign PrintPileRecord to a button (put the button on Record_Template
' before creating copies and every generated pile record sheet carries
' it). Clicking it exports the active record sheet's print area
' (A1:B57 on the template) as a single JPG image:
'   1. A folder picker asks where to save.
'   2. The file is named <Pile ID>_Pile_Record.jpg, e.g.
'      P-1_Pile_Record.jpg. The pile ID comes from B8, falling back to
'      the sheet name if B8 is empty.
'   3. An existing file with that name prompts before overwriting.
'
' How it works: the range is copied as a bitmap (a pixel-exact snapshot
' of the filled-in cells) and written straight from the clipboard to a
' JPG via the Windows GDI+ API. Earlier chart-paste-export versions
' produced blank images on some Excel builds; this route has no
' rendering step and saves exactly what was copied.

Private Const RECORD_TITLE As String = "TIP Testing Site Record Sheet"
Private Const FILE_SUFFIX As String = "_Pile_Record.jpg"

Private Const CF_BITMAP As Long = 2
Private Const JPEG_ENCODER_CLSID As String = _
    "{557CF401-1A04-11D3-9A73-0000F81EF32E}"

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Type GdiplusStartupInput
        GdiplusVersion As Long
        DebugEventCallback As LongPtr
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type

    Private Declare PtrSafe Function OpenClipboard Lib "user32" _
        (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" _
        (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" _
        (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" _
        (token As LongPtr, inputBuf As GdiplusStartupInput, _
         Optional ByVal outputBuf As LongPtr = 0) As Long
    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" _
        (ByVal token As LongPtr)
    Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" _
        (ByVal hbm As LongPtr, ByVal hPal As LongPtr, _
         bitmap As LongPtr) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" _
        (ByVal image As LongPtr) As Long
    Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" _
        (ByVal image As LongPtr, ByVal fileName As LongPtr, _
         clsidEncoder As GUID, ByVal encoderParams As LongPtr) As Long
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32" _
        (ByVal lpszProgID As LongPtr, pclsid As GUID) As Long
#Else
    Private Type GdiplusStartupInput
        GdiplusVersion As Long
        DebugEventCallback As Long
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type

    Private Declare Function OpenClipboard Lib "user32" _
        (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function GetClipboardData Lib "user32" _
        (ByVal wFormat As Long) As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" _
        (ByVal wFormat As Long) As Long
    Private Declare Function GdiplusStartup Lib "gdiplus" _
        (token As Long, inputBuf As GdiplusStartupInput, _
         Optional ByVal outputBuf As Long = 0) As Long
    Private Declare Sub GdiplusShutdown Lib "gdiplus" _
        (ByVal token As Long)
    Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" _
        (ByVal hbm As Long, ByVal hPal As Long, bitmap As Long) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus" _
        (ByVal image As Long) As Long
    Private Declare Function GdipSaveImageToFile Lib "gdiplus" _
        (ByVal image As Long, ByVal fileName As Long, _
         clsidEncoder As GUID, ByVal encoderParams As Long) As Long
    Private Declare Function CLSIDFromString Lib "ole32" _
        (ByVal lpszProgID As Long, pclsid As GUID) As Long
#End If

Public Sub PrintPileRecord()
    Dim recordSheet As Worksheet
    Dim exportRange As Range
    Dim pileId As String
    Dim folderPath As String
    Dim filePath As String
    Dim errorMessage As String

    On Error GoTo CleanFail

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

    ' Ask for the save location. A workbook opened from SharePoint
    ' reports an https:// address as its path - a web location, not a
    ' folder - so it is only offered as a starting point when it is a
    ' real path on this computer.
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Choose where to save " & pileId & FILE_SUFFIX
        If IsLocalPath(ThisWorkbook.Path) Then
            .InitialFileName = ThisWorkbook.Path & Application.PathSeparator
        End If
        If .Show = 0 Then Exit Sub  ' user cancelled
        folderPath = .SelectedItems(1)
    End With

    Do While Len(folderPath) > 1 And _
             Right$(folderPath, 1) = Application.PathSeparator
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop

    If Not IsLocalPath(folderPath) Then
        MsgBox "The picture can only be saved to a folder on this " & _
               "computer, and this is a web address:" & vbCrLf & _
               folderPath & vbCrLf & vbCrLf & _
               "Choose a local folder instead - a OneDrive folder that " & _
               "syncs to this PC is fine.", _
               vbExclamation, "Print pile record"
        Exit Sub
    End If

    filePath = folderPath & Application.PathSeparator & _
               SafeFileName(pileId) & FILE_SUFFIX

    If FileExists(filePath) Then
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

    ' Snapshot the whole range (all of it, in one image - no paging)
    ' onto the clipboard, then write the clipboard bitmap to the JPG.
    exportRange.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    DoEvents
    If IsClipboardFormatAvailable(CF_BITMAP) = 0 Then
        ' Occasionally the first copy does not land; try once more.
        exportRange.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        DoEvents
    End If

    SaveClipboardBitmapAsJpeg filePath  ' raises a VBA error on failure

    Application.CutCopyMode = False

    MsgBox "Saved:" & vbCrLf & filePath, vbInformation, "Print pile record"
    Exit Sub

CleanFail:
    errorMessage = Err.Description
    Application.CutCopyMode = False
    MsgBox "Could not save the pile record image." & vbCrLf & vbCrLf & _
           errorMessage, vbExclamation, "Print pile record"
End Sub

' Writes the bitmap currently on the clipboard to filePath as a JPG
' using GDI+. Raises a VBA error if any step fails.
Private Sub SaveClipboardBitmapAsJpeg(ByVal filePath As String)
    #If VBA7 Then
        Dim gdiplusToken As LongPtr
        Dim hBitmap As LongPtr
        Dim gdipImage As LongPtr
    #Else
        Dim gdiplusToken As Long
        Dim hBitmap As Long
        Dim gdipImage As Long
    #End If
    Dim startupInput As GdiplusStartupInput
    Dim jpegClsid As GUID
    Dim clipboardOpen As Boolean
    Dim gdiplusStarted As Boolean
    Dim failureText As String

    On Error GoTo CleanFail

    If IsClipboardFormatAvailable(CF_BITMAP) = 0 Then
        failureText = "No picture found on the clipboard."
        GoTo CleanFail
    End If

    If OpenClipboard(0) = 0 Then
        failureText = "Could not open the Windows clipboard."
        GoTo CleanFail
    End If
    clipboardOpen = True

    hBitmap = GetClipboardData(CF_BITMAP)
    If hBitmap = 0 Then
        failureText = "Could not read the picture from the clipboard."
        GoTo CleanFail
    End If

    startupInput.GdiplusVersion = 1
    If GdiplusStartup(gdiplusToken, startupInput) <> 0 Then
        failureText = "Could not start the Windows imaging system (GDI+)."
        GoTo CleanFail
    End If
    gdiplusStarted = True

    If GdipCreateBitmapFromHBITMAP(hBitmap, 0, gdipImage) <> 0 Then
        failureText = "Could not convert the clipboard picture."
        GoTo CleanFail
    End If

    CLSIDFromString StrPtr(JPEG_ENCODER_CLSID), jpegClsid
    If GdipSaveImageToFile(gdipImage, StrPtr(filePath), jpegClsid, 0) <> 0 Then
        failureText = "Could not write the JPG file:" & vbCrLf & filePath
        GoTo CleanFail
    End If

    GdipDisposeImage gdipImage
    GdiplusShutdown gdiplusToken
    CloseClipboard
    Exit Sub

CleanFail:
    If Len(failureText) = 0 Then failureText = Err.Description
    If gdipImage <> 0 Then GdipDisposeImage gdipImage
    If gdiplusStarted Then GdiplusShutdown gdiplusToken
    If clipboardOpen Then CloseClipboard
    Err.Raise vbObjectError + 514, "SaveClipboardBitmapAsJpeg", failureText
End Sub

' Windows filenames reject \ / : * ? " < > | and control characters,
' and may not end with a space or a dot. A pile ID carrying a line
' break is the usual hidden cause of "Bad file name or number".
Private Function SafeFileName(ByVal proposedName As String) As String
    Dim character As String
    Dim result As String
    Dim position As Long

    For position = 1 To Len(proposedName)
        character = Mid$(proposedName, position, 1)
        If InStr(1, "\/:*?""<>|", character) > 0 Then
            result = result & "-"
        ElseIf AscW(character) < 32 Then
            result = result & " "
        Else
            result = result & character
        End If
    Next position

    result = Trim$(result)
    Do While Len(result) > 0 And Right$(result, 1) = "."
        result = Trim$(Left$(result, Len(result) - 1))
    Loop

    If Len(result) = 0 Then result = "Pile"
    SafeFileName = result
End Function

' True for a path this machine can write to. A workbook opened
' straight from SharePoint reports an https:// path, which is a web
' address rather than a folder.
Private Function IsLocalPath(ByVal candidate As String) As Boolean
    Dim normalised As String

    normalised = LCase$(Trim$(candidate))
    If Len(normalised) = 0 Then Exit Function
    If Left$(normalised, 7) = "http://" Then Exit Function
    If Left$(normalised, 8) = "https://" Then Exit Function

    IsLocalPath = True
End Function

' Dir$ raises "Bad file name or number" on a path it cannot parse, so
' the existence check never reaches the caller as a crash.
Private Function FileExists(ByVal fullPath As String) As Boolean
    Dim found As String

    On Error Resume Next
    found = Dir$(fullPath)
    On Error GoTo 0

    FileExists = (Len(found) > 0)
End Function

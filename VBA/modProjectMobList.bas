Option Explicit

Private Const CACHE_SHEET As String = "_AssetCache"
Private Const SUBMISSION_LOG_SHEET As String = "_SubmissionLog"
Private Const MASTER_FILE As String = "Apps\A2AM Master Tracker.xlsx"
' Web link to the Master Tracker. The link stored on the Site Details
' sheet ("A2AM Master Tracker" row) takes priority; this constant is
' the fallback if that row is empty or deleted.
Private Const MASTER_LINK As String = _
    "https://a2am.sharepoint.com/:x:/s/A2AdvancedMonitoring/" & _
    "IQCaaWXJAcaUQ7R66vheIRVVAXhTX2pe1rx6ABQlsP7vJ7g"
Private Const SITE_DETAILS_SHEET As String = "Site Details"
Private Const MASTER_LINK_LABEL As String = "Master Tracker"
Private Const MASTER_SHEET As String = "Equipment"
Private Const LEAD_CABLE_SHEET As String = "Lead Cables"
Private Const ITEM_HEADER_ROW As Long = 15
Private Const FIRST_ITEM_ROW As Long = 16
Private Const LAST_ITEM_ROW As Long = 50

Public Sub Button1_Click()
    CreateNewEntry
End Sub

Public Sub CreateNewEntry()
    Dim templateSheet As Worksheet
    Dim entrySheet As Worksheet
    Dim baseName As String
    Dim newName As String
    Dim counter As Long
    Dim movementChoice As VbMsgBoxResult
    Dim movementText As String
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim errorMessage As String

    On Error GoTo CleanFail

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set templateSheet = ThisWorkbook.Worksheets("Template")

    movementChoice = MsgBox( _
        "Choose the movement for this entry:" & vbCrLf & vbCrLf & _
        "Yes = Office -> Site (mobilisation)" & vbCrLf & _
        "No = Site -> Office (demobilisation)" & vbCrLf & _
        "Cancel = do not create an entry", _
        vbYesNoCancel + vbQuestion, _
        "New project movement" _
    )
    If movementChoice = vbCancel Then GoTo CleanExit
    If movementChoice = vbYes Then
        movementText = "Office -> Site"
    Else
        movementText = "Site -> Office"
    End If

    baseName = Format$(Date, "yymmdd")
    newName = baseName
    counter = 1

    Do While SheetExists(newName)
        newName = baseName & "(" & counter & ")"
        counter = counter + 1
    Loop

    templateSheet.Copy _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set entrySheet = ActiveSheet
    entrySheet.Name = newName

    entrySheet.Range("E8").value = Date
    entrySheet.Range("E8").NumberFormat = "yyyy/mm/dd"

    With entrySheet.Range("E11").Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="Office -> Site,Site -> Office"
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Choose movement"
        .errorMessage = _
            "Select Office -> Site or Site -> Office."
    End With
    entrySheet.Range("E11").Value2 = movementText

    entrySheet.Activate

CleanExit:
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating

    If movementChoice = vbCancel Then Exit Sub
    'Refresh the sheet that was just created.  Do not search for an older
    'date-named sheet: a brand-new workbook may not have one.
    RefreshEntryAssets entrySheet
    Exit Sub

CleanFail:
    errorMessage = Err.description
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "The new entry could not be created." & vbCrLf & _
           errorMessage, vbExclamation
End Sub

Public Sub RefreshLatestEntryAssets()
    Dim entrySheet As Worksheet

    Set entrySheet = LatestDateEntrySheet(ThisWorkbook)
    If entrySheet Is Nothing Then
        MsgBox "No valid date-named entry sheet was found.", vbExclamation
        Exit Sub
    End If

    RefreshEntryAssets entrySheet
End Sub

Public Sub RefreshCurrentEntryAssets()
    If ActiveWorkbook Is Nothing Then Exit Sub
    If Not ActiveWorkbook Is ThisWorkbook Then
        MsgBox "Select the project entry sheet first, then run this again.", _
               vbExclamation
        Exit Sub
    End If

    If ActiveSheet.Name = "Template" Or ActiveSheet.Name = "Front Cover" Then
        MsgBox "Select the generated entry sheet first.", vbExclamation
        Exit Sub
    End If

    RefreshEntryAssets ActiveSheet
End Sub

Private Sub RefreshEntryAssets(ByVal entrySheet As Worksheet)
    Dim cacheSheet As Worksheet
    Dim masterBook As Workbook
    Dim masterSheet As Worksheet
    Dim leadCableSheet As Worksheet
    Dim masterPath As String
    Dim masterWasAlreadyOpen As Boolean
    Dim assetsByCategory As Object
    Dim projectLocations As Object
    Dim direction As String
    Dim projectCode As String
    Dim projectName As String
    Dim resolvedProjectName As String
    Dim populatedCount As Long
    Dim omittedCount As Long
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim errorMessage As String

    On Error GoTo CleanFail

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    direction = NormalisedDirection(entrySheet.Range("E11").Value2)
    projectCode = Trim$(CStr(entrySheet.Range("E4").Value2))
    projectName = Trim$(CStr(entrySheet.Range("G4").Value2))

    If direction <> "office -> site" And direction <> "site -> office" Then
        Err.Raise vbObjectError + 1003, , _
            "E11 must be Office -> Site or Site -> Office."
    End If
    If direction = "site -> office" And Len(projectCode) = 0 Then
        Err.Raise vbObjectError + 1004, , _
            "A Project Code is required in E4 for a demobilisation."
    End If

    Set masterBook = OpenMasterTrackerFromLink(masterWasAlreadyOpen)
    If masterBook Is Nothing Then
        Err.Raise vbObjectError + 1001, , _
            "The Master Tracker could not be opened from its web link:" & _
            vbCrLf & MasterTrackerLink() & vbCrLf & vbCrLf & _
            "Check the link in the ""A2AM Master Tracker"" row of the " & _
            SITE_DETAILS_SHEET & " sheet and that you are signed in " & _
            "to SharePoint in Excel."
    End If
    Set masterSheet = masterBook.Worksheets(MASTER_SHEET)
    Set leadCableSheet = masterBook.Worksheets(LEAD_CABLE_SHEET)
    If direction = "office -> site" Then
        Set assetsByCategory = ReadAvailableAssets(masterSheet)
        AddAvailableLeadCables leadCableSheet, assetsByCategory
    Else
        Set projectLocations = ProjectLocationKeys( _
            leadCableSheet, projectCode, projectName, resolvedProjectName _
        )
        If Len(resolvedProjectName) > 0 Then
            projectName = resolvedProjectName
            entrySheet.Range("G4").Value2 = resolvedProjectName
        End If
        Set assetsByCategory = ReadProjectEquipment( _
            masterSheet, projectLocations _
        )
        AddProjectLeadCables _
            leadCableSheet, assetsByCategory, projectCode, projectLocations
        ClearEntryItemRows entrySheet
    End If
    Set cacheSheet = GetOrCreateCacheSheet(ThisWorkbook)

    WriteCacheAndValidation entrySheet, cacheSheet, assetsByCategory
    If direction = "site -> office" Then
        populatedCount = PopulateDemobEntry( _
            entrySheet, assetsByCategory, omittedCount _
        )
    End If
    entrySheet.Activate

    If direction = "site -> office" Then
        MsgBox populatedCount & _
               " asset(s) assigned to " & projectCode & _
               " were populated on " & entrySheet.Name & "." & _
               IIf(omittedCount > 0, vbCrLf & CStr(omittedCount) & _
                   " additional asset(s) could not fit below row " & _
                   LAST_ITEM_ROW & ".", vbNullString) & _
               vbCrLf & "No other entry sheets were changed.", _
               vbInformation
    Else
        MsgBox "Available assets refreshed for: " & _
               entrySheet.Name & vbCrLf & _
               "No other entry sheets were changed.", vbInformation
    End If

CleanExit:
    On Error Resume Next
    If Not masterBook Is Nothing Then
        If Not masterWasAlreadyOpen Then masterBook.Close SaveChanges:=False
    End If
    If Not cacheSheet Is Nothing Then cacheSheet.Visible = xlSheetVeryHidden
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Exit Sub

CleanFail:
    errorMessage = Err.description
    GoTo CleanFailExit

CleanFailExit:
    On Error Resume Next
    If Not masterBook Is Nothing Then
        If Not masterWasAlreadyOpen Then masterBook.Close SaveChanges:=False
    End If
    If Not cacheSheet Is Nothing Then cacheSheet.Visible = xlSheetVeryHidden
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    MsgBox "Available assets could not be refreshed." & vbCrLf & _
           errorMessage, vbExclamation
End Sub

Public Sub GenerateFormEntry()
    Dim entrySheet As Worksheet
    Dim logSheet As Worksheet
    Dim documentsRoot As String
    Dim queueRoot As String
    Dim pendingFolder As String
    Dim outputPath As String
    Dim projectCode As String
    Dim projectName As String
    Dim technician As String
    Dim vanId As String
    Dim direction As String
    Dim movementType As String
    Dim entryDate As String
    Dim assetId As String
    Dim equipmentName As String
    Dim category As String
    Dim confirmation As String
    Dim assetKey As String
    Dim assetJson As String
    Dim payload As String
    Dim keysToLog As Collection
    Dim targetRow As Long
    Dim assetCount As Long
    Dim confirmedCount As Long
    Dim alreadyQueuedCount As Long
    Dim fileNumber As Integer
    Dim key As Variant
    Dim itemColumn As Long
    Dim assetIdColumn As Long
    Dim equipmentNameColumn As Long
    Dim confirmationColumn As Long
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim errorMessage As String
    Dim activeSheetDate As Date

    On Error GoTo CleanFail

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'The button belongs to an entry sheet, so process that sheet directly.
    'Only fall back to the latest dated entry when the macro is run elsewhere.
    If ActiveWorkbook Is ThisWorkbook Then
        If TryDateSheetName(ActiveSheet.Name, activeSheetDate) Then
            Set entrySheet = ActiveSheet
        End If
    End If
    If entrySheet Is Nothing Then
        Set entrySheet = LatestDateEntrySheet(ThisWorkbook)
    End If
    If entrySheet Is Nothing Then
        Err.Raise vbObjectError + 1100, , _
            "No valid date-named entry sheet was found."
    End If

    projectCode = Trim$(CStr(entrySheet.Range("E4").Value2))
    projectName = Trim$(CStr(entrySheet.Range("G4").Value2))
    technician = Trim$(CStr(entrySheet.Range("E9").Value2))
    vanId = Trim$(CStr(entrySheet.Range("E10").Value2))
    direction = NormalisedDirection(entrySheet.Range("E11").Value2)

    itemColumn = EntryColumnByHeader(entrySheet, "Item", 4)
    assetIdColumn = EntryColumnByHeader(entrySheet, "Asset ID", 5)
    equipmentNameColumn = EntryColumnByHeader( _
        entrySheet, "Equipment Name", 6 _
    )
    confirmationColumn = EntryColumnByHeader( _
        entrySheet, "Confirmed master tracker?", 7 _
    )

    If Len(projectCode) = 0 Then
        Err.Raise vbObjectError + 1101, , _
            "The latest entry sheet has no Project Number in E4."
    End If
    If direction <> "office -> site" And direction <> "site -> office" Then
        Err.Raise vbObjectError + 1102, , _
            "The movement in E11 must be Office -> Site or Site -> Office."
    End If
    If direction = "site -> office" Then
        movementType = "Demob"
    Else
        movementType = "Mob"
    End If

    If IsDate(entrySheet.Range("E8").value) Then
        entryDate = Format$(CDate(entrySheet.Range("E8").value), "yyyy-mm-dd")
    Else
        entryDate = Format$(Date, "yyyy-mm-dd")
    End If

    Set logSheet = GetOrCreateSubmissionLog(ThisWorkbook)
    Set keysToLog = New Collection

    For targetRow = FIRST_ITEM_ROW To LAST_ITEM_ROW
        confirmation = LCase$(Trim$(CStr( _
            entrySheet.Cells(targetRow, confirmationColumn).Value2 _
        )))
        assetId = Trim$(CStr( _
            entrySheet.Cells(targetRow, assetIdColumn).Value2 _
        ))

        If confirmation = "yes" And Len(assetId) > 0 Then
            confirmedCount = confirmedCount + 1
            category = CanonicalCategory( _
                entrySheet.Cells(targetRow, itemColumn).Value2 _
            )
            equipmentName = Trim$(CStr( _
                entrySheet.Cells(targetRow, equipmentNameColumn).Value2 _
            ))
            assetKey = entrySheet.Name & "|" & direction & "|" & UCase$(assetId)

            If Not SubmissionAlreadyLogged(logSheet, assetKey) Then
                If assetCount > 0 Then assetJson = assetJson & ","
                assetJson = assetJson & _
                    "{""category"":""" & JsonEscape(category) & """," & _
                    """assetId"":""" & JsonEscape(assetId) & """," & _
                    """equipmentName"":""" & JsonEscape(equipmentName) & """}"
                keysToLog.Add assetKey
                assetCount = assetCount + 1
            Else
                alreadyQueuedCount = alreadyQueuedCount + 1
            End If
        End If
    Next targetRow

    If assetCount = 0 Then
        If confirmedCount = 0 Then
            MsgBox "There are no rows marked Yes with an Asset ID on '" & _
                   entrySheet.Name & "'.", vbInformation
        Else
            MsgBox alreadyQueuedCount & " confirmed row(s) already have " & _
                   "an existing JSON queue file." & vbCrLf & _
                   "No duplicate file was created.", vbInformation
        End If
        GoTo CleanExit
    End If

    documentsRoot = LocalDocumentsRoot(ThisWorkbook.Path)
    queueRoot = documentsRoot & "Apps\ProjectMobQueue"
    pendingFolder = queueRoot & "\Pending"
    EnsureFolderExists queueRoot
    EnsureFolderExists pendingFolder

    Randomize
    outputPath = pendingFolder & "\" & _
        SafeFilePart(projectCode) & "_" & entrySheet.Name & "_" & _
        Format$(Now, "yyyymmdd_hhnnss") & "_" & _
        Format$(Int(Rnd() * 10000), "0000") & ".json"

    payload = _
        "{""submissionId"":""" & _
            JsonEscape(projectCode & "|" & entrySheet.Name & "|" & _
                Format$(Now, "yyyymmddhhnnss")) & """," & _
        """sourceWorkbook"":""" & JsonEscape(ThisWorkbook.Name) & """," & _
        """sheetName"":""" & JsonEscape(entrySheet.Name) & """," & _
        """entryDate"":""" & JsonEscape(entryDate) & """," & _
        """submittedAt"":""" & Format$(Now, "yyyy-mm-dd\THH:nn:ss") & """," & _
        """technician"":""" & JsonEscape(technician) & """," & _
        """vanId"":""" & JsonEscape(vanId) & """," & _
        """direction"":""" & JsonEscape(direction) & """," & _
        """movementType"":""" & JsonEscape(movementType) & """," & _
        """projectCode"":""" & JsonEscape(projectCode) & """," & _
        """projectName"":""" & JsonEscape(projectName) & """," & _
        """assets"":[" & assetJson & "]}"

    fileNumber = FreeFile
    Open outputPath For Output As #fileNumber
    Print #fileNumber, payload
    Close #fileNumber
    fileNumber = 0

    For Each key In keysToLog
        LogSubmission logSheet, CStr(key), outputPath
    Next key
    logSheet.Visible = xlSheetVeryHidden

    MsgBox assetCount & " confirmed asset movement(s) were queued from " & _
           entrySheet.Name & "." & vbCrLf & _
           "All earlier dated sheets were left unchanged.", vbInformation

CleanExit:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    If Not logSheet Is Nothing Then logSheet.Visible = xlSheetVeryHidden
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Exit Sub

CleanFail:
    errorMessage = Err.description
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    If Not logSheet Is Nothing Then logSheet.Visible = xlSheetVeryHidden
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    MsgBox "The form entry could not be generated." & vbCrLf & _
           errorMessage, vbExclamation
End Sub

Private Function LatestDateEntrySheet(ByVal book As Workbook) As Worksheet
    Dim sheet As Worksheet
    Dim sheetDate As Date
    Dim latestDate As Date
    Dim latestIndex As Long
    Dim found As Boolean

    For Each sheet In book.Worksheets
        If TryDateSheetName(sheet.Name, sheetDate) Then
            If Not found Or sheetDate > latestDate Or _
               (sheetDate = latestDate And sheet.Index > latestIndex) Then
                latestDate = sheetDate
                latestIndex = sheet.Index
                Set LatestDateEntrySheet = sheet
                found = True
            End If
        End If
    Next sheet
End Function

Private Function TryDateSheetName( _
    ByVal sheetName As String, _
    ByRef parsedDate As Date _
) As Boolean
    Dim yearNumber As Long
    Dim monthNumber As Long
    Dim dayNumber As Long
    Dim candidate As Date
    Dim datePart As String
    Dim suffix As String

    If Len(sheetName) < 6 Then Exit Function
    datePart = Left$(sheetName, 6)
    If Not datePart Like "######" Then Exit Function

    suffix = Mid$(sheetName, 7)
    If Len(suffix) > 0 Then
        If Left$(suffix, 1) <> "(" Or Right$(suffix, 1) <> ")" Then
            Exit Function
        End If
        If Len(suffix) < 3 Then Exit Function
        If Not (Mid$(suffix, 2, Len(suffix) - 2) Like _
            String$(Len(suffix) - 2, "#")) Then
            Exit Function
        End If
    End If

    dayNumber = CLng(Left$(datePart, 2))
    monthNumber = CLng(Mid$(datePart, 3, 2))
    yearNumber = 2000 + CLng(Right$(datePart, 2))

    On Error GoTo InvalidDate
    candidate = DateSerial(yearNumber, monthNumber, dayNumber)
    On Error GoTo 0

    If Year(candidate) <> yearNumber Then Exit Function
    If Month(candidate) <> monthNumber Then Exit Function
    If Day(candidate) <> dayNumber Then Exit Function

    parsedDate = candidate
    TryDateSheetName = True
    Exit Function

InvalidDate:
    On Error GoTo 0
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet

    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not sheet Is Nothing
End Function

Private Function MasterTrackerPath(ByVal projectWorkbookPath As String) As String
    Const ROOT_MARKER As String = "\A2 Advanced Monitoring - Documents\"
    Dim markerPosition As Long
    Dim documentsRoot As String
    Dim candidate As String
    Dim oneDriveRoot As String
    Dim userProfile As String

    markerPosition = InStr(1, projectWorkbookPath, ROOT_MARKER, vbTextCompare)
    If markerPosition > 0 Then
        documentsRoot = Left$( _
            projectWorkbookPath, _
            markerPosition + Len(ROOT_MARKER) - 1 _
        )
        candidate = documentsRoot & MASTER_FILE
        If Len(Dir$(candidate)) > 0 Then
            MasterTrackerPath = candidate
            Exit Function
        End If
    End If

    oneDriveRoot = Trim$(Environ$("OneDriveCommercial"))
    If Len(oneDriveRoot) > 0 Then
        candidate = oneDriveRoot & _
            "\A2 Advanced Monitoring - Documents\" & MASTER_FILE
        If Len(Dir$(candidate)) > 0 Then
            MasterTrackerPath = candidate
            Exit Function
        End If
    End If

    userProfile = Trim$(Environ$("USERPROFILE"))
    If Len(userProfile) > 0 Then
        candidate = userProfile & _
            "\A2 Advanced Monitoring\A2 Advanced Monitoring - Documents\" & _
            MASTER_FILE
        If Len(Dir$(candidate)) > 0 Then
            MasterTrackerPath = candidate
            Exit Function
        End If

        candidate = userProfile & _
            "\A2 Advanced Monitoring - Documents\" & MASTER_FILE
        If Len(Dir$(candidate)) > 0 Then
            MasterTrackerPath = candidate
            Exit Function
        End If
    End If

    Err.Raise vbObjectError + 1002, , _
        "The locally synced Master Tracker could not be found. Expected:" & _
        vbCrLf & userProfile & _
        "\A2 Advanced Monitoring\A2 Advanced Monitoring - Documents\" & _
        MASTER_FILE & vbCrLf & vbCrLf & _
        "Make sure that file is available on this device in OneDrive."
End Function

Private Function LocalDocumentsRoot(ByVal projectWorkbookPath As String) As String
    Const ROOT_MARKER As String = "\A2 Advanced Monitoring - Documents\"
    Dim markerPosition As Long
    Dim candidate As String
    Dim oneDriveRoot As String
    Dim userProfile As String

    markerPosition = InStr(1, projectWorkbookPath, ROOT_MARKER, vbTextCompare)
    If markerPosition > 0 Then
        candidate = Left$( _
            projectWorkbookPath, _
            markerPosition + Len(ROOT_MARKER) - 1 _
        )
        If Len(Dir$(candidate, vbDirectory)) > 0 Then
            LocalDocumentsRoot = candidate
            Exit Function
        End If
    End If

    oneDriveRoot = Trim$(Environ$("OneDriveCommercial"))
    If Len(oneDriveRoot) > 0 Then
        candidate = oneDriveRoot & "\A2 Advanced Monitoring - Documents\"
        If Len(Dir$(candidate, vbDirectory)) > 0 Then
            LocalDocumentsRoot = candidate
            Exit Function
        End If
    End If

    userProfile = Trim$(Environ$("USERPROFILE"))
    If Len(userProfile) > 0 Then
        candidate = userProfile & _
            "\A2 Advanced Monitoring\A2 Advanced Monitoring - Documents\"
        If Len(Dir$(candidate, vbDirectory)) > 0 Then
            LocalDocumentsRoot = candidate
            Exit Function
        End If
    End If

    Err.Raise vbObjectError + 1103, , _
        "The local synced 'A2 Advanced Monitoring - Documents' folder " & _
        "could not be found."
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function GetOrCreateSubmissionLog(ByVal book As Workbook) As Worksheet
    On Error Resume Next
    Set GetOrCreateSubmissionLog = book.Worksheets(SUBMISSION_LOG_SHEET)
    On Error GoTo 0

    If GetOrCreateSubmissionLog Is Nothing Then
        Set GetOrCreateSubmissionLog = book.Worksheets.Add( _
            After:=book.Worksheets(book.Worksheets.Count) _
        )
        GetOrCreateSubmissionLog.Name = SUBMISSION_LOG_SHEET
        GetOrCreateSubmissionLog.Range("A1").value = "Submission Key"
        GetOrCreateSubmissionLog.Range("B1").value = "Queued At"
        GetOrCreateSubmissionLog.Range("C1").value = "Queue File"
    End If

    GetOrCreateSubmissionLog.Visible = xlSheetVeryHidden
End Function

Private Function SubmissionAlreadyLogged( _
    ByVal logSheet As Worksheet, _
    ByVal submissionKey As String _
) As Boolean
    Dim lastRow As Long
    Dim matchCell As Range
    Dim searchRange As Range
    Dim firstAddress As String
    Dim queueFile As String

    lastRow = logSheet.Cells(logSheet.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Set searchRange = logSheet.Range("A2:A" & lastRow)
    Set matchCell = searchRange.Find( _
        What:=submissionKey, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False _
    )

    If matchCell Is Nothing Then Exit Function
    firstAddress = matchCell.address

    Do
        queueFile = Trim$(CStr( _
            logSheet.Cells(matchCell.Row, "C").Value2 _
        ))

        'A stale hidden-log row must not block a new submission.
        'Only treat it as already queued while its JSON file still exists.
        If Len(queueFile) > 0 Then
            If Len(Dir$(queueFile)) > 0 Then
                SubmissionAlreadyLogged = True
                Exit Function
            End If
        End If

        Set matchCell = searchRange.FindNext(matchCell)
        If matchCell Is Nothing Then Exit Do
    Loop While matchCell.address <> firstAddress
End Function

Private Sub LogSubmission( _
    ByVal logSheet As Worksheet, _
    ByVal submissionKey As String, _
    ByVal queueFile As String _
)
    Dim nextRow As Long
    nextRow = logSheet.Cells(logSheet.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    logSheet.Cells(nextRow, "A").Value2 = submissionKey
    logSheet.Cells(nextRow, "B").value = Now
    logSheet.Cells(nextRow, "B").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    logSheet.Cells(nextRow, "C").Value2 = queueFile
End Sub

Private Function JsonEscape(ByVal value As String) As String
    Dim escaped As String
    escaped = Replace(value, "\", "\\")
    escaped = Replace(escaped, Chr$(34), "\" & Chr$(34))
    escaped = Replace(escaped, vbCrLf, "\n")
    escaped = Replace(escaped, vbCr, "\n")
    escaped = Replace(escaped, vbLf, "\n")
    escaped = Replace(escaped, vbTab, "\t")
    JsonEscape = escaped
End Function

Private Function SafeFilePart(ByVal value As String) As String
    Dim invalidCharacters As Variant
    Dim character As Variant
    Dim result As String

    result = Trim$(value)
    invalidCharacters = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each character In invalidCharacters
        result = Replace(result, CStr(character), "_")
    Next character
    SafeFilePart = result
End Function

Private Function IsWebAddress(ByVal value As String) As Boolean
    Dim normalised As String
    normalised = LCase$(Trim$(value))
    IsWebAddress = _
        Left$(normalised, 8) = "https://" Or _
        Left$(normalised, 7) = "http://"
End Function

' Reads the Master Tracker web link from the Site Details sheet: the
' row labelled "A2AM Master Tracker", first cell to the label's right
' holding a hyperlink or web address. Falls back to the MASTER_LINK
' constant so the macro still works if that row is deleted.
Private Function MasterTrackerLink() As String
    Dim detailsSheet As Worksheet
    Dim labelCell As Range
    Dim valueCell As Range
    Dim lastColumn As Long
    Dim col As Long
    Dim linkText As String

    On Error Resume Next
    Set detailsSheet = ThisWorkbook.Worksheets(SITE_DETAILS_SHEET)
    On Error GoTo 0

    If Not detailsSheet Is Nothing Then
        Set labelCell = detailsSheet.UsedRange.Find( _
            What:=MASTER_LINK_LABEL, LookIn:=xlValues, _
            LookAt:=xlPart, MatchCase:=False)
        If Not labelCell Is Nothing Then
            lastColumn = detailsSheet.UsedRange.Column + _
                         detailsSheet.UsedRange.Columns.Count - 1
            For col = labelCell.Column + 1 To lastColumn
                Set valueCell = detailsSheet.Cells(labelCell.Row, col)
                If valueCell.Hyperlinks.Count > 0 Then
                    linkText = valueCell.Hyperlinks(1).address
                Else
                    linkText = Trim$(CStr(valueCell.value))
                End If
                If IsWebAddress(linkText) Then Exit For
                linkText = vbNullString
            Next col
        End If
    End If

    If Len(linkText) = 0 Then linkText = MASTER_LINK
    MasterTrackerLink = linkText
End Function

' Opens the Master Tracker from its web link. A workbook that is
' already open and contains the Equipment and Lead Cables sheets is
' reused. Share links are tried with a "?download=1" suffix first,
' which makes SharePoint serve the real file instead of a redirect
' page; every candidate is validated by the presence of the Equipment
' sheet. Returns Nothing when no candidate yields the tracker.
Private Function OpenMasterTrackerFromLink( _
    ByRef wasAlreadyOpen As Boolean _
) As Workbook
    Dim book As Workbook
    Dim linkAddress As String
    Dim candidates As Variant
    Dim i As Long
    Dim queryPos As Long

    For Each book In Application.Workbooks
        If Not book Is ThisWorkbook Then
            If BookHasSheet(book, MASTER_SHEET) And _
               BookHasSheet(book, LEAD_CABLE_SHEET) Then
                Set OpenMasterTrackerFromLink = book
                wasAlreadyOpen = True
                Exit Function
            End If
        End If
    Next book

    linkAddress = Trim$(MasterTrackerLink())
    queryPos = InStr(1, linkAddress, "?")
    If queryPos > 0 Then linkAddress = Left$(linkAddress, queryPos - 1)

    candidates = Array(linkAddress & "?download=1", linkAddress)
    For i = LBound(candidates) To UBound(candidates)
        Set book = Nothing
        On Error Resume Next
        Set book = Workbooks.Open( _
            fileName:=CStr(candidates(i)), _
            UpdateLinks:=False, _
            ReadOnly:=True, _
            AddToMru:=False _
        )
        On Error GoTo 0

        If Not book Is Nothing Then
            If BookHasSheet(book, MASTER_SHEET) Then
                Set OpenMasterTrackerFromLink = book
                wasAlreadyOpen = False
                Exit Function
            End If
            'Wrong content (e.g. a share-link redirect page opened as
            'a workbook) - discard and try the next candidate.
            book.Close SaveChanges:=False
        End If
    Next i
End Function

Private Function BookHasSheet( _
    ByVal book As Workbook, _
    ByVal sheetName As String _
) As Boolean
    Dim candidate As Worksheet

    On Error Resume Next
    Set candidate = book.Worksheets(sheetName)
    On Error GoTo 0

    BookHasSheet = Not candidate Is Nothing
End Function

Private Function OpenWorkbookReadOnly( _
    ByVal fullPath As String, _
    ByRef wasAlreadyOpen As Boolean _
) As Workbook
    Dim book As Workbook

    For Each book In Application.Workbooks
        If StrComp(book.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenWorkbookReadOnly = book
            wasAlreadyOpen = True
            Exit Function
        End If
    Next book

    Set OpenWorkbookReadOnly = Workbooks.Open( _
        fileName:=fullPath, _
        UpdateLinks:=False, _
        ReadOnly:=True, _
        AddToMru:=False _
    )
End Function

Private Function ReadAvailableAssets(ByVal equipmentSheet As Worksheet) As Object
    Dim categoryAssets As Object
    Dim assets As Collection
    Dim lastRow As Long
    Dim sourceRow As Long
    Dim currentCategory As String
    Dim rawCategory As String
    Dim assetId As String
    Dim equipmentName As String
    Dim location As String

    Set categoryAssets = CreateObject("Scripting.Dictionary")
    categoryAssets.CompareMode = vbTextCompare

    lastRow = equipmentSheet.Cells( _
        equipmentSheet.Rows.Count, "B" _
    ).End(xlUp).Row

    For sourceRow = 2 To lastRow
        rawCategory = Trim$(CStr(equipmentSheet.Cells(sourceRow, "A").Value2))
        If Len(rawCategory) > 0 Then
            currentCategory = CanonicalCategory(rawCategory)
        End If

        assetId = Trim$(CStr(equipmentSheet.Cells(sourceRow, "B").Value2))
        equipmentName = Trim$(CStr(equipmentSheet.Cells(sourceRow, "C").Value2))
        location = Trim$(CStr(equipmentSheet.Cells(sourceRow, "D").Value2))

        If Len(currentCategory) > 0 And Len(assetId) > 0 Then
            If Len(location) = 0 Or StrComp(location, "Office", vbTextCompare) = 0 Then
                If Not categoryAssets.Exists(currentCategory) Then
                    Set assets = New Collection
                    categoryAssets.Add currentCategory, assets
                End If
                Set assets = categoryAssets(currentCategory)
                assets.Add Array(assetId, equipmentName, location)
            End If
        End If
    Next sourceRow

    Set ReadAvailableAssets = categoryAssets
End Function

Private Function ProjectLocationKeys( _
    ByVal leadCableSheet As Worksheet, _
    ByVal projectCode As String, _
    ByVal projectName As String, _
    ByRef resolvedProjectName As String _
) As Object
    Dim locationKeys As Object
    Dim lastRow As Long
    Dim sourceRow As Long
    Dim rowProjectCode As String
    Dim location As String
    Dim locationKey As String

    Set locationKeys = CreateObject("Scripting.Dictionary")
    locationKeys.CompareMode = vbTextCompare

    locationKey = NormaliseMatchKey(projectName)
    If Len(locationKey) > 0 Then
        locationKeys(locationKey) = True
        resolvedProjectName = projectName
        Set ProjectLocationKeys = locationKeys
        Exit Function
    End If

    lastRow = leadCableSheet.Cells( _
        leadCableSheet.Rows.Count, "B" _
    ).End(xlUp).Row

    For sourceRow = 2 To lastRow
        rowProjectCode = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "G").Value2 _
        ))
        If StrComp(rowProjectCode, projectCode, vbTextCompare) = 0 Then
            location = Trim$(CStr( _
                leadCableSheet.Cells(sourceRow, "F").Value2 _
            ))
            locationKey = NormaliseMatchKey(location)
            If Len(locationKey) > 0 Then
                locationKeys(locationKey) = True
                If Len(resolvedProjectName) = 0 Then
                    resolvedProjectName = location
                End If
            End If
        End If
    Next sourceRow

    Set ProjectLocationKeys = locationKeys
End Function

Private Function ReadProjectEquipment( _
    ByVal equipmentSheet As Worksheet, _
    ByVal projectLocations As Object _
) As Object
    Dim categoryAssets As Object
    Dim assets As Collection
    Dim lastRow As Long
    Dim sourceRow As Long
    Dim currentCategory As String
    Dim rawCategory As String
    Dim assetId As String
    Dim equipmentName As String
    Dim location As String
    Dim locationKey As String

    Set categoryAssets = CreateObject("Scripting.Dictionary")
    categoryAssets.CompareMode = vbTextCompare

    lastRow = equipmentSheet.Cells( _
        equipmentSheet.Rows.Count, "B" _
    ).End(xlUp).Row

    For sourceRow = 2 To lastRow
        rawCategory = Trim$(CStr( _
            equipmentSheet.Cells(sourceRow, "A").Value2 _
        ))
        If Len(rawCategory) > 0 Then
            currentCategory = CanonicalCategory(rawCategory)
        End If

        assetId = Trim$(CStr( _
            equipmentSheet.Cells(sourceRow, "B").Value2 _
        ))
        equipmentName = Trim$(CStr( _
            equipmentSheet.Cells(sourceRow, "C").Value2 _
        ))
        location = Trim$(CStr( _
            equipmentSheet.Cells(sourceRow, "D").Value2 _
        ))
        locationKey = NormaliseMatchKey(location)

        If Len(currentCategory) > 0 And Len(assetId) > 0 Then
            If Len(locationKey) > 0 And _
               projectLocations.Exists(locationKey) Then
                If Not categoryAssets.Exists(currentCategory) Then
                    Set assets = New Collection
                    categoryAssets.Add currentCategory, assets
                End If
                Set assets = categoryAssets(currentCategory)
                assets.Add Array(assetId, equipmentName, location)
            End If
        End If
    Next sourceRow

    Set ReadProjectEquipment = categoryAssets
End Function

Private Sub AddAvailableLeadCables( _
    ByVal leadCableSheet As Worksheet, _
    ByVal categoryAssets As Object _
)
    Const LEAD_CATEGORY As String = "Lead cable"
    Dim assets As Collection
    Dim lastRow As Long
    Dim sourceRow As Long
    Dim leadType As String
    Dim assetId As String
    Dim sensors As String
    Dim cableLength As String
    Dim location As String
    Dim condition As String
    Dim description As String

    If Not categoryAssets.Exists(LEAD_CATEGORY) Then
        Set assets = New Collection
        categoryAssets.Add LEAD_CATEGORY, assets
    End If
    Set assets = categoryAssets(LEAD_CATEGORY)

    lastRow = leadCableSheet.Cells( _
        leadCableSheet.Rows.Count, "B" _
    ).End(xlUp).Row

    For sourceRow = 2 To lastRow
        leadType = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "A").Value2 _
        ))
        assetId = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "B").Value2 _
        ))
        sensors = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "C").Value2 _
        ))
        cableLength = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "D").Value2 _
        ))
        location = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "F").Value2 _
        ))
        condition = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "H").Value2 _
        ))

        If Len(assetId) > 0 Then
            If Len(location) = 0 Or _
               StrComp(location, "Office", vbTextCompare) = 0 Then
                description = LeadCableDescription( _
                    leadType, sensors, cableLength, condition _
                )
                assets.Add Array(assetId, description, location)
            End If
        End If
    Next sourceRow
End Sub

Private Sub AddProjectLeadCables( _
    ByVal leadCableSheet As Worksheet, _
    ByVal categoryAssets As Object, _
    ByVal projectCode As String, _
    ByVal projectLocations As Object _
)
    Const LEAD_CATEGORY As String = "Lead cable"
    Dim assets As Collection
    Dim lastRow As Long
    Dim sourceRow As Long
    Dim leadType As String
    Dim assetId As String
    Dim sensors As String
    Dim cableLength As String
    Dim location As String
    Dim rowProjectCode As String
    Dim condition As String
    Dim description As String
    Dim locationKey As String
    Dim belongsToProject As Boolean

    If Not categoryAssets.Exists(LEAD_CATEGORY) Then
        Set assets = New Collection
        categoryAssets.Add LEAD_CATEGORY, assets
    End If
    Set assets = categoryAssets(LEAD_CATEGORY)

    lastRow = leadCableSheet.Cells( _
        leadCableSheet.Rows.Count, "B" _
    ).End(xlUp).Row

    For sourceRow = 2 To lastRow
        leadType = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "A").Value2 _
        ))
        assetId = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "B").Value2 _
        ))
        sensors = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "C").Value2 _
        ))
        cableLength = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "D").Value2 _
        ))
        location = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "F").Value2 _
        ))
        rowProjectCode = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "G").Value2 _
        ))
        condition = Trim$(CStr( _
            leadCableSheet.Cells(sourceRow, "H").Value2 _
        ))

        locationKey = NormaliseMatchKey(location)
        belongsToProject = False
        If Len(locationKey) > 0 Then
            belongsToProject = projectLocations.Exists(locationKey)
        End If
        If Not belongsToProject And projectLocations.Count = 0 Then
            belongsToProject = _
                StrComp(rowProjectCode, projectCode, vbTextCompare) = 0
        End If

        If Len(assetId) > 0 And belongsToProject Then
            description = LeadCableDescription( _
                leadType, sensors, cableLength, condition _
            )
            assets.Add Array(assetId, description, location)
        End If
    Next sourceRow
End Sub

Private Function LeadCableDescription( _
    ByVal leadType As String, _
    ByVal sensors As String, _
    ByVal cableLength As String, _
    ByVal condition As String _
) As String
    Dim description As String

    description = leadType
    If Len(sensors) > 0 Then
        If Len(description) > 0 Then description = description & " - "
        description = description & sensors & " sensors"
    End If
    If Len(cableLength) > 0 Then
        If Len(description) > 0 Then description = description & " - "
        description = description & cableLength & "m"
    End If
    If Len(condition) > 0 Then
        If Len(description) > 0 Then description = description & " - "
        description = description & condition
    End If

    If Len(description) = 0 Then description = "Lead cable"
    LeadCableDescription = description
End Function

Private Function PopulateDemobEntry( _
    ByVal entrySheet As Worksheet, _
    ByVal categoryAssets As Object, _
    ByRef omittedCount As Long _
) As Long
    Dim categories As Variant
    Dim categoryIndex As Long
    Dim category As String
    Dim assets As Collection
    Dim asset As Variant
    Dim targetRow As Long

    categories = Array( _
        "Analyser", "Laptop", "Splicer", "Fibre Cleaver", _
        "Drumstand", "Gun", "Battery", "Tent", "Table", _
        "Transformer", "Internet", "DeWalt Tower", "Lead cable" _
    )
    targetRow = FIRST_ITEM_ROW

    For categoryIndex = LBound(categories) To UBound(categories)
        category = CStr(categories(categoryIndex))
        If categoryAssets.Exists(category) Then
            Set assets = categoryAssets(category)
            For Each asset In assets
                If targetRow <= LAST_ITEM_ROW Then
                    entrySheet.Cells(targetRow, "D").Value2 = category
                    entrySheet.Cells(targetRow, "E").Value2 = asset(0)
                    ClearCellOrMergedArea _
                        entrySheet.Cells(targetRow, "G")
                    targetRow = targetRow + 1
                    PopulateDemobEntry = PopulateDemobEntry + 1
                Else
                    omittedCount = omittedCount + 1
                End If
            Next asset
        End If
    Next categoryIndex
End Function

Private Function EntryColumnByHeader( _
    ByVal entrySheet As Worksheet, _
    ByVal headerText As String, _
    ByVal fallbackColumn As Long _
) As Long
    Dim targetKey As String
    Dim candidateKey As String
    Dim headerColumn As Long
    Dim lastHeaderColumn As Long

    targetKey = NormaliseMatchKey(headerText)
    lastHeaderColumn = entrySheet.Cells( _
        ITEM_HEADER_ROW, entrySheet.Columns.Count _
    ).End(xlToLeft).Column
    If lastHeaderColumn < 10 Then lastHeaderColumn = 10

    For headerColumn = 1 To lastHeaderColumn
        candidateKey = NormaliseMatchKey( _
            entrySheet.Cells(ITEM_HEADER_ROW, headerColumn).Value2 _
        )
        If candidateKey = targetKey Then
            EntryColumnByHeader = headerColumn
            Exit Function
        End If
    Next headerColumn

    EntryColumnByHeader = fallbackColumn
End Function

Private Function NormaliseMatchKey(ByVal rawValue As Variant) As String
    Dim value As String
    Dim character As String
    Dim result As String
    Dim position As Long

    value = LCase$(Trim$(CStr(rawValue)))
    For position = 1 To Len(value)
        character = Mid$(value, position, 1)
        If character Like "[a-z0-9]" Then
            result = result & character
        End If
    Next position

    NormaliseMatchKey = result
End Function

Private Function GetOrCreateCacheSheet(ByVal book As Workbook) As Worksheet
    On Error Resume Next
    Set GetOrCreateCacheSheet = book.Worksheets(CACHE_SHEET)
    On Error GoTo 0

    If GetOrCreateCacheSheet Is Nothing Then
        Set GetOrCreateCacheSheet = book.Worksheets.Add( _
            After:=book.Worksheets(book.Worksheets.Count) _
        )
        GetOrCreateCacheSheet.Name = CACHE_SHEET
    End If

    GetOrCreateCacheSheet.Visible = xlSheetVisible
End Function

Private Sub WriteCacheAndValidation( _
    ByVal entrySheet As Worksheet, _
    ByVal cacheSheet As Worksheet, _
    ByVal categoryAssets As Object _
)
    Dim targetRow As Long
    Dim cacheRow As Long
    Dim categoryIndex As Long
    Dim category As String
    Dim listName As String
    Dim listStartRow As Long
    Dim listEndRow As Long
    Dim finalCacheRow As Long
    Dim assets As Collection
    Dim asset As Variant
    Dim nameFormula As String
    Dim assetValidationFormula As String
    Dim categories As Variant
    Dim categoryValidationList As String
    Dim confirmationRange As Range

    cacheSheet.Cells.ClearContents
    cacheSheet.Range("A1").Value2 = "Category"
    cacheSheet.Range("B1").Value2 = "Asset ID"
    cacheSheet.Range("C1").Value2 = "Equipment Name"
    cacheSheet.Range("D1").Value2 = "Location"
    cacheRow = 2

    categories = Array( _
        "Analyser", "Laptop", "Splicer", "Fibre Cleaver", _
        "Drumstand", "Gun", "Battery", "Tent", "Table", _
        "Transformer", "Internet", "DeWalt Tower", "Lead cable" _
    )
    categoryValidationList = _
        "Analyser,Laptop,Splicer,Fibre Cleaver,Drumstand,Gun," & _
        "Battery,Tent,Table,Transformer,Internet,DeWalt Tower," & _
        "Lead cable"

    For categoryIndex = LBound(categories) To UBound(categories)
        category = CStr(categories(categoryIndex))
        listName = ValidationListName(category)

        DeleteWorkbookName ThisWorkbook, listName
        listStartRow = cacheRow

        If categoryAssets.Exists(category) Then
            Set assets = categoryAssets(category)
            For Each asset In assets
                cacheSheet.Cells(cacheRow, "A").Value2 = category
                cacheSheet.Cells(cacheRow, "B").Value2 = asset(0)
                cacheSheet.Cells(cacheRow, "C").Value2 = asset(1)
                cacheSheet.Cells(cacheRow, "D").Value2 = asset(2)
                cacheRow = cacheRow + 1
            Next asset
        End If

        If cacheRow = listStartRow Then
            cacheSheet.Cells(cacheRow, "A").Value2 = category
            cacheSheet.Cells(cacheRow, "B").Value2 = vbNullString
            cacheRow = cacheRow + 1
        End If

        listEndRow = cacheRow - 1
        ThisWorkbook.Names.Add _
            Name:=listName, _
            RefersTo:="='" & CACHE_SHEET & "'!$B$" & _
                listStartRow & ":$B$" & listEndRow
    Next categoryIndex

    ' Accept the US spelling already used on some entry sheets.
    DeleteWorkbookName ThisWorkbook, "Available_Analyzer"
    ThisWorkbook.Names.Add _
        Name:="Available_Analyzer", _
        RefersTo:="=" & ValidationListName("Analyser")

    cacheSheet.Cells(cacheRow, "A").Value2 = "Blank"
    cacheSheet.Cells(cacheRow, "B").Value2 = vbNullString
    DeleteWorkbookName ThisWorkbook, "Available_Blank"
    ThisWorkbook.Names.Add _
        Name:="Available_Blank", _
        RefersTo:="='" & CACHE_SHEET & "'!$B$" & cacheRow
    finalCacheRow = cacheRow

    For targetRow = FIRST_ITEM_ROW To LAST_ITEM_ROW
        With entrySheet.Cells(targetRow, "D").Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:=categoryValidationList
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
            .ErrorTitle = "Unknown category"
            .errorMessage = _
                "Select an equipment category from the dropdown list."
        End With

        assetValidationFormula = _
            "=INDIRECT(IF($D" & targetRow & "=""""," & _
            """Available_Blank"",""Available_""&" & _
            "SUBSTITUTE($D" & targetRow & ","" "",""_"")))"

        With entrySheet.Cells(targetRow, "E").Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:=assetValidationFormula
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
            .ErrorTitle = "Unavailable asset"
            .errorMessage = _
                "Select an available asset from the dropdown list."
        End With

        nameFormula = "=IF(E" & targetRow & "="""",""""," & _
            "IFERROR(XLOOKUP(E" & targetRow & "," & _
            "'" & CACHE_SHEET & "'!$B$2:$B$" & finalCacheRow & "," & _
            "'" & CACHE_SHEET & "'!$C$2:$C$" & finalCacheRow & ",""""),""""))"
        entrySheet.Cells(targetRow, "F").Formula2 = nameFormula

        Set confirmationRange = entrySheet.Cells(targetRow, "G")
        If confirmationRange.MergeCells Then
            Set confirmationRange = confirmationRange.MergeArea
        End If

        With confirmationRange.Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:="Yes,No"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    Next targetRow

    cacheSheet.Columns("A:D").AutoFit
    cacheSheet.Visible = xlSheetVeryHidden
End Sub

Private Sub ClearEntryItemRows(ByVal entrySheet As Worksheet)
    Dim targetRow As Long
    Dim targetColumn As Long

    For targetRow = FIRST_ITEM_ROW To LAST_ITEM_ROW
        For targetColumn = 4 To 7
            ClearCellOrMergedArea _
                entrySheet.Cells(targetRow, targetColumn)
        Next targetColumn
    Next targetRow
End Sub

Private Sub ClearCellOrMergedArea(ByVal targetCell As Range)
    If targetCell.MergeCells Then
        targetCell.MergeArea.ClearContents
    Else
        targetCell.ClearContents
    End If
End Sub

Private Sub DeleteWorkbookName(ByVal book As Workbook, ByVal nameText As String)
    On Error Resume Next
    book.Names(nameText).Delete
    On Error GoTo 0
End Sub

Private Function CanonicalCategory(ByVal rawValue As Variant) As String
    Dim value As String
    value = LCase$(Trim$(CStr(rawValue)))

    Select Case value
        Case "analyser", "analysers", "analyzer", "analyzers"
            CanonicalCategory = "Analyser"
        Case "laptop", "laptops"
            CanonicalCategory = "Laptop"
        Case "splicer", "splicers"
            CanonicalCategory = "Splicer"
        Case "fibre cleaver", "fibre cleavers", "cleaver", "cleavers"
            CanonicalCategory = "Fibre Cleaver"
        Case "drumstand", "drumstands", "drum stand", "drum stands"
            CanonicalCategory = "Drumstand"
        Case "gun", "guns"
            CanonicalCategory = "Gun"
        Case "battery", "batteries"
            CanonicalCategory = "Battery"
        Case "tent", "tents"
            CanonicalCategory = "Tent"
        Case "table", "tables", "splicing table", "splicing tables"
            CanonicalCategory = "Table"
        Case "transformer", "transformers"
            CanonicalCategory = "Transformer"
        Case "internet", "internet dongle", "internet dongles"
            CanonicalCategory = "Internet"
        Case "dewalt tower", "dewalt towers", "dewalt tower (toolbox)"
            CanonicalCategory = "DeWalt Tower"
        Case "lead cable", "lead cables", "lead"
            CanonicalCategory = "Lead cable"
    End Select
End Function

Private Function ValidationListName(ByVal category As String) As String
    Select Case category
        Case "Analyser": ValidationListName = "Available_Analyser"
        Case "Laptop": ValidationListName = "Available_Laptop"
        Case "Splicer": ValidationListName = "Available_Splicer"
        Case "Fibre Cleaver": ValidationListName = "Available_Fibre_Cleaver"
        Case "Drumstand": ValidationListName = "Available_Drumstand"
        Case "Gun": ValidationListName = "Available_Gun"
        Case "Battery": ValidationListName = "Available_Battery"
        Case "Tent": ValidationListName = "Available_Tent"
        Case "Table": ValidationListName = "Available_Table"
        Case "Transformer": ValidationListName = "Available_Transformer"
        Case "Internet": ValidationListName = "Available_Internet"
        Case "DeWalt Tower": ValidationListName = "Available_DeWalt_Tower"
        Case "Lead cable": ValidationListName = "Available_Lead_Cable"
    End Select
End Function

Private Function NormalisedDirection(ByVal rawValue As Variant) As String
    Dim value As String
    value = LCase$(Trim$(CStr(rawValue)))
    value = Replace(value, "‰ Õ", "->")
    value = Replace(value, "‰ÛÒ", "-")
    value = Replace(value, "‰ÛÓ", "-")
    value = Replace(value, " ", vbNullString)

    If value = "office->site" Then
        NormalisedDirection = "office -> site"
    ElseIf value = "site->office" Then
        NormalisedDirection = "site -> office"
    Else
        NormalisedDirection = value
    End If
End Function

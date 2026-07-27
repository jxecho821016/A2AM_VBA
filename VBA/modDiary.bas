Option Explicit

' Site Diary entry creation.
'
' Duplicates the "Diary_Template" sheet and names the copy yymmdd-Diary
' (e.g. 260727-Diary). Up to MAX_ENTRIES_PER_DAY entries are allowed per
' day; extras get a "(n)" suffix (260727-Diary(2) ... (5)). A sixth
' attempt is blocked with a reminder instead of creating a duplicate.
'
' The new copy is captured by position (last sheet), never via
' ActiveSheet - with ScreenUpdating off, ActiveSheet is unreliable and
' was the root cause of copies being lost, misnamed, or left hidden.

Private Const DIARY_TEMPLATE_SHEET As String = "Diary_Template"
Private Const MAX_ENTRIES_PER_DAY As Long = 5

Public Sub CreateDiaryEntry()
    Dim templateSheet As Worksheet
    Dim newSheet As Worksheet
    Dim baseName As String
    Dim newName As String
    Dim entryNumber As Long
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim errorMessage As String

    On Error GoTo CleanFail

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set templateSheet = ThisWorkbook.Worksheets(DIARY_TEMPLATE_SHEET)
    baseName = Format$(Date, "yymmdd") & "-Diary"

    entryNumber = 1
    newName = baseName
    Do While DiarySheetExists(newName)
        entryNumber = entryNumber + 1
        If entryNumber > MAX_ENTRIES_PER_DAY Then
            Application.EnableEvents = previousEnableEvents
            Application.ScreenUpdating = previousScreenUpdating
            MsgBox "There are already " & MAX_ENTRIES_PER_DAY & _
                   " diary entries for today (" & baseName & ")." & vbCrLf & _
                   "No additional entry was created.", vbExclamation, _
                   "Daily entry limit reached"
            Exit Sub
        End If
        newName = baseName & "(" & entryNumber & ")"
    Loop

    templateSheet.Copy _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)

    ' Capture the copy by position, not ActiveSheet. The copy always
    ' lands as the new last sheet regardless of where Diary_Template
    ' sits in the tab order or whether ScreenUpdating is off.
    Set newSheet = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    newSheet.Visible = xlSheetVisible
    newSheet.Name = newName
    newSheet.Activate

CleanExit:
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

CleanFail:
    errorMessage = Err.Description
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "The new diary entry could not be created." & vbCrLf & _
           errorMessage, vbExclamation
End Sub

Private Function DiarySheetExists(ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    DiarySheetExists = Not sheet Is Nothing
End Function

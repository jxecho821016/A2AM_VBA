Option Explicit

' Front page navigation, and a repair for buttons that point at the
' wrong workbook.
'
' Front_Page returns to the front sheet. Assign it to the "Front Page"
' button on Record_Template, Diary_Template and Template.
'
' RepairButtonMacros fixes a problem that appears whenever a project
' workbook is created by copying an earlier one: Excel rewrites each
' copied button's assigned macro to name the ORIGINAL file, so the
' button stores something like
'     'FOP-0058-Project_Mob_List.xlsm'!Front_Page
' instead of plain "Front_Page". Clicking it then opens that other
' workbook rather than running the macro here. The repair strips the
' workbook prefix from every button in this workbook, and reports what
' it changed. Run it once after copying a workbook for a new project.

Private Const FRONT_PAGE_SHEET As String = "Site Details"

Public Sub Front_Page()
    Dim frontSheet As Worksheet

    Set frontSheet = FindSheet(FRONT_PAGE_SHEET)
    If frontSheet Is Nothing Then
        MsgBox "This workbook has no """ & FRONT_PAGE_SHEET & """ sheet." & _
               vbCrLf & vbCrLf & "Sheets found: " & SheetNameList() & _
               vbCrLf & vbCrLf & _
               "Rename the front sheet, or change the FRONT_PAGE_SHEET " & _
               "constant at the top of modNavigation.", _
               vbExclamation, "Front page"
        Exit Sub
    End If

    frontSheet.Visible = xlSheetVisible
    frontSheet.Activate
    frontSheet.Range("A1").Select
End Sub

Public Sub RepairButtonMacros()
    Dim sheet As Worksheet
    Dim shape As Object
    Dim currentAction As String
    Dim bareName As String
    Dim report As String
    Dim fixedCount As Long

    On Error Resume Next

    For Each sheet In ThisWorkbook.Worksheets
        For Each shape In sheet.Shapes
            currentAction = vbNullString
            currentAction = shape.OnAction

            ' A macro name carrying "!" is qualified with a workbook -
            ' for a macro in this workbook that prefix is never needed,
            ' and when it names another file the button opens that file.
            If InStr(currentAction, "!") > 0 Then
                bareName = Mid$(currentAction, InStrRev(currentAction, "!") + 1)
                If Len(bareName) > 0 Then
                    shape.OnAction = bareName
                    report = report & vbCrLf & "  " & sheet.Name & _
                             " - " & currentAction & "  ->  " & bareName
                    fixedCount = fixedCount + 1
                End If
            End If
        Next shape
    Next sheet

    On Error GoTo 0

    If fixedCount = 0 Then
        MsgBox "Every button already points at a macro in this " & _
               "workbook - nothing needed changing.", _
               vbInformation, "Repair buttons"
    Else
        MsgBox fixedCount & " button(s) repointed at this workbook:" & _
               vbCrLf & report & vbCrLf & vbCrLf & _
               "Save the workbook to keep these changes.", _
               vbInformation, "Repair buttons"
    End If
End Sub

Private Function FindSheet(ByVal wantedName As String) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In ThisWorkbook.Worksheets
        If LCase$(Trim$(candidate.Name)) = LCase$(Trim$(wantedName)) Then
            Set FindSheet = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Function SheetNameList() As String
    Dim candidate As Worksheet
    Dim names As String

    For Each candidate In ThisWorkbook.Worksheets
        If Len(names) > 0 Then names = names & ", "
        names = names & candidate.Name
    Next candidate

    SheetNameList = names
End Function

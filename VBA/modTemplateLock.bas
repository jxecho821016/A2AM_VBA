Option Explicit

' Template protection.
'
' Protecting the workbook's STRUCTURE greys out right-click > Move or
' Copy, Insert, Delete and Rename, so nobody can duplicate
' Record_Template, Diary_Template or Template by hand. Cell editing and
' data entry are unaffected - structure protection only blocks
' sheet-level operations.
'
' The existing creation macros cannot copy sheets while that protection
' is on, so the buttons are pointed at the Run... wrappers below
' instead. Each one unlocks the structure, runs the original macro
' untouched, and locks it again - even if the macro fails.
'
' SETUP (once):
'   1. Paste this module into the workbook (Alt+F11 > Insert > Module).
'   2. Reassign the buttons (right-click each > Assign Macro):
'        Mobilisation / Demobilisation New Entry -> RunCreateNewEntry
'        Site Diary Template New Entry           -> RunCreateDiaryEntry
'        Generate Pile Record                    -> RunCreatePileRecords
'        Generate Form Entry (on Template)       -> RunGenerateFormEntry
'   3. Run ProtectTemplates once, then save the workbook.
'
' To edit a template yourself later, run UnlockForEditing, make the
' change, then run ProtectTemplates again. Note that any button click
' re-applies the protection.
'
' This stops accidental and casual duplication. It is not security: the
' password sits in this module, so also lock the VBA project itself
' (Tools > VBAProject Properties > Protection) to keep it out of sight.

Private Const STRUCTURE_PASSWORD As String = "A2AM-templates"

Public Sub ProtectTemplates()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=STRUCTURE_PASSWORD
    On Error GoTo 0

    ThisWorkbook.Protect Password:=STRUCTURE_PASSWORD, _
                         Structure:=True, Windows:=False

    MsgBox "The templates are protected." & vbCrLf & vbCrLf & _
           "Right-click > Move or Copy is now unavailable, so the " & _
           "templates can only be duplicated with the buttons." & _
           vbCrLf & vbCrLf & "Save the workbook to keep this.", _
           vbInformation, "Templates protected"
End Sub

Public Sub UnlockForEditing()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=STRUCTURE_PASSWORD
    On Error GoTo 0

    MsgBox "Sheet protection is off, so the templates can be edited " & _
           "and copied by hand." & vbCrLf & vbCrLf & _
           "Run ProtectTemplates when you are finished. Clicking any " & _
           "of the buttons also switches it back on.", _
           vbInformation, "Templates unlocked"
End Sub

' ---- button wrappers: unlock, run the original macro, lock again ----

Public Sub RunCreateNewEntry()
    AllowSheetChanges
    On Error Resume Next
    Application.Run "CreateNewEntry"
    On Error GoTo 0
    RestoreProtection
End Sub

Public Sub RunCreateDiaryEntry()
    AllowSheetChanges
    On Error Resume Next
    Application.Run "CreateDiaryEntry"
    On Error GoTo 0
    RestoreProtection
End Sub

Public Sub RunCreatePileRecords()
    AllowSheetChanges
    On Error Resume Next
    Application.Run "CreatePileRecords"
    On Error GoTo 0
    RestoreProtection
End Sub

Public Sub RunGenerateFormEntry()
    AllowSheetChanges
    On Error Resume Next
    Application.Run "GenerateFormEntry"
    On Error GoTo 0
    RestoreProtection
End Sub

Private Sub AllowSheetChanges()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=STRUCTURE_PASSWORD
    On Error GoTo 0
End Sub

' Always re-applies protection, so a failed run can never leave the
' templates copyable.
Private Sub RestoreProtection()
    On Error Resume Next
    ThisWorkbook.Protect Password:=STRUCTURE_PASSWORD, _
                         Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

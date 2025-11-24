# Edit Single Record Form – VBA Module

## Description

Template for an Access form that edits a single record with Save/Delete buttons and refreshes a parent form's subform when closed.

## Code

```
Option Compare Database
Option Explicit

' Template: Edit Single Record Form
'
' Assumptions:
'   - The form is bound to a single record (navigation disabled)
'   - Includes Save and Delete buttons (btnSave, btnDelete)
'   - Has a parent form with a subform that lists all records


' DELETE BUTTON CLICK
' Deletes the current record (after confirmation) and closes the form.
Private Sub btnDelete_Click()

    ' Custom confirmation dialog
    If MsgBox("Delete this return tracking?", vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Exit Sub
    End If

    ' Temporarily disable Access delete prompts
    DoCmd.SetWarnings False

    ' Delete the record
    DoCmd.RunCommand acCmdDeleteRecord

    ' Restore warnings
    DoCmd.SetWarnings True

    ' Close this form
    DoCmd.Close acForm, Me.Name

End Sub


' SAVE BUTTON CLICK
' Saves the current record and closes the form.
Private Sub btnSave_Click()

    On Error GoTo ErrHandler

    ' Save the record
    DoCmd.RunCommand acCmdSaveRecord

    ' Optional:
    ' MsgBox "Record saved!", vbInformation

    ' Close the form
    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrHandler:
    MsgBox "Error saving record: " & Err.Description, vbCritical

End Sub


' FORM CLOSE EVENT
' When this edit form closes, refresh the parent form's subform (if parent is open).
Private Sub Form_Close()

    Dim parentFormName As String
    Dim subFormName As String
    Dim frmParent As Form

    ' Parent and subform control names
    parentFormName = "ReturnTracking_MainForm"
    subFormName = "ReturnTrackingQuery_Subform"

    ' Only requery if parent form is open
    If CurrentProject.AllForms(parentFormName).IsLoaded Then

        ' Reference parent form
        Set frmParent = Forms(parentFormName)

        ' Refresh the listing subform
        frmParent(subFormName).Form.Requery

    End If

End Sub


' FORM LOAD EVENT
' Disable Delete button when on a new (unsaved) record.
Private Sub Form_Load()
    Me.btnDelete.Enabled = Not Me.NewRecord
End Sub
```

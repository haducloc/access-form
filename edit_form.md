# Edit Single Record Form – Custom Form - Form Code Module VBA

## Description

- This template is for an Access form that edits a **single record**, with **Save** and **Delete** buttons.
- The form can be opened **directly**, or it can be opened from the **parent form** (`ReturnTracking_MainForm`).
- The form supports both **editing an existing record** and **creating a new record**.
- The **Delete** button is enabled only when editing an existing record; it is disabled for new records.
- In the **Save** button’s Click event, you can perform validation and display error messages as needed.
- When the form closes, it will automatically refresh the **subform on the parent form** (if the parent form is open) so the latest data is displayed.

---

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
' When this form closes, refresh the parent form's subform (if the parent form is open).
Private Sub Form_Close()

    Dim parentFormName As String
    Dim subFormName As String
    Dim frmParent As Form

    ' Parent form and subform CONTROL names
    parentFormName = "ReturnTracking_MainForm"

    ' NOTE: This is the subform CONTROL name, not the form object name.
    subFormName = "ReturnTrackingQuery_SubForm"

    ' Only refresh if the parent form is currently open
    If FormLoaded(parentFormName) Then

        ' Get reference to the parent form
        Set frmParent = Forms(parentFormName)

        ' Refresh the listing subform
        frmParent(subFormName).Form.Requery

    End If

End Sub


' FORM LOAD EVENT
' Disable the Delete button when the form is adding a new (unsaved) record.
Private Sub Form_Load()
    Me.btnDelete.Enabled = Not Me.NewRecord
End Sub
```

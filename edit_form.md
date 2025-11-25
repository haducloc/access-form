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

```vba
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
    HandleDeleteClick Me
End Sub


' SAVE BUTTON CLICK
' Saves the current record and closes the form.
Private Sub btnSave_Click()
    HandleSaveClick Me
End Sub

' FORM BEFORE UPDATE
' Disable saving automatically for any input value changes
Private Sub Form_BeforeUpdate(Cancel As Integer)
    HandleFormBeforeUpdate Me, Cancel
End Sub

' FORM CLOSE EVENT
' When this form closes, requery the parent form's subform
' — but only if the parent form is currently open.
Private Sub Form_Close()
    Dim parentFormName As String: parentFormName = "ReturnTracking_MainForm"
    Dim subformControlName As String: subformControlName = "ReturnTrackingQuery_SubForm"

    HandleFormCloseRefresh parentFormName, subformControlName
End Sub

' FORM LOAD EVENT
' Disable the Delete button when the form is adding a new (unsaved) record.
Private Sub Form_Load()
    Me.btnDelete.Enabled = Not Me.NewRecord
End Sub
```

Option Compare Database
Option Explicit

' ============================================================================
'  Template: Edit Single Record Form
'
'  Assumptions:
'    • Form is bound to a single record (navigation disabled)
'    • Contains Save and Delete buttons (btnSave, btnDelete)
'    • Has a parent form with a subform listing all records
' ============================================================================


' ----------------------------------------------------------------------------
' DELETE BUTTON CLICK
' Deletes the current record (after custom confirmation) and closes the form.
' ----------------------------------------------------------------------------
Private Sub btnDelete_Click()

    ' Custom confirmation dialog
    If MsgBox("Delete this return tracking?", vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Exit Sub
    End If

    ' Temporarily disable Access system warnings (built-in delete prompt)
    DoCmd.SetWarnings False

    ' Delete the record
    DoCmd.RunCommand acCmdDeleteRecord

    ' RESTORE warnings immediately
    DoCmd.SetWarnings True

    ' Close this form
    DoCmd.Close acForm, Me.Name

End Sub


' ----------------------------------------------------------------------------
' SAVE BUTTON CLICK
' Saves the current record and closes the form.
' ----------------------------------------------------------------------------
Private Sub btnSave_Click()

    On Error GoTo ErrHandler

    ' Save current record
    DoCmd.RunCommand acCmdSaveRecord

    ' Optional:
    ' MsgBox "Record saved!", vbInformation

    ' Close this form
    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrHandler:
    MsgBox "Error saving record: " & Err.Description, vbCritical

End Sub


' ----------------------------------------------------------------------------
' FORM CLOSE EVENT
' When this edit form closes, requery the parent form's subform so the list
' updates instantly with any changes (save/delete).
' ----------------------------------------------------------------------------
Private Sub Form_Close()

    Dim parentFormName As String
    Dim subFormName As String
    Dim frmParent As Form

    ' Parent form and subform control names
    parentFormName = "ReturnTracking_MainForm"
    subFormName = "ReturnTrackingQuery_Subform"

    ' Only requery if parent is open
    If CurrentProject.AllForms(parentFormName).IsLoaded Then

        ' Get reference to parent form
        Set frmParent = Forms(parentFormName)

        ' Requery subform
        frmParent(subFormName).Form.Requery

    End If

End Sub


' ----------------------------------------------------------------------------
' FORM LOAD EVENT
' Disable Delete button when on a new, unsaved record.
' ----------------------------------------------------------------------------
Private Sub Form_Load()
    Me.btnDelete.Enabled = Not Me.NewRecord
End Sub

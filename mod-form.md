```vba
Option Compare Database
Option Explicit


'============================================================
' FormExists
' Returns True if a form exists in the project.
'============================================================
Public Function FormExists(FormName As String) As Boolean
    On Error Resume Next
    Dim ao As AccessObject: Set ao = CurrentProject.AllForms(FormName)
    FormExists = (Err.Number = 0)
    Err.Clear: On Error GoTo 0
End Function


'============================================================
' FormLoaded
' Returns True if a form exists AND is loaded/open.
'============================================================
Public Function FormLoaded(FormName As String) As Boolean
    On Error Resume Next
    Dim ao As AccessObject: Set ao = CurrentProject.AllForms(FormName)
    If Err.Number <> 0 Then
        FormLoaded = False
    Else
        FormLoaded = ao.IsLoaded
    End If
    Err.Clear: On Error GoTo 0
End Function


'============================================================
' HandleSaveClick
' Saves record using Tag="InSaveClickContext" to allow saving,
' then closes the form. Must be called from btnSave_Click.
'============================================================
Public Sub HandleSaveClick(frm As Form)
    On Error GoTo ErrHandler
    
    frm.Tag = "InSaveClickContext"
    DoCmd.RunCommand acCmdSaveRecord
    frm.Tag = ""
    
    DoCmd.Close acForm, frm.Name
    Exit Sub

ErrHandler:
    MsgBox "Unexpected error: " & Err.Description, vbCritical
    frm.Tag = ""
End Sub


'============================================================
' HandleDeleteClick
' Confirms deletion, executes delete, and closes the form.
'============================================================
Public Sub HandleDeleteClick(frm As Form, confirmMsg As String)
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
    DoCmd.Close acForm, frm.Name
End Sub


'============================================================
' HandleFormBeforeUpdate
' Blocks all Access auto-save attempts except during save click.
' Call from Form_BeforeUpdate.
'============================================================
Public Sub HandleFormBeforeUpdate(frm As Form, Cancel As Integer)
    If frm.Tag <> "InSaveClickContext" Then Cancel = True
End Sub

'-------------------------------------------------------------
' RefreshParentSubform
'   Refreshes/Requeries a subform control on a parent form,
'   but ONLY if the parent form exists and is currently loaded.
'
'   parentFormName      = name of the parent form (string)
'   subformControlName  = name of the SUBFORM CONTROL
'-------------------------------------------------------------
Public Sub RefreshParentSubform(parentFormName As String, subformControlName As String)
    On Error GoTo Cleanup

    ' Ensure parent form exists and is open
    If Not FormExists(parentFormName) Then GoTo Cleanup
    If Not FormLoaded(parentFormName) Then GoTo Cleanup

    ' Refresh the specified subform control
    Forms(parentFormName)(subformControlName).Form.Requery
    Forms(parentFormName)(subformControlName).Form.Refresh

Cleanup:
    Exit Sub
End Sub
```

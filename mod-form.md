```vba
Option Compare Database
Option Explicit


'============================================================
' FormExists
' Returns True if a form exists in the project.
'============================================================
Public Function FormExists(FormName As String) As Boolean
    ' Access throws an error here if the form does not exist.
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

'============================================================
' HandleFormCloseRefresh
'
' Called from the Form_Close event of an edit or child form.
' If the specified parent form is open, this procedure requeries
' the specified subform control within that parent form.
'
' The routine is intentionally fail-safe: any errors (such as
' missing controls, wrong names, or forms not being loaded)
' are ignored to avoid interrupting the close operation.
'============================================================
Public Sub HandleFormCloseRefresh(parentFormName As String, subformControlName As String)

    Dim frmParent As Form
    On Error GoTo ExitHandler   ' Fail softly — never block form close.

    ' Proceed only if the parent form is open and loaded.
    If Not FormLoaded(parentFormName) Then Exit Sub

    ' Get a reference to the parent form instance.
    Set frmParent = Forms(parentFormName)

    ' Ensure the named control exists AND is a subform control.
    If frmParent.Controls(subformControlName).ControlType = acSubform Then
        frmParent(subformControlName).Form.Requery
    End If

ExitHandler:
    ' Silent exit by design; parent form refresh is non-critical.
    Exit Sub
End Sub
```

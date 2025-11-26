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
Public Function FormLoaded(FormName As String) As Form
    On Error Resume Next
    Set FormLoaded = Forms(FormName)
    On Error GoTo 0
End Function


'============================================================
' ControlExists
' Returns True if a control exists on an OPEN form.
'============================================================
Public Function ControlExists(FormName As String, ControlName As String) As Boolean
    On Error Resume Next
    Dim ctl As Control
    Set ctl = Forms(FormName).Controls(ControlName)
    ControlExists = (Err.Number = 0)
    Err.Clear: On Error GoTo 0
End Function


'============================================================
' GetControl
' Returns a control if exists on an OPEN form.
'============================================================
Public Function GetControl(frm As Form, ControlName As String) As Control
    On Error Resume Next
    Set GetControl = frm.Controls(ControlName)
    On Error GoTo 0
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
    On Error GoTo ErrHandler
    
    Dim parentForm As Form: Set parentForm = FormLoaded(parentFormName)
    If parentForm Is Nothing Then GoTo ErrHandler
    
    Dim subFormCtrl As Control
    Set subFormCtrl = GetControl(parentForm, subformControlName)
    If subFormCtrl Is Nothing Then GoTo ErrHandler
    
    subFormCtrl.Form.Requery
    subFormCtrl.Form.Refresh

ErrHandler:
    Exit Sub
End Sub
```

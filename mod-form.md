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
' HandleAccessError
' Replaces Access validation popups with custom message and Undo.
' Call from Form_Error.
'============================================================
Public Sub HandleAccessError(frm As Form, DataErr As Integer)
    Dim badValue As String
    
    On Error Resume Next
    badValue = frm.ActiveControl.Text
    If Err.Number <> 0 Then badValue = ""
    Err.Clear: On Error GoTo 0
    
    If badValue <> "" Then
        MsgBox "Invalid value: " & badValue, vbCritical, "Validation Error"
    Else
        MsgBox "Invalid value.", vbCritical, "Validation Error"
    End If
    
    frm.Undo
End Sub


'============================================================
' NavFirst
' Moves to the first record (safe navigation).
'============================================================
Public Sub NavFirst(frm As Form)
    On Error Resume Next
    DoCmd.GoToRecord , frm.Name, acFirst
    Err.Clear
End Sub


'============================================================
' NavPrevious
' Moves to previous record; if at first, stays at first.
'============================================================
Public Sub NavPrevious(frm As Form)
    On Error Resume Next
    DoCmd.GoToRecord , frm.Name, acPrevious
    If Err.Number <> 0 Then DoCmd.GoToRecord , frm.Name, acFirst
    Err.Clear
End Sub


'============================================================
' NavNext
' Moves to next record; if at last, stays at last.
'============================================================
Public Sub NavNext(frm As Form)
    On Error Resume Next
    DoCmd.GoToRecord , frm.Name, acNext
    If Err.Number <> 0 Then DoCmd.GoToRecord , frm.Name, acLast
    Err.Clear
End Sub


'============================================================
' NavLast
' Moves to the last record (safe navigation).
'============================================================
Public Sub NavLast(frm As Form)
    On Error Resume Next
    DoCmd.GoToRecord , frm.Name, acLast
    Err.Clear
End Sub


'============================================================
' NavNew
' Moves to new blank record (safe).
'============================================================
Public Sub NavNew(frm As Form)
    On Error Resume Next
    DoCmd.GoToRecord , frm.Name, acNewRec
    Err.Clear
End Sub
```

Option Compare Database
Option Explicit


'============================================================
'  FormExists
'  Returns True if a form exists in the project.
'============================================================
Public Function FormExists(FormName As String) As Boolean
    On Error Resume Next

    Dim ao As AccessObject
    Set ao = CurrentProject.AllForms(FormName)

    FormExists = (Err.Number = 0)

    Err.Clear
    On Error GoTo 0
End Function


'============================================================
'  FormLoaded
'  Returns True if the form exists AND is currently open.
'============================================================
Public Function FormLoaded(FormName As String) As Boolean
    On Error Resume Next

    Dim ao As AccessObject
    Set ao = CurrentProject.AllForms(FormName)

    If Err.Number <> 0 Then
        FormLoaded = False
    Else
        FormLoaded = ao.IsLoaded
    End If

    Err.Clear
    On Error GoTo 0
End Function


'============================================================
'  HandleSaveClick
'  Performs Save button behavior:
'     - Marks save context
'     - Saves the record
'     - Clears the save flag
'     - Closes the form
'
'  MUST BE CALLED FROM A FORM:
'     Call HandleSaveClick(Me)
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
'  HandleDeleteClick
'  Prompts user, deletes record, restores warnings, closes form.
'
'  Usage:
'     Call HandleDeleteClick(Me, "Delete this record?")
'============================================================
Public Sub HandleDeleteClick(frm As Form, confirmMsg As String)

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Exit Sub
    End If

    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True

    DoCmd.Close acForm, frm.Name
End Sub


'============================================================
'  HandleFormBeforeUpdate
'  Blocks ALL Access auto-saves except during HandleSaveClick.
'
'  MUST BE CALLED FROM A FORM'S BeforeUpdate:
'
'     Private Sub Form_BeforeUpdate(Cancel As Integer)
'         HandleFormBeforeUpdate Me, Cancel
'     End Sub
'============================================================
Public Sub HandleFormBeforeUpdate(frm As Form, Cancel As Integer)
    If frm.Tag <> "InSaveClickContext" Then
        Cancel = True   ' block Access auto-save
    End If
End Sub

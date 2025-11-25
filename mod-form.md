# Access VBA Utility Functions  
- Define common utility functions

---

```vba
Option Compare Database
Option Explicit

' Returns True if a form with the given name exists in the current project.
Function FormExists(FormName As String) As Boolean
    ' Suppress errors (e.g., if the form name does not exist)
    On Error Resume Next

    Dim ao As AccessObject

    ' Attempt to reference the form object in the project
    Set ao = CurrentProject.AllForms(FormName)

    ' If no error occurred, the form exists
    FormExists = (Err.Number = 0)

    ' Clear any error and restore normal error handling
    Err.Clear
    On Error GoTo 0
End Function

' Returns True if the form exists and is currently loaded (open) in memory.
Function FormLoaded(FormName As String) As Boolean
    ' Suppress errors for invalid form names
    On Error Resume Next

    Dim ao As AccessObject

    ' Attempt to reference the form definition
    Set ao = CurrentProject.AllForms(FormName)

    ' If form does not exist: return False
    If Err.Number <> 0 Then
        FormLoaded = False
    Else
        ' Form exists — check if it is currently open (loaded)
        FormLoaded = ao.IsLoaded
    End If

    ' Clear any error and restore normal behavior
    Err.Clear
    On Error GoTo 0
End Function

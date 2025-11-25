```vba
' =============================
'   CONTROL STATE STRUCT
' =============================
Public Type ControlState
    ConvertedValue As Variant
    IsValid As Boolean
End Type

Public Function MakeState(val As Variant, IsValid As Boolean) As ControlState
    Dim cs As ControlState
    cs.ConvertedValue = val
    cs.IsValid = IsValid
    MakeState = cs
End Function

' =============================
'   GETTERS ("" IS ALWAYS VALID)
' =============================

' ---- STRING ----
Public Function GetString(ctrl As Control) As ControlState
    Dim s As String
    s = Trim(Nz(ctrl.Value, ""))
    GetString = MakeState(s, True)
End Function

' ---- INTEGER ----
Public Function GetInt(ctrl As Control) As ControlState
    Dim v As Variant: v = Nz(ctrl.Value, "")

    If v = "" Then
        GetInt = MakeState(Null, True)
    ElseIf IsNumeric(v) And CDbl(v) = Fix(CDbl(v)) Then
        GetInt = MakeState(CLng(v), True)
    Else
        GetInt = MakeState(Null, False)
    End If
End Function


' ---- DOUBLE ----
Public Function GetDouble(ctrl As Control) As ControlState
    Dim v As Variant: v = Nz(ctrl.Value, "")

    If v = "" Then
        GetDouble = MakeState(Null, True)
    ElseIf IsNumeric(v) Then
        GetDouble = MakeState(CDbl(v), True)
    Else
        GetDouble = MakeState(Null, False)
    End If
End Function

' ---- DATE ----
Public Function GetDate(ctrl As Control) As ControlState
    Dim v As Variant: v = Nz(ctrl.Value, "")

    If v = "" Then
        GetDate = MakeState(Null, True)
    ElseIf IsDate(v) Then
        GetDate = MakeState(DateValue(v), True)
    Else
        GetDate = MakeState(Null, False)
    End If
End Function

' ---- TIME ----
Public Function GetTime(ctrl As Control) As ControlState
    Dim v As Variant: v = Nz(ctrl.Value, "")

    If v = "" Then
        GetTime = MakeState(Null, True)
    ElseIf IsDate(v) Then
        GetTime = MakeState(TimeValue(v), True)
    Else
        GetTime = MakeState(Null, False)
    End If
End Function

' ---- DATETIME ----
Public Function GetDateTime(ctrl As Control) As ControlState
    Dim v As Variant: v = Nz(ctrl.Value, "")

    If v = "" Then
        GetDateTime = MakeState(Null, True)
    ElseIf IsDate(v) Then
        GetDateTime = MakeState(CDate(v), True)
    Else
        GetDateTime = MakeState(Null, False)
    End If
End Function

' ---- BOOLEAN ----
Public Function GetBool(ctrl As Control) As ControlState
    Dim v As String
    v = LCase(Trim(Nz(ctrl.Value, "")))

    If v = "" Then
        GetBool = MakeState(Null, True)
        Exit Function
    End If

    Select Case v
        Case "true", "yes", "y", "1"
            GetBool = MakeState(True, True)
        Case "false", "no", "n", "0"
            GetBool = MakeState(False, True)
        Else
            GetBool = MakeState(Null, False)
    End Select
End Function
```

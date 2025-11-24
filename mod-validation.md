# Validation Module – `modValidation`
A compact, control-based validation framework for Microsoft Access forms.  
Supports **text**, **integer**, **double**, and **date** validation with:

- Automatic trimming  
- Required/optional logic  
- Writing cleaned values back into the control  
- Returning a `ValidationResult` structure  


---

## `ValidationResult` Structure

```vb
Public Type ValidationResult
    IsValid As Boolean     ' TRUE if validation passed
    Value As Variant       ' Cleaned/converted value (Long, Double, String, Date, or Null)
    Message As String      ' Optional failure message
End Type
```


---

## Internal Helper

```vb
' CV() – CleanValue helper
' - Converts Null → ""
' - Trims leading/trailing spaces
Private Function CV(val As Variant) As String
    CV = Trim(Nz(val, ""))
End Function
```


---

# STRING VALIDATION
Validation for text-based controls.

---

## `ValidateRequiredText`
Ensures text input is **not blank**, trims it, and writes it back.

```vb
Public Function ValidateRequiredText(ctrl As Control) As ValidationResult
    Dim s As String: s = CV(ctrl.Value)

    If s = "" Then
        ValidateRequiredText.Message = ctrl.Name & " is required."
        Exit Function
    End If

    ctrl.Value = s
    ValidateRequiredText.IsValid = True
    ValidateRequiredText.Value = s
End Function
```


---

## `ValidateMaxLength`
Ensures text length ≤ maxLen.

```vb
Public Function ValidateMaxLength(ctrl As Control, maxLen As Long, Optional Required As Boolean = False) As ValidationResult
    Dim s As String: s = CV(ctrl.Value)

    If s = "" Then
        If Required Then ValidateMaxLength.Message = ctrl.Name & " is required." _
        Else ValidateMaxLength.IsValid = True
        Exit Function
    End If

    If Len(s) > maxLen Then
        ValidateMaxLength.Message = ctrl.Name & " cannot exceed " & maxLen & " characters."
        Exit Function
    End If

    ctrl.Value = s
    ValidateMaxLength.IsValid = True
    ValidateMaxLength.Value = s
End Function
```


---

## `ValidateMinLength`
Ensures text length ≥ minLen.

```vb
Public Function ValidateMinLength(ctrl As Control, minLen As Long, Optional Required As Boolean = False) As ValidationResult
    Dim s As String: s = CV(ctrl.Value)

    If s = "" Then
        If Required Then ValidateMinLength.Message = ctrl.Name & " is required." _
        Else ValidateMinLength.IsValid = True
        Exit Function
    End If

    If Len(s) < minLen Then
        ValidateMinLength.Message = ctrl.Name & " must be at least " & minLen & " characters."
        Exit Function
    End If

    ctrl.Value = s
    ValidateMinLength.IsValid = True
    ValidateMinLength.Value = s
End Function
```


---

# INTEGER VALIDATION
Validates integers (whole numbers).

---

### INTERNAL HELPER  
Common logic for integer-based validators.

```vb
Private Function IntResult(ctrl As Control, Required As Boolean) As ValidationResult
    Dim s As String: s = CV(ctrl.Value)

    If s = "" Then
        If Required Then Exit Function
        IntResult.IsValid = True
        IntResult.Value = Null
        Exit Function
    End If

    If Not IsNumeric(s) Then Exit Function
    If CLng(s) <> Val(s) Then Exit Function   ' must be integer only

    ctrl.Value = CLng(s)
    IntResult.IsValid = True
    IntResult.Value = CLng(s)
End Function
```


---

## `ValidateInteger`
Accepts any integer.

```vb
Public Function ValidateInteger(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    ValidateInteger = IntResult(ctrl, Required)
End Function
```


---

## `ValidatePositiveInteger`
Accepts integers > 0.

```vb
Public Function ValidatePositiveInteger(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    Dim r As ValidationResult: r = IntResult(ctrl, Required)
    If Not r.IsValid Then Exit Function
    If Not IsNull(r.Value) And r.Value <= 0 Then Exit Function

    r.IsValid = True
    ValidatePositiveInteger = r
End Function
```


---

## `ValidateNotNegativeInteger`
Accepts integers ≥ 0.

```vb
Public Function ValidateNotNegativeInteger(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    Dim r As ValidationResult: r = IntResult(ctrl, Required)
    If Not r.IsValid Then Exit Function
    If Not IsNull(r.Value) And r.Value < 0 Then Exit Function

    r.IsValid = True
    ValidateNotNegativeInteger = r
End Function
```


---

# DOUBLE VALIDATION
Validates decimal numbers.

---

### INTERNAL HELPER  
Common logic for all double validators.

```vb
Private Function DblResult(ctrl As Control, Required As Boolean) As ValidationResult
    Dim s As String: s = CV(ctrl.Value)

    If s = "" Then
        If Required Then Exit Function
        DblResult.IsValid = True
        DblResult.Value = Null
        Exit Function
    End If

    If Not IsNumeric(s) Then Exit Function

    ctrl.Value = CDbl(s)
    DblResult.IsValid = True
    DblResult.Value = CDbl(s)
End Function
```


---

## `ValidateDouble`
Accepts any decimal value.

```vb
Public Function ValidateDouble(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    ValidateDouble = DblResult(ctrl, Required)
End Function
```


---

## `ValidatePositiveDouble`
Accepts decimals > 0.

```vb
Public Function ValidatePositiveDouble(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    Dim r As ValidationResult: r = DblResult(ctrl, Required)
    If Not r.IsValid Then Exit Function
    If Not IsNull(r.Value) And r.Value <= 0 Then Exit Function

    r.IsValid = True
    ValidatePositiveDouble = r
End Function
```


---

## `ValidateNotNegativeDouble`
Accepts decimals ≥ 0.

```vb
Public Function ValidateNotNegativeDouble(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    Dim r As ValidationResult: r = DblResult(ctrl, Required)
    If Not r.IsValid Then Exit Function
    If Not IsNull(r.Value) And r.Value < 0 Then Exit Function

    r.IsValid = True
    ValidateNotNegativeDouble = r
End Function
```


---

# DATE VALIDATION
Handles date picker controls and Date/Time fields.

---

## `ValidateDate`
Ensures a valid date when required, or accepts Null if optional.

```vb
Public Function ValidateDate(ctrl As Control, Optional Required As Boolean = True) As ValidationResult

    If IsNull(ctrl.Value) Or CV(ctrl.Value) = "" Then
        If Required Then Exit Function
        ValidateDate.IsValid = True
        ValidateDate.Value = Null
        Exit Function
    End If

    If Not IsDate(ctrl.Value) Then Exit Function

    ctrl.Value = CDate(ctrl.Value)
    ValidateDate.IsValid = True
    ValidateDate.Value = ctrl.Value
End Function
```


---

# Usage Examples

## Required trimmed text
```vb
If Not ValidateRequiredText(Me!CtName).IsValid Then
    MsgBox "CtName is required", vbCritical
    Exit Sub
End If
```

## Required non-negative integer
```vb
If Not ValidateNotNegativeInteger(Me!ApplicantID).IsValid Then
    MsgBox "ApplicantID must be >= 0", vbCritical
    Exit Sub
End If
```

## Optional positive double
```vb
If Not ValidatePositiveDouble(Me!Amount, False).IsValid Then
    MsgBox "Amount must be > 0", vbCritical
    Exit Sub
End If
```

## Required date
```vb
If Not ValidateDate(Me!BirthDate).IsValid Then
    MsgBox "Birthdate is invalid", vbCritical
    Exit Sub
End If

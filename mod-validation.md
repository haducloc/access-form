# Validation Module – `modValidation`
A compact, control-based validation framework for Microsoft Access forms.  
Supports **text**, **integer**, **double**, and **date** validation with:

- Automatic trimming  
- Required/optional validation logic  
- Writes cleaned values back into the control  
- Returns a standardized `ValidationResult` structure  
- Implemented as a reusable library module that can be used in any Form code module or imported into other Access databases  

---

```vb
'===============================================
' Validation Framework (Final Simplified Version)
' With full comments and documentation
'===============================================

Option Compare Database
Option Explicit

'-----------------------------------------------
' ValidationResult Type
'
' Represents the outcome of a validation check.
' - IsValid: True if validation passed
' - Value: The cleaned/converted value (or Null)
' - ErrorMessage: Message describing why validation failed
'-----------------------------------------------
Public Type ValidationResult
    IsValid As Boolean
    Value As Variant
    ErrorMessage As String
End Type


'-----------------------------------------------
' cleanValue()
'
' Returns a trimmed string version of any input.
' Converts:
'   Null → ""
'   Numeric values → their string representation
'   Trimming removes leading/trailing spaces
'-----------------------------------------------
Private Function cleanValue(val As Variant) As String
    cleanValue = Trim(Nz(val, ""))
End Function


'-----------------------------------------------
' Fail()
'
' Creates a standardized failed ValidationResult.
' Ensures:
'   - IsValid = False
'   - Value = Null
'   - ErrorMessage = provided message
'
' This keeps failure handling consistent across
' all validators.
'-----------------------------------------------
Private Function Fail(ctrl As Control, msg As String) As ValidationResult
    Dim result As ValidationResult
    result.IsValid = False
    result.Value = Null
    result.ErrorMessage = msg
    Fail = result
End Function


'================================================
' STRING VALIDATION
'================================================

'-----------------------------------------------
' ValidateRequiredText()
'
' Ensures text is present (after trimming).
' Writes the cleaned value back into ctrl.Value.
'-----------------------------------------------
Public Function ValidateRequiredText(ctrl As Control) As ValidationResult
    Dim result As ValidationResult
    Dim stringValue As String: stringValue = cleanValue(ctrl.Value)

    If stringValue = "" Then
        ValidateRequiredText = Fail(ctrl, ctrl.Name & " is required.")
        Exit Function
    End If

    ctrl.Value = stringValue           ' normalize value
    result.IsValid = True
    result.Value = stringValue
    ValidateRequiredText = result
End Function


'-----------------------------------------------
' ValidateMaxLength()
'
' Ensures the trimmed text does not exceed maxLen.
' If Required=False, empty values are allowed.
'-----------------------------------------------
Public Function ValidateMaxLength(ctrl As Control, maxLen As Long, Optional Required As Boolean = False) As ValidationResult
    Dim result As ValidationResult
    Dim stringValue As String: stringValue = cleanValue(ctrl.Value)

    ' Handle empty value
    If stringValue = "" Then
        If Required Then
            ValidateMaxLength = Fail(ctrl, ctrl.Name & " is required.")
        Else
            result.IsValid = True
            ValidateMaxLength = result
        End If
        Exit Function
    End If

    ' Length enforcement
    If Len(stringValue) > maxLen Then
        ValidateMaxLength = Fail(ctrl, ctrl.Name & " cannot exceed " & maxLen & " characters.")
        Exit Function
    End If

    ctrl.Value = stringValue
    result.IsValid = True
    result.Value = stringValue
    ValidateMaxLength = result
End Function


'-----------------------------------------------
' ValidateMinLength()
'
' Ensures text meets minimum length requirement.
' If Required=False, empty values are allowed.
'-----------------------------------------------
Public Function ValidateMinLength(ctrl As Control, minLen As Long, Optional Required As Boolean = False) As ValidationResult
    Dim result As ValidationResult
    Dim stringValue As String: stringValue = cleanValue(ctrl.Value)

    ' Allow empty when not required
    If stringValue = "" Then
        If Required Then
            ValidateMinLength = Fail(ctrl, ctrl.Name & " is required.")
        Else
            result.IsValid = True
            ValidateMinLength = result
        End If
        Exit Function
    End If

    ' Enforce minimum length
    If Len(stringValue) < minLen Then
        ValidateMinLength = Fail(ctrl, ctrl.Name & " must be at least " & minLen & " characters.")
        Exit Function
    End If

    ctrl.Value = stringValue
    result.IsValid = True
    result.Value = stringValue
    ValidateMinLength = result
End Function


'================================================
' INTEGER VALIDATION
'================================================

'-----------------------------------------------
' IntResult()
'
' Core integer validation logic shared by:
'   - ValidateInteger
'   - ValidatePositiveInteger
'   - ValidateNotNegativeInteger
'
' Enforces:
'   - Optional/required logic
'   - Numeric format
'   - Strict integer check (no decimals, no "1e3")
'
' Writes cleaned integer back into ctrl.Value
'-----------------------------------------------
Private Function IntResult(ctrl As Control, Required As Boolean) As ValidationResult
    Dim result As ValidationResult
    Dim stringValue As String: stringValue = cleanValue(ctrl.Value)

    ' Handle empty value
    If stringValue = "" Then
        If Required Then
            IntResult = Fail(ctrl, ctrl.Name & " is required.")
        Else
            ctrl.Value = Null                      ' normalize optional
            result.IsValid = True
            result.Value = Null
            IntResult = result
        End If
        Exit Function
    End If

    ' Must be numeric
    If Not IsNumeric(stringValue) Then
        IntResult = Fail(ctrl, ctrl.Name & " must be a number.")
        Exit Function
    End If

    ' Strict integer: ensure the string representation matches an integer
    If CStr(CLng(stringValue)) <> Trim$(stringValue) Then
        IntResult = Fail(ctrl, ctrl.Name & " must be a whole number.")
        Exit Function
    End If

    ' Valid integer
    ctrl.Value = CLng(stringValue)
    result.IsValid = True
    result.Value = CLng(stringValue)
    IntResult = result
End Function


'-----------------------------------------------
' ValidateInteger()
'
' Accepts any whole number (required by default).
'-----------------------------------------------
Public Function ValidateInteger(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    ValidateInteger = IntResult(ctrl, Required)
End Function


'-----------------------------------------------
' ValidatePositiveInteger()
'
' Requires integer > 0 when present.
' Optional empty values allowed when Required=False.
'-----------------------------------------------
Public Function ValidatePositiveInteger(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    Dim result As ValidationResult
    result = IntResult(ctrl, Required)

    If Not result.IsValid Then
        ValidatePositiveInteger = result
        Exit Function
    End If

    If Not IsNull(result.Value) And result.Value <= 0 Then
        ValidatePositiveInteger = Fail(ctrl, ctrl.Name & " must be greater than 0.")
        Exit Function
    End If

    ValidatePositiveInteger = result
End Function


'-----------------------------------------------
' ValidateNotNegativeInteger()
'
' Requires integer >= 0 when present.
'-----------------------------------------------
Public Function ValidateNotNegativeInteger(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    Dim result As ValidationResult
    result = IntResult(ctrl, Required)

    If Not result.IsValid Then
        ValidateNotNegativeInteger = result
        Exit Function
    End If

    If Not IsNull(result.Value) And result.Value < 0 Then
        ValidateNotNegativeInteger = Fail(ctrl, ctrl.Name & " must be 0 or greater.")
        Exit Function
    End If

    ValidateNotNegativeInteger = result
End Function


'================================================
' DOUBLE (DECIMAL) VALIDATION
'================================================

'-----------------------------------------------
' DblResult()
'
' Shared logic for validating decimal numbers.
' Ensures optional/required behavior and numeric format.
'-----------------------------------------------
Private Function DblResult(ctrl As Control, Required As Boolean) As ValidationResult
    Dim result As ValidationResult
    Dim stringValue As String: stringValue = cleanValue(ctrl.Value)

    ' Handle empty
    If stringValue = "" Then
        If Required Then
            DblResult = Fail(ctrl, ctrl.Name & " is required.")
        Else
            ctrl.Value = Null
            result.IsValid = True
            result.Value = Null
            DblResult = result
        End If
        Exit Function
    End If

    ' Must be numeric
    If Not IsNumeric(stringValue) Then
        DblResult = Fail(ctrl, ctrl.Name & " must be a number.")
        Exit Function
    End If

    ' Valid double
    ctrl.Value = CDbl(stringValue)
    result.IsValid = True
    result.Value = CDbl(stringValue)
    DblResult = result
End Function


'-----------------------------------------------
' ValidateDouble()
'
' Accepts any numeric decimal value.
'-----------------------------------------------
Public Function ValidateDouble(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    ValidateDouble = DblResult(ctrl, Required)
End Function


'-----------------------------------------------
' ValidatePositiveDouble()
'
' Requires value > 0 when provided.
'-----------------------------------------------
Public Function ValidatePositiveDouble(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    Dim result As ValidationResult
    result = DblResult(ctrl, Required)

    If Not result.IsValid Then
        ValidatePositiveDouble = result
        Exit Function
    End If

    If Not IsNull(result.Value) And result.Value <= 0 Then
        ValidatePositiveDouble = Fail(ctrl, ctrl.Name & " must be greater than 0.")
        Exit Function
    End If

    ValidatePositiveDouble = result
End Function


'-----------------------------------------------
' ValidateNotNegativeDouble()
'
' Requires value >= 0 when provided.
'-----------------------------------------------
Public Function ValidateNotNegativeDouble(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    Dim result As ValidationResult
    result = DblResult(ctrl, Required)

    If Not result.IsValid Then
        ValidateNotNegativeDouble = result
        Exit Function
    End If

    If Not IsNull(result.Value) And result.Value < 0 Then
        ValidateNotNegativeDouble = Fail(ctrl, ctrl.Name & " must be 0 or greater.")
        Exit Function
    End If

    ValidateNotNegativeDouble = result
End Function


'================================================
' DATE VALIDATION
'================================================

'-----------------------------------------------
' ValidateDate()
'
' Ensures control contains a valid date.
' Optional empty behavior when Required=False.
'-----------------------------------------------
Public Function ValidateDate(ctrl As Control, Optional Required As Boolean = True) As ValidationResult
    Dim result As ValidationResult
    Dim stringValue As String: stringValue = cleanValue(ctrl.Value)

    ' Null or empty
    If IsNull(ctrl.Value) Or stringValue = "" Then
        If Required Then
            ValidateDate = Fail(ctrl, ctrl.Name & " is required.")
        Else
            ctrl.Value = Null
            result.IsValid = True
            result.Value = Null
            ValidateDate = result
        End If
        Exit Function
    End If

    ' Must be a valid date
    If Not IsDate(ctrl.Value) Then
        ValidateDate = Fail(ctrl, ctrl.Name & " must be a valid date.")
        Exit Function
    End If

    ' Normalize to VBA Date
    ctrl.Value = CDate(ctrl.Value)
    result.IsValid = True
    result.Value = ctrl.Value
    ValidateDate = result
End Function
```

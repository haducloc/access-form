```vba
Option Compare Database
Option Explicit

' ===========================================
' EscapeAccessWildcards
' Escapes Access wildcard characters (* ? # [ ])
' so they act as literal characters inside LIKE.
'
' Use this ONLY in "plain value" mode.
' Wildcards must NOT be escaped in pattern mode (prefix "~").
'
' Escape rules:
'   * -> [*]
'   ? -> [?]
'   # -> [#]
'   [ -> [[]
'   ] -> []]
'
' Parameters:
'   s - raw literal text
'
' Returns:
'   Escaped version of s suitable for LIKE literal search
' ===========================================
Private Function EscapeAccessWildcards(s As String) As String
    s = Replace(s, "[", "[[]")
    s = Replace(s, "]", "[]]")
    s = Replace(s, "*", "[*]")
    s = Replace(s, "?", "[?]")
    s = Replace(s, "#", "[#]")
    EscapeAccessWildcards = s
End Function


' ===========================================
' SQL_String
'
' Builds a WHERE clause for text fields.
'
' Behavior:
'   - If IsValid = False ? skip field (Exit Function)
'   - If empty           ? skip field
'
'   PLAIN VALUE MODE:
'       User enters:   john
'       SQL produced:   LIKE '*john*'
'       ? Auto "contains" search
'       ? Access wildcards are escaped
'
'   PATTERN MODE:
'       User enters:   ~jo*n
'       SQL produced:   LIKE 'jo*n'
'       ? "~" prefix removed
'       ? Wildcards NOT escaped
'       ? Allows advanced patterns like ~[A-Z]* or ~??n*
'
' Parameters:
'   fieldName - table column name
'   state     - ControlState from GetString()
'
' Returns:
'   Partial SQL (e.g., "Name LIKE '*john*'")
'   or "" if no filtering needed
' ===========================================
Public Function SQL_String(fieldName As String, state As ControlState) As String
    If state.IsValid = False Then Exit Function
    If state.ConvertedValue = "" Then Exit Function

    Dim s As String
    s = CStr(state.ConvertedValue)

    ' Always escape apostrophes first
    s = Replace(s, "'", "''")

    If Left$(s, 1) = "~" Then
        ' PATTERN MODE — use user pattern exactly
        s = Mid$(s, 2) ' remove "~"
        SQL_String = fieldName & " LIKE '" & s & "'"

    Else
        ' PLAIN VALUE MODE — escape wildcards & auto contains search
        s = EscapeAccessWildcards(s)
        SQL_String = fieldName & " LIKE '*" & s & "*'"
    End If
End Function


' ===========================================
' SQL_Number
'
' Builds a WHERE clause for numeric fields.
'
' Behavior:
'   - If IsValid = False ? skip
'   - If Null value ? skip
'
' Produces:
'   Age = 25
'
' Parameters:
'   fieldName - table column name
'   state     - ControlState from GetInt() or GetDouble()
' ===========================================
Public Function SQL_Number(fieldName As String, state As ControlState) As String
    If state.IsValid = False Then Exit Function
    If IsNull(state.ConvertedValue) Then Exit Function

    SQL_Number = fieldName & " = " & state.ConvertedValue
End Function


' ===========================================
' SQL_Date
'
' Builds a WHERE clause for date fields.
'
' Produces:
'   HireDate = #2024-01-01#
'
' Uses ISO yyyy-mm-dd format (recommended).
'
' Parameters:
'   fieldName - table column name
'   state     - ControlState from GetDate()
' ===========================================
Public Function SQL_Date(fieldName As String, state As ControlState) As String
    If state.IsValid = False Then Exit Function
    If IsNull(state.ConvertedValue) Then Exit Function

    SQL_Date = fieldName & " = #" & Format(state.ConvertedValue, "yyyy-mm-dd") & "#"
End Function


' ===========================================
' SQL_Bool
'
' Builds WHERE clause for boolean fields.
'
' Produces:
'   IsActive = True
'   or
'   IsActive = False
'
' Parameters:
'   fieldName - table column name
'   state     - ControlState from GetBool()
' ===========================================
Public Function SQL_Bool(fieldName As String, state As ControlState) As String
    If state.IsValid = False Then Exit Function
    If IsNull(state.ConvertedValue) Then Exit Function

    If state.ConvertedValue = True Then
        SQL_Bool = fieldName & " = True"
    Else
        SQL_Bool = fieldName & " = False"
    End If
End Function


' ===========================================
' SQL_Join
'
' Combines multiple SQL filters into a single WHERE clause.
'
' Rules:
'   - SQL_xxx functions always return "" when unused.
'   - This function ignores empty strings.
'   - Automatically inserts "AND" between conditions.
'
' Example:
'   SQL_Join("Name LIKE '*john*'", "", "Age = 30")
'
' Returns:
'   (Name LIKE '*john*') AND (Age = 30)
'
' Parameters:
'   list() - ParamArray of SQL filters
'
' Returns:
'   Single WHERE clause string
' ===========================================
Public Function SQL_Join(ParamArray list()) As String
    Dim result As String
    Dim i As Long
    Dim item As String

    For i = LBound(list) To UBound(list)
        item = list(i)   ' always a String

        If item <> "" Then
            If result <> "" Then result = result & " AND "
            result = result & "(" & item & ")"
        End If
    Next i

    SQL_Join = result
End Function
```

# Multi-records/Datasheet Form – Form Code Module VBA

## Description

- This form is bound to a **multi-record dataset** (table or query).
- The form is displayed in **datasheet view**.
- When the user **double-clicks** a row, an Edit Form opens and the selected record's primary key (`ApplicantID`) is passed using a `WHERE` clause.

### Assumptions

- The form displays **multiple records**
- The user double-clicks a row to edit that specific record
- The key field **ApplicantID** exists in the form's RecordSource
- The key can be numeric or string.

---

## Code

```
Option Compare Database
Option Explicit

' --------------------------------------------------------------
' Form_DblClick
'
' Purpose:
'   When the user double-clicks any row in the datasheet,
'   this procedure opens the Edit Form and loads the selected
'   record by passing a WHERE condition based on ApplicantID.
'
' Notes:
'   - ApplicantID must exist in this form's recordset
'   - The Edit Form must be named "ReturnTracking_EditForm"
'   - The WHERE condition determines which record is loaded
' --------------------------------------------------------------
Private Sub Form_DblClick(Cancel As Integer)

    Dim whereCondition As String

    ' ------------------------------------------------------
    ' If ApplicantID is NUMERIC (most common case)
    ' ------------------------------------------------------
    whereCondition = "ApplicantID = " & Me!ApplicantID


    ' ------------------------------------------------------
    ' If ApplicantID is STRING / TEXT instead:
    '   (Uncomment and use this version)
    '
    '   whereCondition = "ApplicantID = '" & Me!ApplicantID & "'"
    '
    ' IMPORTANT:
    '   - String values MUST be wrapped in single quotes.
    '   - Do NOT wrap numeric values in quotes.
    ' ------------------------------------------------------


    ' Open the edit form and pass the WHERE condition
    If FormExists("ReturnTracking_EditForm") Then
        DoCmd.OpenForm "ReturnTracking_EditForm", , , whereCondition
    End If

End Sub
```

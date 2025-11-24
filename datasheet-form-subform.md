# Datasheet SubForm – VBA Module

## Description

This template is for a datasheet-style subform.  
When the user **double-clicks** a row, an Edit Form opens and the selected record’s primary key (`ApplicantID`) is passed using a `WHERE` clause.

### Assumptions

- The subform is bound to **multiple records**
- The user double-clicks a row to edit that specific record
- The key field **ApplicantID** exists in the form's RecordSource
- `ApplicantID` is usually numeric, but notes are included for string keys

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
    DoCmd.OpenForm "ReturnTracking_EditForm", , , whereCondition

End Sub
```

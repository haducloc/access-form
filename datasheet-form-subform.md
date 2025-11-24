# Datasheet SubForm – VBA Module

## Description

This template is for a datasheet-style subform.  
Double-clicking a record opens an Edit Form and passes the primary key (`ApplicantID`) using a `WHERE` clause.

## Code

```
Option Compare Database
Option Explicit

' Template: Datasheet SubForm
'
' Assumptions:
'   - The form is bound to multiple records
'   - Double-clicking a row opens the Edit Form
'   - ApplicantID is the primary key

Private Sub Form_DblClick(Cancel As Integer)

    Dim whereCondition As String

    ' ------------------------------------------------------
    ' If ApplicantID is NUMERIC (most common)
    ' ------------------------------------------------------
    whereCondition = "ApplicantID = " & Me!ApplicantID

    ' ------------------------------------------------------
    ' If ApplicantID is STRING/TEXT instead:
    '
    '   whereCondition = "ApplicantID = '" & Me!ApplicantID & "'"
    '
    ' (String values must be wrapped in single quotes)
    ' ------------------------------------------------------

    ' Open the edit form
    DoCmd.OpenForm "ReturnTracking_EditForm", , , whereCondition

End Sub
```

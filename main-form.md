# Main/App Form – Custom Form - Form Code Module VBA

## Description

- This form allows users to **search** for existing records and **add new** records.
- It contains two main sections:
  - **Top section** → Search inputs (ApplicantID, CtName, etc.)
  - **Bottom section** → A subform that displays multiple records as the search results.
- On the **Search** button click, you may also perform input validation.  
  For example, if ApplicantID must be numeric, you can validate the input and display an error message if it is not a valid number.

### Controls Used

- **Search button** → `btnSearch`
- **Add New button** → `btnAddNew`
- **Subform control** → `ReturnTrackingQuery_SubForm`  
  (This is the **subform CONTROL name**, not the form object name.)

---

## Code

```
Option Compare Database

' --------------------------------------------------------------
' btnAddNew_Click
'
' Purpose:
'   Opens the edit form in "Add New" mode so the user can create
'   a new return tracking record.
' --------------------------------------------------------------
Private Sub btnAddNew_Click()

    DoCmd.OpenForm "ReturnTracking_EditForm", , , , acFormAdd

End Sub


' --------------------------------------------------------------
' btnSearch_Click
'
' Purpose:
'   Builds a filter based on the user's search inputs and applies
'   that filter to the subform to display matching records.
'
' Search Fields:
'   txtAppID   - ApplicantID (numeric)
'   txtCtName  - CtName (text)
'
' Subform:
'   ReturnTrackingQuery_SubForm - Displays the search results
' --------------------------------------------------------------
Private Sub btnSearch_Click()

    Dim filter As String
    Dim subformName As String

    ' Subform control name on the main form
    subformName = "ReturnTrackingQuery_SubForm"

    ' Start with no filter
    filter = ""

    ' --------------------------------------------------------
    ' Build the filter conditions
    ' --------------------------------------------------------

    ' Filter by ApplicantID (numeric)
    If Nz(Me.txtAppID, "") <> "" Then
        filter = filter & "[ApplicantID] = " & Me.txtAppID
    End If

    ' Filter by CtName (text with wildcard search)
    If Nz(Me.txtCtName, "") <> "" Then

        ' Add AND if there is already another filter condition
        If filter <> "" Then
            filter = filter & " AND "
        End If

        filter = filter & "[CtName] LIKE '*" & Me.txtCtName & "*'"
    End If

    ' --------------------------------------------------------
    ' Apply or clear the filter on the subform
    ' --------------------------------------------------------
    If filter = "" Then

        ' No criteria → show all records
        Me(subformName).Form.FilterOn = False

    Else

        ' Apply the filter and enable filtering
        Me(subformName).Form.Filter = filter
        Me(subformName).Form.FilterOn = True

    End If

End Sub
```

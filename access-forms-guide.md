# Access Forms for Managing Records with Custom Search

- Includes a **main form** with a search section and a results grid (implemented as a **datasheet subform**).  
- Includes an **edit form** for adding new records or modifying existing ones.  
- Includes a **datasheet form** used to display multiple records and embedded as a subform in the main form.  
- Assumes the supporting table already exists. In this guide, the table is **ReturnTracking**, which contains fields such as `ApplicantID` (PK), `CtName`, and others.

---

# Step 1: Create a Query for the Dataset

- Purpose: To provide the recordSource for the subform on the **main form**.  
- Determine the tables involved. In this guide, only the **ReturnTracking** table is used.  
- Click SQL Query, Write SQL Query, Apply SQL sorting as needed.
- This query must have **no parameters**.
- Property Sheet, Set ODBC Timeout to 600 ( 10 minutes) if link tables/views used.
- Save the query as **ReturnTrackingQuery**.

---

# Step 2: Create the Datasheet Form

- Select the **ReturnTrackingQuery** view
- Create a datasheet-style form using **More Forms → Datasheet**.  
- The RecordSource will be **ReturnTrackingQuery**
- Save the form. In this guide, name it **ReturnTrackingQuery_SubForm**.  
- Open the form to verify that it displays correctly.  
- Switch to **Design View** and adjust its properties:

### Data Tab  
- Data Entry: **No**  
- Allow Additions: **No**  
- Allow Deletions: **No**  
- Allow Edits: **No**  
- Allow Filters: **No**

### Format Tab  
- Record Selectors: **No**  
- Navigation Buttons: **Yes**

### Form Code Module
- On Design View, Click on View Code menu on top right navigation.
- Open the form’s code module and apply the **datasheet-form code template**.  
- Modify the code to match your field names, control names, validation rules, and form names.

---

# Step 3: Create the Edit Form

- Create a **Single Form** using the Form Wizard.  
- Select the **ReturnTracking** table and include all necessary fields.  
- Choose the **Columnar** layout.  
- Save the form as **ReturnTracking_EditForm**.  
- Open the form to confirm it displays properly.  
- Switch to **Design View**, Adjust form size to have bottom space for two command buttons:
  - Save button → name **btnSave**
  - Delete button → name **btnDelete**
  - Notes: When drag a Button into the form, click Cancel to indicate this button will handled manually.

### Data Tab  
- Data Entry: **No**  
- Allow Additions: **Yes**  
- Allow Deletions: **Yes**  
- Allow Edits: **Yes**  
- Allow Filters: **No**

### Format Tab  
- Record Selectors: **No**  
- Navigation Buttons: **No**
### Other Tab  
- Pop Up: **Yes**  
- Modal: **Yes**
- Shortcut Menu: **No**

### Form Code Module
- Open the form’s code module and apply the **edit-form code template**.  
- Adjust the code as needed to match your fields, controls, and validation rules.

---

# Step 4: Create the Main Form

- Create a **Blank Form** (Single Form).  
- Add search input controls to the top of the form. Use consistent names such as:  
  - `txtCtName`, `cboReturnType`, `chkReturnComplete`, etc.

- Add command buttons for actions:  
  - Search button → caption **Search**, name **btnSearch**  
  - Add New button → caption **Add New**, name **btnAddNew**

- Drag the datasheet form (**ReturnTrackingQuery_SubForm**) onto the bottom section of the main form.  
  - Access will embed it inside a subform control automatically.  
  - Name the subform control. By default it use form object name as the subform control name.

- Open the form to ensure the subform loads correctly.  
- Switch to **Design View** and update the properties:

### Data Tab  
- Data Entry: **No**  
- Allow Additions: **No**  
- Allow Deletions: **No**  
- Allow Edits: **Yes**  
- Allow Filters: **No**

### Format Tab  
- Record Selectors: **No**  
- Navigation Buttons: **No**
  
### Other Tab  
- Shortcut Menu: **No**

### Form Code Module
- Open the form’s code module and apply the **main-form code template**.  
- Adjust the code as needed for your search fields, subform control name, and validation logic.

---

# NOTES: VBA — When to Use `.` vs `!`

## `.` (Dot)
Use dot for:
- Properties  
- Methods  
- Anything supported by IntelliSense

Examples:
    Me.Requery  
    Form.TextBox1.Value  
    rs.MoveNext

## `!` (Bang)
Use bang for:
- Controls referenced by name  
- Fields in recordsets  
- Collection items looked up at runtime

Examples:
    Me!TextBox1  
    rs!FirstName  
    Forms!OrderForm!Total

## Quick Rule
- `.` = known members (compile-time)  
- `!` = named items (runtime lookup)

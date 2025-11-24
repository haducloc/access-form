# Access Forms for Managing Records with Custom Search

- Includes a **main form** containing a search section and a results grid (implemented as a **datasheet subform**).  
- Includes an **edit form** for adding new records or editing existing ones.  
- Includes a **datasheet form** used to display multiple records and embedded as a subform in the main form.  
- Assumes the supporting table already exists. In this guide, the table is **ReturnTracking**, which contains fields such as `ApplicantID` (PK), `CtName`, and others.

---

# Step 1: Create a Query for the Dataset

- Create a query that will supply the dataset for the records.  
- You may join other tables if needed, but for this example the query is based solely on the **ReturnTracking** table.  
- You may apply SQL sorting as needed.  
- This query must have **no parameters**.  
- Save the query with the name **ReturnTrackingQuery**.

---

# Step 2: Create the Datasheet Form

- Create a datasheet-style form using **More Forms → Datasheet**.  
- Set the RecordSource to the query **ReturnTrackingQuery**.  
- Include all fields or selected fields as needed.  
- Save the form. In this guide, name it **ReturnTrackingQuery_Subform**.  
- Open the form to verify correct behavior.  
- Switch to **Design View** and adjust properties:

### Data Tab  
- Data Entry → **No**  
- Allow Additions → **No**  
- Allow Deletions → **No**  
- Allow Edits → **No**  
- Allow Filters → **No**

### Format Tab  
- Record Selectors → **No**  
- Navigation Buttons → **No**

- Open the form’s code module and apply the **datasheet-form code template**.  
- Modify the code if needed so control names match your fields or validation rules.

---

# Step 3: Create the Edit Form

- Create a **Single Form** using the Form Wizard.  
- Select the **ReturnTracking** table and include all fields.  
- Choose **Columnar** layout.  
- Save the form. For this guide, name it **ReturnTracking_EditForm**.  
- Open the form to verify correct behavior.  
- Switch to **Design View** and add two command buttons at the bottom of the form:
  - Save button → name it **btnSave**  
  - Delete button → name it **btnDelete**

### Data Tab  
- Data Entry → **No**  
- Allow Additions → **No**  
- Allow Deletions → **No**  
- Allow Edits → **No**  
- Allow Filters → **No**

### Format Tab  
- Record Selectors → **No**  
- Navigation Buttons → **No**

- Open the form’s code module and apply the **edit-form code template**.  
- Modify the code if needed to match field names, validation rules, and control names.

---

# Step 4: Create the Main Form

- Create a **Blank Form** (Single Form).  
- Add search input controls to the top area of the form and name them using a consistent convention, such as:  
  - `txtUserName`  
  - `cboUserType`  
  - `chkUserActive`  
  - `subformUserList` (for the subform control)

- Add command buttons for search and add-new operations:  
  - Search button → caption **Search**, name **btnSearch**  
  - Add New button → caption **Add New**, name **btnAddNew**

- Drag the datasheet form (**ReturnTrackingQuery_Subform**) onto the bottom of the main form.  
  - Access will automatically embed it inside a subform control.  
  - Name the subform control appropriately (usually the same name).

- Open the form to verify the subform loads correctly.

- Switch to **Design View** and update the properties:

### Data Tab  
- Data Entry → **No**  
- Allow Additions → **No**  
- Allow Deletions → **No**  
- Allow Edits → **No**  
- Allow Filters → **No**

### Format Tab  
- Record Selectors → **No**  
- Navigation Buttons → **No**

- Open the main form’s code module and apply the **main-form code template**.  
- Modify names and logic as needed to match your search fields, subform control name, and validation rules.


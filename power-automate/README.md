# Power Automate Approval Flow Specs

This directory contains the automation specifications and step-by-step blueprints for the single-level Manager Approval and PDF Generation flow associated with the Travel Request App.

---

## Flow Trigger & Architecture

- **Trigger:** SharePoint Online — *When an item is created* in the `Travel Request` list.
- **Auditing & Status Update:** Writes approval comments, responder details, and approval dates back to SharePoint.
- **Dynamic PDF Generation:** Pulls child cost line items, converts them to a styled PDF using OneDrive, saves it in SharePoint, and emails the PDF as an attachment to the submitter upon approval.

---

## Step-by-Step Flow Blueprint

### 1. Trigger
- **Connector:** SharePoint
- **Action:** When an item is created
- **Site Address:** *Your SharePoint Site URL*
- **List Name:** `Travel Request`

### 2. Initialize Variables
- **`RequestID`** (String): Holds the custom generated tracking code (e.g. `TRV-2026-105`).
- **`ManagerEmail`** (String): Initially set to the `Approver` field value from the trigger output (`triggerOutputs()?['body/Approver']`).
- **`Status`** (String): Tracks the current execution status.

### 3. Get Manager Details (Fallback Backup)
- A conditional check evaluates if the traveler's manager email is empty. If so, it queries Office 365 Users to resolve the manager's UPN and set the `ManagerEmail` variable.

### 4. Create Approval Card
- **Connector:** Approvals
- **Action:** Start and wait for an approval
- **Approval Type:** `Approve/Reject - First to respond`
- **Assigned To:** `variables('ManagerEmail')`

### 5. PDF Generation & SharePoint Upload (On Approved Path)
1. **Get items (Child Details):** Queries the child details list where `TravelRequest/Id` matches the parent `ID`.
2. **Select columns:** Map Date, Title, Category, Description, and Amount.
3. **Create HTML table:** Generate the raw HTML table of mapped expense items.
4. **Get items (Admin Options):** Query company branding options where `App eq 'TravelRequest'`.
5. **Compose HTML Content:** Standard CSS styling injected with traveler meta information and itemized cost breakdown tables.
6. **Create temp HTML file (OneDrive):** Save output to `/temp` folder in OneDrive.
7. **Convert file (OneDrive):** Convert the OneDrive HTML file to PDF.
8. **Create file (SharePoint):** Save the converted PDF in SPO `Travel Request Documents/GeneratedPDF`.
9. **Delete temp file (OneDrive):** Clean up the OneDrive temp HTML file.
10. **Send email (Outlook V2):** Send approval notification with the converted SPO PDF file as an attachment.

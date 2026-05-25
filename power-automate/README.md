# Power Automate Approval Flow Specs

This directory contains the automation specifications and step-by-step blueprints for the single-level Manager Approval and PDF Generation flow associated with the Travel Request App.

> [!TIP]
> **Complete Configuration Instructions:**
> For detailed step-by-step setup guides on configuring SharePoint Online lists, compiling the SPFx web part, and deploying this Power Automate flow, refer to the parent [Step-by-Step-guide.md](../Step-by-Step-guide.md) in the project root.

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
- **`EmailAttachments`** (Array): Holds the array of dynamic email attachments (PDF on approval, empty on rejection).
- **`TravelStatus`** (String): Stores the runtime status (`Approved` or `Rejected`) of the request.

### 3. Fetch & Format Details
1. **Get items (Child Details):** Queries the child details list where `TravelRequestID` matches the parent `ID`.
2. **Select columns (PDF):** Map Date, Title, Category, Description, and Amount for the PDF report.
3. **Create HTML table:** Generate the raw HTML table of mapped expense items.
4. **Select column (Markdown):** Format child expense items as Markdown table rows.
5. **Join rows & Compose Markdown table:** Combine the rows to form the Markdown details table.

### 4. Create Approval Card
- **Connector:** Approvals
- **Action:** Start and wait for an approval
- **Approval Type:** `Approve/Reject - First to respond`
- **Assigned To:** `triggerBody()?['Approver']`
- **Details:** Includes general traveler details and the dynamically formatted Markdown details table output.

### 5. Conditional Branches & Notifications
- **Approved Path (`If yes`):**
  1. **Update item status:** Set status to `Approved`.
  2. **Set variable:** Set `TravelStatus` variable to `Approved`.
  3. **Get items (Admin Options):** Query company branding options where `App eq 'TravelRequest'`.
  4. **Compose HTML Content:** Standard CSS styling injected with traveler meta information and itemized cost breakdown tables.
  5. **Create temp HTML file (OneDrive):** Save output to `/temp` folder in OneDrive.
  6. **Convert file (OneDrive):** Convert the OneDrive HTML file to PDF.
  7. **Delete temp file (OneDrive):** Clean up the OneDrive temp HTML file.
  8. **Append to array variable (EmailAttachments):** Add the PDF attachment name and content bytes to the array.
- **Rejected Path (`If no`):**
  1. **Update item status:** Set status to `Rejected`.
  2. **Set variable:** Set `TravelStatus` variable to `Rejected`.

### 6. Dynamic Notifications & Audit (Outside Condition)
- **Send email (Outlook V2):** Sent to the traveler using the dynamic `TravelStatus` in the subject/body, with the HTML details table in the body, and the `EmailAttachments` array variable mapped to the attachments parameter.
- **Log Approval Details (Update Item):** Write responder email, comments, and decision date back to the SharePoint Parent item.
- **Terminate:** Stop the flow execution with a status of `Succeeded`.

### 7. Optional Enhancements & Extensions
- **Microsoft Teams Integration**: Post a Teams message card dynamically displaying the approval status and comments.
- **Try-Catch Error Scoping**: Wrap core actions in a try-catch scope to handle and log runtime exceptions to the `Flow Error Logs` list.
- **Font Icons in PDF**: Load CDN icon libraries and dynamically render font icons (e.g. `plane` or `briefcase`) from the branding table in the PDF report header.

---

## 👨‍💻 Developer & Credits

This automation suite was designed and implemented by **WRVishnu**. For more enterprise-level automation templates, customized SPFx components, and Microsoft 365 cloud engineering solutions, visit **[www.wrvishnu.com](http://www.wrvishnu.com)**.



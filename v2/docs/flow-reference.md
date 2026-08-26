# Power Automate — flow reference

Flow name: **Wayleave application**
Connections: Microsoft Forms · Excel Online (Business) · Office 365 Outlook

## Actions

### 1. When a new response is submitted
- Form Id: **Wayleave Submission**

Re-select the form here whenever you add or rename a form question, otherwise new
questions appear in dynamic content as raw ids (`body/ra41a046…`).

### 2. Get response details
- Form Id: **Wayleave Submission**
- Response Id: `Response Id` from the trigger

### 3. Run script
- Location: OneDrive for Business
- File: `/Wayleave_Application_Tracker.xlsx`
- Script: `WayleaveUpsert`

| Parameter | Mapped to |
|---|---|
| submissionType | Submission Type |
| projectRef | Project Ref / Submission Ref (Q2) **+** Project Ref / Submission Ref (Q10) |
| projectTitle | Project Title |
| wayleaveType | Wayleave Type |
| district | District |
| agency | Agency |
| jkrRequired | JKR required? |
| applicationLink | Application File Submission |
| approvalRef | Approval Ref |
| approvalLink | Approval File Submission |
| notifyEngineers | Notify engineer (Q9) **+** Notify engineer (Q13) |
| submittedOn | `utcNow()` |

Two chips in one box is intentional: only one section is ever filled per submission,
so concatenating the pair yields whichever was answered.

Do **not** map the entries named "Application" or "Approval" — those are section
headings and are always empty.

### 4. Apply to each
- Select an output: `engineerEmails` from **Run script**

An empty array means zero iterations — no email, no failure. This is why there is
no Condition guarding the loop.

#### 4a. Send an email (V2)  *(inside the loop)*
- To: `Current item`
- Subject: `Wayleave Application - Action Required`
- Body: see `docs/email-template.html`

## Reading a failed run

Open the run, click **Run script**, scroll to **OUTPUTS**:

- `note` is populated → a table or row could not be found; the message says which
- `engineerEmails` is empty → no engineer selected, or a name did not match the lookup table
- `action` = `Approval - No Matching Project` → approval submitted for a Project Ref not in the tracker

`Send an email (V2)` reporting *"there are no items to repeat"* is not an error —
it means `engineerEmails` came back empty.

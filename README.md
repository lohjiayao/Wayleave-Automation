# Wayleave Application Tracker — Automation

Automates wayleave application and approval tracking at Sarawak Energy. A technician submits a Microsoft Form; Power Automate runs an Office Script that writes to Excel and works out who to notify; matched engineers get an email with the project details.

No manual data entry, no manual notification emails, and no duplicate rows when a project is resubmitted.

## Demo

[![Wayleave Tracker automation demo](https://img.youtube.com/vi/DmcPfcErNbY/hqdefault.jpg)](https://youtu.be/DmcPfcErNbY)

▶ **[Watch the demo](https://youtu.be/DmcPfcErNbY)** — form submission through to tracker update and engineer notification.

---

## How it works

```
Microsoft Form submitted
        │
        ▼
Power Automate trigger  ──►  Get response details
        │
        ▼
Run script (WayleaveUpsert.ts)          ← all Excel logic lives here
        │   • decides Application vs Approval
        │   • adds or updates the tracker row (no duplicates)
        │   • writes the audit log row
        │   • resolves engineer names → email addresses
        │   • returns project context for the email
        ▼
Apply to each  (over engineerEmails)
        └── Send an email (V2)          ← one email per selected engineer
```

Four actions in the flow. Everything that could go wrong in Excel is handled inside one script instead of a chain of Filter array / Condition / Update row / Add row steps.

---

## Components

### 1. Microsoft Form — "Wayleave Submission"

One form, two sections, branched by the first question.

| Q | Question | Type | Section |
|---|---|---|---|
| 1 | Submission Type | Choice: Application / Approval | — |
| 2 | Project Ref / Submission Ref | Text | Application |
| 3 | Project Title | Text | Application |
| 4 | Wayleave Type | Choice | Application |
| 5 | District | Choice | Application |
| 6 | Agency | Choice | Application |
| 7 | JKR required? | Choice: Yes / No | Application |
| 8 | Application File Submission | File upload | Application |
| 9 | Notify engineer | Multi-choice | Application |
| 10 | Project Ref / Submission Ref | Text | Approval |
| 11 | Approval File Submission | File upload | Approval |
| 12 | Approval Ref | Text | Approval |
| 13 | Notify engineer | Multi-choice | Approval |

Nothing is marked Required. If no engineer is selected, no email is sent — the flow handles that without erroring.

### 2. Excel workbook — `Wayleave_Application_Tracker.xlsx`

Three Excel Tables. The script finds them by inspecting headers, not by table name, so renaming a table won't break anything.

| Table | Identified by | Purpose |
|---|---|---|
| Tracker | has a `Project Ref…` column and an `Approval…` column | One row per project. Current state. |
| Submission Log | has a `Timestamp` column | Every submission, append-only audit trail. |
| Engineer lookup | has an `Engineer Email` column | Maps form option values → real email addresses. |

**Engineer lookup table** — the header must literally contain `Engineer Email`, and the name values must match the form's option values (matching is case-insensitive):

| Engineer Name | Engineer Email | Section |
|---|---|---|
| loh | name@example.com | Distribution Planning |
| guofu | name@example.com | Distribution Planning |
| jiayao | name@example.com | Projects |

To add or remove an engineer, edit this table and the form's choice options. The flow needs no changes.

### 3. Office Script — `v2/scripts/WayleaveUpsert.ts`

Lives in the workbook under **Automate**. Takes 12 parameters, returns a result object.

**Inputs**

| Parameter | Source |
|---|---|
| submissionType | Q1 |
| projectRef | Q2 **and** Q10 (both chips in one box — only one is ever filled) |
| projectTitle | Q3 |
| wayleaveType | Q4 |
| district | Q5 |
| agency | Q6 |
| jkrRequired | Q7 |
| applicationLink | Q8 |
| approvalRef | Q12 |
| approvalLink | Q11 |
| notifyEngineers | Q9 **and** Q13 (both chips in one box) |
| submittedOn | `utcNow()` |

**Outputs**

| Field | Meaning |
|---|---|
| `action` | Added / Updated / Approval Updated / Approval - No Matching Project |
| `matched` | whether an existing tracker row was found |
| `engineerEmails` | addresses to email — drives the Apply to each loop |
| `engineerNames` | the names that resolved |
| `documentLink` | clean URL, JSON wrapper stripped |
| `documentName` | original filename |
| `outProjectRef` … `outApprovalRef` | project details read back from the tracker row |
| `note` | diagnostic message when a table or row can't be found |

**Behaviour**

- *Application, new Project Ref* → new tracker row
- *Application, existing Project Ref* → updates that row in place, no duplicate
- *Approval, existing Project Ref* → sets only Approval Ref, Approval Date, Document Link - Approval; leaves everything else untouched
- *Approval, unknown Project Ref* → writes nothing, explains why in `note`
- *No engineer selected* → `engineerEmails` is empty, loop runs zero times, no email

Every submission gets a log row regardless of outcome.

---

## Setup

1. **Excel** — open the workbook → **Automate** → **New Script** → paste `v2/scripts/WayleaveUpsert.ts` → name it `WayleaveUpsert` → Save.
2. **Verify tables** — three Excel Tables must exist with the headers described above. The engineer table's email column header must contain `Engineer Email`.
3. **Power Automate** — build four actions:
   - When a new response is submitted (Forms)
   - Get response details (Forms)
   - Run script (Excel Online Business) → select the workbook and `WayleaveUpsert`, map the 12 parameters
   - Apply to each over `engineerEmails` → Send an email (V2), To = Current item
4. **Email body** — paste `v2/docs/email-template.html` via the code-view (`</>`) button, then replace each placeholder with the matching **Run script** output.

---

## Gotchas

**Form file uploads arrive as JSON**, not a URL:

```json
[{"name":"file.jpg","link":"https://...","id":"...","size":35491, ...}]
```

The script strips this to the bare link. Use the script's `documentLink` output in the email, never the raw form field.

**Multi-select answers are text, not arrays.** Forms may send `["guofu","loh"]`, `guofu;loh`, or `guofu, loh`. The script accepts all three plus empty. Do not use `split()` in Power Automate — it throws on null.

**Section headings appear in the dynamic content list.** "Application" and "Approval" show up alongside real questions and are always empty. Don't map them.

**New form questions show as raw IDs** (`body/ra41a046…`) until the flow re-reads the form schema. Re-select the form in the trigger to refresh.

**Two questions can share a name.** There are two "Project Ref / Submission Ref" and two "Notify engineer" entries. Put *both* chips in the same parameter box — Power Automate concatenates them, and since only one is ever filled you get the filled one.

---

## Testing

| Case | Expected |
|---|---|
| Application, new Project Ref, engineer selected | new tracker row · log row · email sent |
| Application, same Project Ref again | tracker row updated, not duplicated · new log row |
| Application, no engineer selected | tracker written · log written · no email, no error |
| Approval, existing Project Ref | only approval columns change · email shows full project details |
| Approval, unknown Project Ref | nothing written · `note` explains why |

---

## Repository layout

```
.
├── README.md
├── v1/                         earlier version — see v1/CHANGELOG.md
│   ├── CHANGELOG.md
│   ├── POWER_AUTOMATE_SETUP.md
│   ├── TROUBLESHOOTING.md
│   └── WORKBOOK_STRUCTURE.md
└── v2/                         current version, described above
    ├── scripts/
    │   └── WayleaveUpsert.ts   Office Script — all Excel logic
    └── docs/
        ├── email-template.html notification email body
        └── flow-reference.md   Power Automate parameter mapping
```

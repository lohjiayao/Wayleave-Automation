# Power Automate Setup

## Flow structure

```
When a new response is submitted          (Microsoft Forms)
│
Get response details                      (Microsoft Forms)
│
Run script  →  UpsertWayleaveRow          (Excel Online Business)
│
Condition   →  length(coalesce(...emails, '')) is greater than 0
│
├── True ──── Send an email (V2)
│             ├── Run script → MarkNotified   status = "Sent"
│             └── Run script → MarkNotified   status = "Failed"
│                 (Run after: has failed / has timed out)
│
└── False ─── Send an email (V2)   [alert to yourself]
              └── Run script → MarkNotified   status = "No recipients"
```

There is deliberately **no** `Add a row into a table` action. See the design
notes in the README.

---

## Action configuration

### Run script — UpsertWayleaveRow

Map each script parameter to the matching field from *Get response details*:

| Script parameter  | Form question                  |
|-------------------|--------------------------------|
| `projectRef`      | Project Ref / Submission Ref   |
| `projectTitle`    | Project Title                  |
| `wayleaveType`    | Wayleave Type                  |
| `district`        | District                       |
| `jkrRequired`     | JKR Required?                  |
| `agency`          | Agency                         |
| `documentLink`    | Document Link                  |
| `notifyEngineers` | Notify engineer                |

The script returns:

| Key         | Meaning                                          |
|-------------|--------------------------------------------------|
| `action`    | `inserted` or `updated`                          |
| `row`       | Sheet row number in Wayleave Tracker             |
| `logRow`    | Sheet row number in Submission Log               |
| `ref`       | The trimmed Project Ref                          |
| `emails`    | Semicolon-joined email addresses, ready for `To` |
| `unmatched` | Names not found in EngineerList                  |

### Condition

| Left | Operator | Right |
|------|----------|-------|
| `length(coalesce(body('Run_script')?['result']?['emails'], ''))` | is greater than | `0` |

`coalesce` guards against a null result. Without it, `length(null)` raises
`InvalidTemplate` and the run terminates at the condition instead of taking the
False branch.

> **Note:** `body('Run_script')` uses the action's *internal* name, not its
> display name. If the action is renamed or a second Run script is added, the
> key becomes `Run_script_1`. Confirm in **Code view**.

### Send an email (V2) — True branch

| Field   | Value |
|---------|-------|
| To      | `body('Run_script')?['result']?['emails']` |
| Subject | `[Wayleave] @{body('Run_script')?['result']?['ref']}` |
| Body    | Project details from *Get response details*, plus the document link |

### Send an email (V2) — False branch

This is an alert to yourself, so the **To** field must be a literal address.
Leaving it empty causes Power Automate to substitute `anonymous`, which fails
with `expected format 'string/email'`.

| Field   | Value |
|---------|-------|
| To      | *(your own address, typed literally)* |
| Subject | `[Wayleave] No recipients for @{body('Run_script')?['result']?['ref']}` |
| Body    | `Unmatched names: @{body('Run_script')?['result']?['unmatched']}` |

### Run script — MarkNotified

| Script parameter | Value |
|------------------|-------|
| `projectRef`     | `body('Run_script')?['result']?['ref']` |
| `status`         | `Sent`, `Failed`, or `No recipients` depending on branch |
| `note`           | Optional. On failure paths, pass `unmatched` or the error. |

For the failure instance, right-click the action → **Configure run after** →
tick *has failed* and *has timed out*. Without this the flow simply terminates
on a failed send and the row is left reading `Pending`, which is exactly the
case where an accurate status matters most.

---

## Testing

Submit the same Project Ref twice. Expected result:

- 1 row in Wayleave Tracker, with the second submission's values
- 2 rows in Submission Log
- 2 emails sent
- Notification Status reading `Sent`, with a real date in the Sent Date column
  (right-aligned, not left-aligned — left alignment means it was written as text)

To verify the header-resolution logic, insert a dummy column in the middle of
the tracker, submit a test response, confirm the values still land in the
correct columns, then delete the dummy column.

# Workbook Structure

File: `Wayleave_Application_Tracker.xlsx` (stored in SharePoint)

The workbook itself is not committed to this repository — it contains staff
email addresses and internal SharePoint document links. This file documents its
structure so the scripts can be understood without it.

---

## Sheet: Wayleave Tracker

Table name: `WayleaveTracker`
One row per Project Ref. Updated in place on resubmission.

| Col | Header                        | Written by | Notes |
|-----|-------------------------------|------------|-------|
| A   | Project Ref / Submission Ref  | Script     | Match key. Case-insensitive. |
| B   | Project Title                 | Script     | |
| C   | Wayleave Type                 | Script     | |
| D   | District                      | Script     | |
| E   | Confirmation Date             | Manual     | Entered by the planner later. |
| F   | Submission Date               | Script     | Excel serial, `dd/mm/yyyy`. |
| G   | Submission < 10 Days          | Formula    | Calculated column. |
| H   | JKR Required                  | Script     | |
| I   | Agency                        | Script     | |
| J   | Approval Ref                  | Manual     | |
| K   | Approval Date                 | Manual     | |
| L   | JKR Duration (Days)           | Formula    | Calculated column. |
| M   | Remarks                       | Both       | MarkNotified appends here. |
| N   | Document Link - Application   | Script     | |
| O   | Document Link - Approval      | Manual     | |
| P   | Notify Engineer(s)            | Script     | Normalised to `name; name`. |
| Q   | Engineer Notified?            | MarkNotified | |
| R   | Notification Sent Date        | MarkNotified | Excel serial, `dd/mm/yyyy hh:mm`. |
| S   | Notification Status           | MarkNotified | `Pending` / `Sent` / `Failed`. |

### Calculated columns

```excel
G:  =IF(AND(E2<>"",F2<>""),IF((F2-E2)<10,"Yes","No"),"")
L:  =IF(AND(E2<>"",K2<>""),K2-E2,"")
```

Both are re-applied by the script after a row is inserted, because writing a
null into a calculated column can clear its formula. The script derives the
column letters from the live header positions, so the formulas survive a column
being moved.

---

## Sheet: Submission Log

Table name: `SubmissionLog`
Append-only. One row per submission, including resubmissions.

| Col | Header                       |
|-----|------------------------------|
| A   | Timestamp                    |
| B   | Project Ref / Submission Ref |
| C   | Project Title                |
| D   | Wayleave Type                |
| E   | District                     |
| F   | Agency                       |
| G   | JKR Required                 |
| H   | Document Link                |
| I   | Notify Engineer(s)           |

The tracker shows current state; the log shows what came in and when. Keeping
both means a resubmission does not erase the record of the original.

---

## Sheet: Lists

Dropdown sources plus the engineer directory.

| Range | Contents |
|-------|----------|
| A     | Wayleave Type options |
| B     | District (Sarawak) options |
| C     | JKR Required options |
| D     | Agency options |
| F:H   | `EngineerList` table — Name, Email, Section |

**Only `EngineerList` is read by the scripts.** Columns A–D exist purely for
Excel's data validation dropdowns, so options can be added or removed freely
without touching any code.

`EngineerList` is read fresh on every run, so adding, removing, or renaming an
engineer takes effect on the next submission with no code change. The `Name`
values must match the Form's choice labels exactly — matching is
case-insensitive and trims surrounding whitespace, but internal spacing counts.

---

## Sheet: Legend

Reference notes for users of the workbook. Not read by any script.

---

## Blank padding rows

Both `WayleaveTracker` and `SubmissionLog` were created with several hundred
pre-formatted blank rows. The scripts therefore fill the first blank row rather
than calling `addRow`, which would otherwise append below the padding and leave
a large empty gap. If the padding is ever removed, the scripts fall back to
`addRow` automatically.

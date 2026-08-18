## ▶️ [WATCH THE DEMO VIDEO](https://youtu.be/rpPtETyhqXs)

[![Watch the demo](https://img.youtube.com/vi/rpPtETyhqXs/maxresdefault.jpg)](https://youtu.be/rpPtETyhqXs)

---

# Wayleave Application Tracker

Automation for logging wayleave applications submitted through Microsoft Forms
into an Excel Online workbook, with duplicate handling and automatic engineer
notification by email.

Built during industrial training in the Distribution Planning department,
Sarawak Energy Berhad (Sibu).

---

## What it does

1. An engineer submits a wayleave application through a Microsoft Form.
2. Power Automate picks up the response and calls an Office Script.
3. The script writes to three places in one call:
   - **Wayleave Tracker** — one row per Project Ref. Existing refs are updated
     in place rather than duplicated, so resubmissions do not create clutter.
   - **Submission Log** — append-only audit trail. Every submission gets a row,
     including resubmissions.
   - **EngineerList** — read only, to turn selected engineer names into email
     addresses.
4. The flow emails the selected engineers, then writes the notification status
   back to the tracker.

---

## Repository structure

```
wayleave-tracker/
├── README.md
├── CHANGELOG.md
├── .gitignore
├── scripts/
│   ├── UpsertWayleaveRow.ts    Main script. Insert-or-update + log + email lookup.
│   └── MarkNotified.ts         Runs after the email. Writes back Q/R/S columns.
├── docs/
│   ├── POWER_AUTOMATE_SETUP.md Flow structure, actions, expressions.
│   ├── WORKBOOK_STRUCTURE.md   Sheets, tables, columns, formulas.
│   └── TROUBLESHOOTING.md      Known failure modes and how to diagnose them.
└── logbook/
    └── 2026-08-18.txt          Daily internship logbook entry.
```

---

## Requirements

- Excel workbook stored in SharePoint or OneDrive (Office Scripts do not run on
  local files)
- Power Automate with the Office 365 Outlook and Excel Online (Business)
  connectors
- A Microsoft Form with questions matching the script parameters

---

## Setup

1. Open the workbook in Excel Online, go to **Automate → New Script**, paste in
   `scripts/UpsertWayleaveRow.ts`, and save it as `UpsertWayleaveRow`.
2. Repeat for `scripts/MarkNotified.ts`.
3. Build the flow as described in `docs/POWER_AUTOMATE_SETUP.md`.
4. Confirm the engineer names in the Form's choice list match the `Name` column
   of the `EngineerList` table exactly.

---

## Design notes

**Why the script does the insert, not the "Add a row into a table" action.**
The script performs an insert-or-update. If a separate Add-a-row action runs
first, the row already exists by the time the script looks for it, so it can
never insert, and a resubmission produces two rows for the same reference.
The flow must not contain that action.

**Why dates are written as serial numbers.**
`setValue(new Date())` writes a text string such as `Tue Aug 18 2026 ...`.
The calculated columns subtract dates, so a text value gives `#VALUE!`.
The script converts to an Excel serial number (epoch 30 Dec 1899) and applies a
UTC+8 offset, because the Power Automate runtime is UTC and a late-afternoon
submission would otherwise be dated to the previous day.

**Why columns are resolved by header name.**
Hardcoded column indices break silently when someone inserts a column — writes
land one column off with no error. The script reads the header row at runtime
and builds a header-to-position map, so reordering and inserting columns is
safe. Only renaming a header requires a code change, and that raises a named
error rather than corrupting data.

---

## References

- **Microsoft Form (design view):** https://forms.cloud.microsoft/Pages/DesignPageV2.aspx?subpage=design&FormId=lWO9y-HouEez3XevwjPMVTjC9G2MWc9PrE4RzL3MpFtUMFJCQ0VDTVFOUVM2S0JVRDJETEtJTlMzOS4u&Token=e09d0e3bed304ae8b528820ef831370f
- **Excel workbook (SharePoint):** https://power2grow-my.sharepoint.com/:x:/r/personal/jiayao_loh_sarawakenergy_com/Documents/Wayleave_Application_Tracker.xlsx?d=wf234b2bd76b14fe894279d92e0b38614&csf=1&web=1&e=iGgPsx

---

## Known limitations

- Engineer name matching is case-insensitive and trims surrounding whitespace,
  but internal spacing must match (`jia yao` does not match `jiayao`).
- Deleting and recreating a question in the Form changes its underlying ID and
  silently breaks the corresponding flow parameter. Edit questions in place.
- Concurrent submissions are not locked. At the expected volume this is
  acceptable, but two submissions landing in the same second could theoretically
  target the same blank row.

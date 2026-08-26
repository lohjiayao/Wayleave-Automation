# Changelog

## v4 — Self-adjusting columns + Forms JSON parsing

- Columns are now resolved by header name at runtime instead of hardcoded
  indices. Inserting or reordering columns in Excel no longer causes silent
  misaligned writes.
- All sheet, table, and header names consolidated into a single `CONFIG` block.
- A missing header now raises a named error instead of corrupting data.
- Calculated-column formulas are rebuilt from live header positions, so they
  follow a column that has been moved.
- `splitNames()` handles Microsoft Forms' JSON array output
  (`["guofu","loh"]`) as well as plain semicolon or comma separated text.

## v3 — Submission Log + engineer email lookup

- Script now appends to `SubmissionLog` on every run, replacing the removed
  `Add a row into a table` action. Fills blank padding rows top-down rather
  than appending past row 500.
- Log timestamps keep hours and minutes; tracker dates remain date-only.
- Added `resolveEmails()`, reading `EngineerList` live on each run and
  returning a semicolon-joined `emails` string ready for the Send-an-email
  `To` field, plus an `unmatched` list for names that failed to resolve.
- Added `MarkNotified.ts` to write back Engineer Notified, Sent Date, and
  Status after the email action, including a failure path.

## v2 — Correctness fixes

- Dates written as Excel serial numbers instead of JS `Date` objects, fixing
  `#VALUE!` in the two calculated columns.
- UTC+8 offset applied, since the Power Automate runtime is UTC.
- Blank Project Ref now rejected — previously it matched the first blank
  padding row and overwrote it.
- Ref matching made case-insensitive.
- Null checks added on worksheet and table lookups.
- Calculated-column formulas re-applied after insert.
- Blank padding rows reused instead of appending past them.
- Sent-date no longer stamped before the email is actually sent.

## v1 — Initial script

- Insert-or-update against `WayleaveTracker` keyed on Project Ref.

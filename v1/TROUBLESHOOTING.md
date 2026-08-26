# Troubleshooting

Failure modes encountered while building this, and how to diagnose them.

---

## Duplicate rows appear for the same Project Ref

**Cause:** an `Add a row into a table` action running before the Run script
action. The script performs insert-or-update, so by the time it looks for the
ref, the row already exists — it can never insert, and a resubmission produces a
second row while the script updates the first.

**Fix:** delete that action. The script owns both tables.

---

## Engineers receive no email; the flow takes the False branch

**Cause (most common):** Microsoft Forms returns multi-select answers as a
**JSON array string**, not a semicolon-delimited list:

```
["guofu","loh","jiayao"]
```

Splitting that on `;` yields `["guofu"` and `"loh"]` — brackets and quotes
included — which match nothing in `EngineerList`.

**Fix:** `splitNames()` in `UpsertWayleaveRow.ts` detects a leading `[` and
parses the JSON, falling back to a plain split for single-select or manually
typed values.

**Cause (second):** the name in the Form does not match the `Name` column of
`EngineerList`. Matching ignores case and surrounding whitespace but not
internal spacing — `jiayao` will not match `jia yao`.

**How to diagnose:** open the run history, click the Run script action, and read
its raw **Outputs**. The `unmatched` field names exactly which entries failed to
resolve.

---

## `Input parameter 'emailMessage/To' ... expected format 'string/email'`

**Cause:** an empty **To** field. Power Automate substitutes `anonymous`, which
is not a valid address.

Typically seen on the False-branch alert email, where dynamic content was used
by mistake. That branch runs *because* no addresses were resolved, so the
address must be typed literally.

---

## Condition shows a red error but evaluated in ~0.1s

The Condition is reporting that a child action failed, not that it failed
itself. Check the actions inside the True and False branches.

---

## `#VALUE!` in Submission < 10 Days or JKR Duration

**Cause:** a date written as text. `setValue(new Date())` in an Office Script
produces a string like `Tue Aug 18 2026 00:10:00 GMT+0000`, and subtracting text
gives `#VALUE!`.

**Fix:** convert to an Excel serial number (epoch 30 Dec 1899), as
`excelSerial()` does.

**Quick visual check:** dates written correctly are right-aligned in Excel.
Left-aligned means text.

---

## Dates are one day behind

The Power Automate runtime is UTC; Sarawak is UTC+8. A submission after 16:00
local time is dated to the previous day unless the offset is applied. See
`CONFIG.tzOffsetHours`.

---

## Values land in the wrong columns after editing the sheet

Older versions used hardcoded column indices, so inserting a column shifted
every write by one with no error raised. The current version resolves columns
by header name at runtime.

If a header is renamed, the script now raises:

```
Column "Approval Date" not found in table "WayleaveTracker".
Either restore the header or update the CONFIG block in this script.
```

Fix by restoring the header or updating `CONFIG`.

---

## New rows appear hundreds of rows down

Both tables contain pre-formatted blank padding rows, and `addRow` appends below
all of them. The scripts fill the first blank row instead.

---

## A flow parameter is suddenly null after editing the Form

Power Automate binds dynamic content to question **IDs**, not titles.

| Change                          | Result |
|---------------------------------|--------|
| Rename a question's text        | Safe |
| Reorder questions               | Safe |
| Add a question                  | Safe |
| Delete a question               | Breaks the mapped parameter |
| Delete **and recreate** it      | Breaks silently — new ID, identical appearance |

Always edit questions in place rather than deleting and re-adding.

---

## Notification Status is stuck on Pending

`UpsertWayleaveRow` runs before the email is sent, so it correctly writes
`Pending`. `MarkNotified` must run afterwards to update it.

If the status is stuck only when sending fails, the failure-path instance of
`MarkNotified` is missing its **Configure run after** settings — it needs
*has failed* and *has timed out* ticked.

---

## Emails silently do not arrive

If the Outlook connector is authenticated with a corporate account, sending to
external addresses may be restricted by tenant policy. The run can still show a
green tick. Check the raw outputs of the send action rather than trusting the
status icon.

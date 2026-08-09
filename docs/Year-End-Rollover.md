# Year-end rollover — the ten-minute procedure

The city's fiscal year ends September 30. This procedure closes the year
into its own archive file and rolls the master workbook forward. Do it in
the first days of October, after the last September trips are reconciled.

**Total time: about ten minutes.** The same checklist is on the workbook's
Instructions sheet.

## 1. Finish the year (± 2 min)

Open the **Ledger**. Every row that had travel should read **Yes** in
`Reconciled?`. Chase anything `Open`: missing receipts (queue 4), missing
actuals, or a packet never sent to Finance. A trip that will genuinely
close after October 1 can roll into the new year — note it and carry it
forward in step 4.

## 2. Archive (± 1 min)

Dashboard → **Archive Year**. The macro saves a complete copy to the
archive folder as `TPD_Training_Master_FY26.xlsm`. That file **is** the
permanent record for the year — auditable, self-contained, never edited
again.

## 3. Advance the settings (± 2 min)

On **Settings**:

| Setting | Change |
|---|---|
| Fiscal year label | `FY26` → `FY27` |
| Fiscal year start | `10/01/2025` → `10/01/2026` |
| Tracking # prefix | `T26-` → `T27-` |
| Training budget | the newly adopted figure |

Check the per diem rates and mileage rate against current policy while
you're there — October is when rates change.

## 4. Clear the records (± 3 min)

On **Requests**, select the *typed values* of all data rows — the input
columns only (A–V, X–Z, AC, AE–AL, AN–AW, BF–BI) — and press Delete.
**Do not delete rows, and do not touch the formula columns (W, AA, AB,
AD, AJ, AM, AR, AT) or the gray helper columns (AX–BE).**

A trip straddling the year (approved in September, travels in November):
re-enter it with a new-year tracking number, or re-import its original
form from the Processed folder. Note the old number in Coordinator Notes.

## 5. Verify and save (± 2 min)

- Dashboard tiles: Spent $0, Encumbered $0 (or exactly your carried-over
  trips), Remaining = the new budget.
- YTD Report months now start with October of the new year.
- Save. Done.

If anything looks wrong, close **without saving** — the archive from
step 2 and the last backup mean nothing is ever lost.

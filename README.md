# TPD Travel & Training Management System

Every year, the Tyler Police Department sends officers to dozens of training
courses and conferences — and every trip generates the same paperwork: a
request form, chain-of-command approvals, a city travel form for Finance,
registration, hotel booking, per diem calculations, and a post-trip
reconciliation of every dollar spent. This project replaces the old scatter of
per-trip spreadsheets and hand-typed ledgers with **one accountable pipeline —
from an officer's request to a closed-out, reconciled training account** —
built entirely in Excel so it runs on tools the department already owns and
can be maintained by whoever holds the Training Coordinator job next.

## The two pieces

| Who | File | What they do |
|---|---|---|
| Officers | `workbooks/TPD_Travel_Request_Form.xlsm` | Fill in the white boxes, click **Submit to Training Unit**. That's the whole process from their side. |
| Training Coordinator | `workbooks/TPD_Training_Master.xlsm` | Everything else — see below. |

**The master workbook**, one record per trip:

- **One-click import** of emailed request forms, each assigned a permanent
  tracking number (`T26-001`, …). Forms are read by *field name*, so the
  officer form's layout can evolve without breaking intake.
- **Every document generated from that single record**: the internal
  approval **Application**, the city's **2-90 travel forms** (pre-trip
  estimate and post-trip reconciliation), and the attendee's
  **Notification** sheet — all calculated to the penny under city travel
  policy, including which meals are reimbursable based on departure and
  return times.
- **Buttons draft the emails** (supervisor approval, officer notification,
  post-trip settlement) — always **opening for review, never sending
  automatically**.
- **A dashboard of what needs doing today**: awaiting approval, needs
  registration, needs hotel, receipts overdue, and which packets must reach
  Finance by the **Tuesday-noon** deadline to get a check cut that week —
  with those deadlines pushed to the Outlook calendar (new items only,
  never duplicated).
- **The ledger and year-to-date report write themselves** from the same
  records, so "where do we stand on the training budget?" is always current
  and always correct.

## Design principles

1. **Formulas, not code, do the math.** Per diem, deadlines, totals, ledger,
   YTD — all spreadsheet formulas: stable, transparent, hard to break. VBA
   only does what formulas can't: email drafts, printing, calendar pushes,
   file imports.
2. **Every setting a future coordinator might change lives on the Settings
   sheet** — email addresses, per diem rates and meal cutoffs, mileage rate,
   deadlines, budget, folder paths, email wording. The system survives a
   change in personnel without anyone touching a line of code.
3. **Three sheets stay untouched**: **Application**, **Pre 2-90** and
   **Post 2-90**. They are protected, fill themselves from the selected
   record, and are never typed on.
4. **Each fiscal year archives cleanly** into its own file; a documented
   ten-minute procedure rolls the system into the next year
   (`docs/Year-End-Rollover.md`).

## Repository layout

```
workbooks/   TPD_Training_Master.xlsx       finished workbook (formulas, no macros yet)
             TPD_Travel_Request_Form.xlsx   finished officer form (no macros yet)
src/vba/     master/*.bas                   macro source for the master workbook
             request_form/*.bas             macro source for the officer form
tools/       Install-VBA.ps1                one-time build: .xlsx + VBA -> .xlsm
build/       build_*.py                     scripts that generate the .xlsx files
docs/        Setup-and-VBA-Install.md       getting from repo to running system
             Coordinator-Guide.md           the Training Coordinator's manual
             Officer-Quick-Start.md         hand this to officers
             Year-End-Rollover.md           the 10-minute fiscal-year rollover
```

## Getting started

1. On a Windows PC with Excel, clone/download this repo and run
   `tools\Install-VBA.ps1` (see `docs/Setup-and-VBA-Install.md`). It imports
   the VBA and produces the two `.xlsm` files.
2. Open the master workbook's **Settings** sheet and set every **yellow**
   cell (real email addresses, current per diem and mileage rates, this
   year's budget, folder paths).
3. Delete the four `SAMPLE —` demo records from the Requests sheet
   (values only — leave formula columns alone).
4. Set the Training Unit's address on the request form's hidden
   **Form Settings** sheet, then distribute the form to officers.

Git note: the VBA lives as readable text in `src/vba/`, so changes to the
automation are reviewable diffs; the workbooks are rebuilt from
`build/*.py` (Python + openpyxl) if the layout ever needs to change.

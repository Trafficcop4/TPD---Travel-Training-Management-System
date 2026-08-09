# Setup — from repository to running system

Excel workbooks in version control can't carry compiled macros, so the repo
stores the workbooks (`workbooks/*.xlsx`) and the macro source
(`src/vba/**/*.bas`) separately. A one-time step on any Windows PC with
Excel joins them into the finished `.xlsm` files.

## 1. Prerequisites

- Windows PC with Microsoft Excel (2016 or later) and Outlook — the same
  machine profile the Training Coordinator uses is ideal.
- This repository, cloned or downloaded as a ZIP and extracted.
- One Excel setting, enabled **once** (you can turn it back off afterwards):

  > File → Options → Trust Center → Trust Center Settings → Macro Settings →
  > ☑ **Trust access to the VBA project object model**

## 2. Run the installer

Open PowerShell in the repo's `tools` folder and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-VBA.ps1
```

It imports every `.bas` module, adds the buttons (the Dashboard strip on the
master; the Submit button on the officer form), and writes:

- `workbooks\TPD_Training_Master.xlsm`
- `workbooks\TPD_Travel_Request_Form.xlsm`

Re-running is safe — it always starts from the `.xlsx` sources and replaces
the `.xlsm` outputs.

### Manual fallback (no PowerShell allowed)

1. Open the `.xlsx` in Excel, press **Alt+F11**, then File → Import File…
   and import each `.bas` from the matching `src/vba/` folder.
2. Save As → **Excel Macro-Enabled Workbook (*.xlsm)**.
3. Add buttons: Developer → Insert → Button (Form Control), draw it, and
   assign the macro. The master's Dashboard names the seven macros; the
   officer form needs one button for `SubmitToTrainingUnit`.

## 3. First-run configuration (the master workbook)

Open `TPD_Training_Master.xlsm` and work down the **Settings** sheet.
Every **yellow** cell is a placeholder that must be set before go-live:

- Training Unit mailbox, coordinator name/phone, Finance email.
- Per diem meal rates and the incidental rate — set to **current City of
  Tyler travel policy** (the shipped numbers are the standard GSA split,
  marked as placeholders).
- Mileage rate — current IRS/city rate.
- This fiscal year's adopted training budget.
- The four folder paths (request inbox, processed, packets, archive).

Then delete the four `SAMPLE —` demonstration records: on **Requests**,
select the *typed values* in rows 2–5 (columns A–V, X–Z, AC, AE–AL, AN–AW,
BF–BI) and press Delete. **Never delete whole rows or the formula columns**
(W, AA, AB, AD, AJ, AM, AR, AT and the gray helper columns AX–BE).

## 4. First-run configuration (the officer form)

The form's **Form Settings** sheet is hidden (right-click a sheet tab →
Unhide). Set the Training Unit email there, re-hide it, and distribute the
`.xlsm` to officers (shared drive or email template — their choice of copy
is fine; the form travels with each request).

## 5. Macro security at the department

The finished `.xlsm` files trigger Excel's macro warning on first open.
Either have IT add the Training Unit's folder as a **Trusted Location**
(File → Options → Trust Center → Trusted Locations), or right-click each
`.xlsm` → Properties → **Unblock** once. Officers get the same one-time
prompt on the request form.

## What the macros are allowed to do

- Emails are **always drafts** — `.Display`, never `.Send`.
- The calendar push only creates items for records whose `Cal Pushed?`
  column isn't `Y`, then marks them — re-running never duplicates.
- Nothing deletes data: the archive macro saves a *copy*; row clearing at
  rollover is a documented manual step, on purpose.

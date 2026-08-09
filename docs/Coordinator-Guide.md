# The Training Coordinator's Guide

This is the manual for `TPD_Training_Master.xlsm`. The one-page version
lives on the workbook's **Instructions** sheet; this document is the full
story, written so the *next* coordinator can take over cold.

## The one idea that runs everything

**One row on the Requests sheet = one trip.** Every form, email, deadline,
ledger line and report number is calculated *from that row*. You never type
the same fact twice, and you never type on a generated form. Pick a record
in the Dashboard's yellow **ACTIVE** box and the Application, both 2-90s,
and the Notification sheet fill themselves.

## The life of a request

1. **It arrives.** An officer clicks Submit on their request form; the
   attachment lands in the request inbox folder (Settings → Folders). Save
   attachments there, then click **Import Requests**. Each form is read by
   field name, appended to Requests with the next tracking number
   (`T26-014`), status `Awaiting Approval`, and the file is moved to the
   Processed folder renamed to its tracking number. The original form is
   the audit copy — never delete it.

2. **Approval.** With the record active, click **Approval Email** — an
   Outlook draft to the supervisor opens with the packet PDF (Application +
   Pre 2-90 + Notification) attached. Review, send. When the approval comes
   back: Status → `Approved`, date in the Approved column. (Denied? Status
   → `Denied` — the record stays for the audit trail and drops out of every
   queue and total.)

3. **Booking.** Register the officer and book the hotel; mark
   `Registered? = Y` and `Hotel Booked? = Y`. The record leaves queues 2
   and 3 the moment you do.

4. **Money to Finance, before the trip.** Queue 5 shows every packet due
   at Finance within a week — the date shown is the **Tuesday**, and the
   real deadline is **noon** that day; that's what gets a check cut the
   same week. Click **Print Packet** for the PDF, send it, then enter the
   date in `Sent to Finance`.

5. **Tell the traveler.** **Officer Notification** drafts the email with
   the Notification sheet attached: what's prepaid, their per diem, their
   receipts-due date. **Push to Outlook** puts the Finance Tuesday-noon
   deadline and the receipts-due date on your calendar — only for records
   not already pushed (`Cal Pushed?` flips to `Y`).

6. **After the trip.** Enter actuals in the Actual columns as receipts come
   in, plus anything the city paid directly and any advance. The Balance
   Due column shows who owes whom; **Settlement Email** drafts the note to
   the traveler. Receipts date → `Receipts Rec'd` (queue 4 stops chasing),
   final packet (now including the Post 2-90) to Finance, date it, Status →
   `Closed`. The Ledger's Reconciled? column reads `Yes` and the trip's
   dollars are final in every report.

## The Dashboard queues, precisely

| Queue | A record appears when… |
|---|---|
| 1 Awaiting approval | Status = `Awaiting Approval` |
| 2 Needs registration | Status = `Approved` and Registered? ≠ Y |
| 3 Needs hotel | Status = `Approved`, Nights > 0, Hotel Booked? ≠ Y |
| 4 Receipts overdue | past (return + receipt days) with no Receipts Rec'd date, and not Closed/Cancelled/Denied |
| 5 Finance packet due | Finance Tuesday is within 7 days (or past) and no Sent-to-Finance date |

Queues list the first eight records each; the header count shows the true
total. For the full picture, filter the Requests sheet itself — every
column has a filter arrow.

## Per diem, exactly as calculated

Rates and cutoffs live on Settings — change them there and every open
record recalculates. A trip earns, per city policy:

- **Departure day:** breakfast if leaving at/before the breakfast cutoff
  (default 6:00 AM), lunch if at/before noon, dinner if at/before 6:00 PM.
- **Return day:** breakfast if back at/after 8:00 AM, lunch at/after
  1:00 PM, dinner at/after 7:00 PM.
- **Full days between:** all three meals.
- **Incidentals:** the daily rate × every travel day, first through last.
- **Same-day trips:** a meal counts only if the departure *and* return
  cutoffs for that meal are both met.

The estimate appears in `Per Diem Est` the moment departure/return dates
*and times* exist — no times, no number, on purpose: make officers give
you times.

## Money definitions (so the reports always reconcile)

- **Spent** = sum of Total Act (every record with actuals entered).
- **Encumbered** = Total Est of records that are alive (not Denied/
  Cancelled/Closed) and have no actuals yet.
- **Remaining** = Budget − Spent − Encumbered.
- **Balance Due Traveler** = Total Act − Paid Direct by City − Advance.
  Positive: city owes the traveler. Negative: the traveler owes the city.

## The three protected sheets

**Application**, **Pre 2-90**, and **Post 2-90** are protected (no
password) and must stay untouched — they fill from the active record. The
2-90s here are functional recreations of the Finance form; if Finance
issues a revised 2-90, the account-code boxes (yellow) still take typed
values, and a layout change should be done deliberately: unprotect, adjust,
re-protect, and note the change in git.

## Entering historical or paper records

Type straight into the next empty Requests row: tracking number by hand
(keep the year's sequence), status, and whatever facts you have. Formula
columns light up as their inputs appear. Records migrated from the old
per-trip spreadsheets were entered exactly this way — and correcting the
old ledger's math errors required nothing but typing the real receipts into
the Actual columns.

## Housekeeping

- **More than 250 requests?** Select the last row's formula + helper cells
  (W, AA, AB, AD, AJ, AM, AR, AT, AX–BE) and fill down. Extend the Ledger
  the same way.
- **Backups:** the workbook is one file — the department's normal drive
  backup covers it. The year-end archive is the permanent record.
- **A formula cell got typed over?** Copy the same cell from any other row
  and paste it back. Every formula column is identical down its column.

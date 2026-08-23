# BPOC Academy Management Workbook (V6) — Coordinator Guide

One workbook runs one academy class, end to end: schedule, roster, grades,
spelling, writing, attendance/makeup, skills, PT, incidents, counseling,
medical, flags, awards, agency emails, TCOLE audit readiness, and every
printable (sign-ins, evals, spelling tests, writing handouts, transcripts).

**Color language:** white boxes with a blue border are yours to type (what
you type shows in blue); gray cells calculate; yellow cells are pending data
(the PT points rubric). Gray `sys*` tabs are the locked engine (password
`TPDAcademy`) — never edit them. The one exception is **sysAwards**, which is
deliberately left editable in its override (E) and notes (G) columns.
**Layout rule:** column A is empty and rows 1–4 are reserved for your own
headers on every sheet. Tables start at row 5, data at row 6.

## Setting up a new academy

1. **Settings** — class label, start/end dates, weights (40/30/10/20),
   caps, retest window, flag thresholds. The **PT rubric block** stays
   yellow until the approved points chart arrives — paste it in and set
   *PT Final Min Points*; the Final-Exam gate activates automatically.
2. **Lists / Agencies / Instructors** — confirm dropdowns and contacts.
   Instructors need name, PID, type, certs; a non-TCOLE instructor is not
   audit-ready until the **SME letter** from the Training Sgt is On File.
3. **ChapterMaster** — TCOLE minimums and planned hours are seeded. The
   per-chapter **training file** columns (SME lesson plan, instructor bio,
   sign-in sheets, assessment, grade sheet, eval) and the **Special TCOLE
   Requirement** column are your audit record; `File Complete?` turns Yes
   when everything required is On File.
4. **Control** — holidays compute from the start date; add any extra
   closure dates. The class-day calendar (Day #, Date, Week #) drives
   retest deadlines, writing dates and sign-in sheets.
5. **Schedule** — one row per time block (date, start, end,
   chapter/activity, instructor). Delivered hours reconcile to ChapterMaster
   live; the audit sheet checks totals against **exactly 736**. Excess is
   never reported inside the BPOC — it goes to the reporting course the
   **Addendum** sheet's "Report Under" column names: #2040 Arrest and
   Control, #2046 Professional Police Driving, #2055 Firearms, and #101
   Addendum to BPOC for everything else.
6. **ExamPlan / WritingMaster / SpellingMaster** — pick this academy's
   exams; writing assignments carry their prompts and **compute assigned/due
   dates by matching each topic's chapter to the schedule** (override
   columns always win); the 12 spelling word lists are stored and printable.
7. **Cadets** — PID, names, agency, status. PID never changes; capacity 50.

## Daily operation

- **ExamScores** — one row per attempt. A failed first attempt shows the
  **Retest Due By** date (5 class days, policy 300.5) and goes red when
  overdue. A passed retest records at the 70 cap automatically; a failed
  retest flags a dismissal review.
- **Missed an exam?** Use the **Absence** column (last column):
  - **Unexcused** — key **Raw Score 0** as well. The zero *is* the record,
    and it starts the ordinary 5-class-day retest clock, exactly like any
    other failing first attempt. Row Check refuses the row until the 0 is
    keyed. A **second** unexcused missed exam raises
    `REMOVAL TRIGGER` on the WatchList and on the Audit sheet.
  - **Excused** — leave Raw Score **blank**. There is no zero and no clock;
    the first attempt is simply taken later. Retest Status reads
    *EXCUSED - 1st attempt pending* so the exam the cadet still owes cannot
    disappear, and **graduation is blocked** until it is taken. When they
    sit it, key the score on that same row — the block clears itself.
- **Spelling** — scores per test; class stats compute; an average under 75
  flags INTERVENTION (policy 300.4.B) — document your response on
  **Counseling**.
- **Attendance / Makeup** — exception log in minutes (classroom) and
  sessions (PT). Makeup is minute-for-minute; `Cl Owed` must reach zero
  before graduation. Print the day's **SignIn** sheet for the paper record
  (TCOLE requires original sign-ins **with PID**).
- **Skills / PT / Medical / Incidents / Counseling / DismissalLog** — log
  as events happen; the engine rolls everything up.
- **Writing** — type **X** when received (lowercase x auto-capitalizes;
  blank = not done); red means past the computed due date.

## Watching the class

- **Dashboard** — KPIs, watch list, class-average trend charts.
- **WatchList** — every flagged cadet with the reasons spelled out; each
  threshold is a Settings value.
- **CadetProfile** — pick a cadet, get the whole picture on one printable
  page (grades, projections, attendance, flags, latest incidents/counseling).
- **GradChecklist / sysChecks** — the graduation gate: 70 in each category
  *and* every category actually recorded (major, minor, spelling and the
  final exam), under caps, makeup complete, skills qualified, writing
  current, no open chain-of-command incident review, final PT **passed**,
  all certificate copies on file, no open dismissal review. A final PT that
  was never assessed reads "Not taken" and a final PT rubric that has not
  been entered on Settings reads "Pending" — both block, neither is a pass.
- **sysAwards** — Valedictorian, Physical Fitness, Top Gun, Grit computed
  from the data; your override cell always wins (Grit especially is
  decision support, not a verdict).

## Agency emails

Set *Current Exam #* / *Current Spelling #* on Settings, review
**EmailPreview**, then the button builds per-agency Outlook drafts (home
agency gets everyone). Drafts open for review — nothing is ever auto-sent.
V6 adds: class average beside each score, retest-cap notes, attendance and
writing status, spelling-intervention notices, and a **"since your last
report"** digest of negative incidents and counseling pulled from EmailLog
timestamps.

## Printing

**PrintCenter** lists every printable with its button (after VBA install):
sign-in sheet, eval stack (one critique per active cadet, auto-filled from
the schedule), any spelling test or key, weekly writing handouts with
prompts, profiles, transcripts, ranking, grad checklist, audit packet,
schedule.

## End of academy

1. Enter final PT points and the Final Exam scores.
2. GradChecklist: every gate column Yes and **Blocking Issues reading
   "Eligible"** (the engine writes that word for a passing cadet — the cell is
   never blank); print
   **Transcript** for every cadet (button).
3. Confirm the **Audit** sheet. Everything should read OK / Yes / N/A with
   ONE expected exception: the delivered-hours check reads
   **"OVER - report excess as #101"** for any academy that delivers more than
   736 hours, which is what this academy does and what this guide tells you to
   do — report the BPOC at exactly 736 and the excess under the course numbers
   the **Addendum** sheet names. "OVER" is the correct, expected state, not a
   defect; "SHORT" is the one that must never be left standing.
4. Record awards (override cells on sysAwards).
5. Save-As the finished file (e.g. `BPOC-2026-01 FINAL.xlsm`), then run
   **New Academy Reset** on a fresh copy for the next class.

## Rebuilding from source

```
python3 build/bpoc/build_bpoc.py        # writes workbooks/BPOC_Academy_Management_V6.xlsx
python3 build/bpoc/verify_build.py      # structural checks
powershell -File tools/Install-BPOC-VBA.ps1   # on Windows+Excel: builds the .xlsm
```

## Daily operations (v6.1)

- **DailyLog** is the digital daily report: one row per training day —
  present count, AM/PM notes, Issues flag, and whether the class leader's
  signed report was scanned — **the scan is the legal original**, confirmed
  by TPD records management for training documents (and department documents
  generally). Keep that determination on file; the Audit sheet's "Scanned
  documents established as legal originals" item is where you record it. The
  counters (incidents, early departures, memos received) compute themselves.
- **SignIn** is now the one-page Daily Report & Roster: AM roll call +
  sign-in grid + PM changes + signatures. Print on demand.
- **Missed time clears by event**: every counting Attendance event stays
  OPEN until that same cadet's Makeup rows linked to its EventID cover the
  balance — then it shows CLEARED with the date. The Dashboard lists every
  OPEN event.
- **What counts against the cap is the `Excused?` column, not the Reason.**
  The engine's `Counts?` formula reads `Excused?` (and the one documented
  exception: a `PT Modified` event whose documentation is Received). It does
  not look at the Reason column at all — Reason is reporting detail. So an
  agency recall counts against the cap and must be made up **only if you log
  it as `Counted`**. If a recall is genuinely excused, mark it `Excused` and
  it neither counts nor needs makeup. Earlier wording here said agency
  recalls always count as if the engine enforced it; it does not, and the
  choice is deliberately yours on each row.
- **Makeup "Row Check"** is the last column on the Makeup sheet. Credit only
  applies when it reads OK. It refuses another cadet's event (WRONG CADET),
  an EventID that no longer exists (NO SUCH EVENT), a Type the caps do not
  credit (TYPE NOT CREDITED — the dropdown now offers only Classroom and PT,
  because those are the only two the caps credit), minutes booked
  against a PT event, which is counted in sessions (UNIT MISMATCH), a row
  with **no makeup Date** (a dateless credit used to stamp a fabricated
  "CLEARED 12/30" on the attendance ledger) and a zero or **negative**
  credit (which used to pass as OK and *increase* the cadet's owed time).
  That keeps the CLEARED banner and the `Cl Owed` / `PT Net` balances in
  step.
- **Memos**: assigned for deficiencies, linked to the triggering Incident/
  Attendance/Counseling ID (or stand-alone), due N class days later
  (Settings), overdue memos flag the cadet on the WatchList.
- **Selective agency reporting**: nothing reaches an agency email digest
  unless YOU mark it — Incidents "Report to Agency?", Counseling "Agency
  Notified?", Memos "Report to Agency?". Blank = academy teaching moment.

# BPOC Academy Management Workbook (V6) — Coordinator Guide

One workbook runs one academy class, end to end: schedule, roster, grades,
spelling, writing, attendance/makeup, skills, PT, incidents, counseling,
medical, flags, awards, agency emails, TCOLE audit readiness, and every
printable (sign-ins, evals, spelling tests, writing handouts, transcripts).

**Color language:** blue cells are yours to type; gray cells calculate;
yellow cells are pending data (the PT points rubric). Gray `sys*` tabs are
the locked engine (password `TPDAcademy`) — never edit them.
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
   live; the audit sheet checks totals against **exactly 736** (excess must
   be reported as Addendum course #101, never inside the BPOC).
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
- **GradChecklist / sysChecks** — the graduation gate: 70 in each category,
  under caps, makeup complete, skills qualified, writing current, final PT
  passed, no open dismissal review.
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
2. GradChecklist all-Yes; print **Transcript** for every cadet (button).
3. Confirm the **Audit** sheet is all OK / all Yes.
4. Record awards (override cells on sysAwards).
5. Save-As the finished file (e.g. `BPOC-2026-01 FINAL.xlsm`), then run
   **New Academy Reset** on a fresh copy for the next class.

## Rebuilding from source

```
python3 build/bpoc/build_bpoc.py        # writes workbooks/BPOC_Academy_Management_V6.xlsx
python3 build/bpoc/verify_build.py      # structural checks
powershell -File tools/Install-BPOC-VBA.ps1   # on Windows+Excel: builds the .xlsm
```

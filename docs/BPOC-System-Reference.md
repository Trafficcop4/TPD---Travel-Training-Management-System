# BPOC Academy Management Workbook — System Reference

How the workbook is built, what every naming convention means, how the
pieces feed each other, and what comes out of the printer.

This is the "how it works" document. For the day-to-day routine see
`BPOC-Coordinator-Guide.md`; for audit mapping see `BPOC-Audit-Prep.md`.

- **Scope:** one workbook = one academy class. Save-As at the end, then
  **New Academy Reset** on the copy.
- **Size:** 60 sheets, 293 named ranges, ~63,000 formulas (7,250 of them rewritten by the prefixer for Excel compatibility).
- **Requires:** Microsoft 365 Excel (uses `LET`, `XLOOKUP`, `FILTER`,
  `TAKE`, `HSTACK`). Macros come from `tools/Install-BPOC-VBA.ps1`.

---

## 1. The layout contract

Every sheet obeys the same geometry. Nothing in the workbook works if
this is broken, because every formula is written against it.

| Where | What lives there |
|---|---|
| Column A | **Always empty** — a left margin so nothing touches the edge |
| Rows 1–4 | **Yours.** Row 1 col B holds the `◄ Dashboard` link; row 4 holds the gray usage note. Put your own headers/notes here |
| Row 5 | **Table headers.** Never move or rename these — formulas and the VBA read them |
| Row 6+ | **Data.** First data row is always 6 |

**Cadet grids are row-mirrored.** On every per-cadet sheet, a cadet
occupies *the same row number* across all of them:

```
Cadets!row 12  ==  Spelling!row 12  ==  Writing!row 12  ==  PT!row 12
               ==  Certifications!row 12  ==  sysGrades!row 12  ==  sysChecks!row 12
```

Rows 6–55 = 50 cadet slots. This is why engine sheets can say
`Cadets!$B6` instead of doing a lookup — and why **you must never sort
or delete rows** on a cadet grid. Change Status to `Separated` instead;
the row grays out and drops from every count automatically.

**Log capacities** (one row per event, oldest at top):

| Sheet | Rows | Sheet | Rows |
|---|---|---|---|
| ExamScores | 6–1505 | Skills | 6–605 |
| Schedule | 6–905 | Makeup | 6–505 |
| Attendance | 6–805 | EmailLog | 6–505 |
| Incidents / Counseling | 6–405 | Memos | 6–305 |
| Medical | 6–205 | Control class days | 6–175 |
| DailyLog | 6–175 | DismissalLog | 6–105 |

---

## 2. Sheet taxonomy — read the tab colors

| Tab color | Class | Meaning | Protected? |
|---|---|---|---|
| 🟡 Gold | **CONFIG** | Set up once per academy, then leave alone | No (except InputGuide) |
| 🔵 Blue | **INPUT** | You type here daily | No |
| 🟢 Green | **OUTPUT** | Results and printables — read, don't type | Yes (picker cells unlocked) |
| ⚫ Gray | **SYSTEM** | `sys*` calculation engine | Yes — never edit |

**Cell colors inside a sheet:** white box with border = you type it
(your entries show blue); gray fill = calculated; yellow = data pending
(the PT points rubric).

Protection password is `TPDAcademy`. Protected sheets are still fully
**selectable and scrollable** — only editing is blocked, except the
deliberately unlocked picker cells (`C5` on the printables).

`sysListsHelper` is **veryHidden** on purpose: it carries a 1,600-row
spill that feeds the Memos dropdown. Leave it hidden.

---

## 3. Naming conventions

### 3.1 Sheet names
- **`sys*`** — calculation engine, locked, no input, never printed.
  Everything else is either config, input, or an output/printable.
- **`*Master`** — a reference library that outlives one academy
  (`ExamMaster`, `SkillsMaster`, `SpellingMaster`, `WritingMaster`,
  `ChapterMaster`). Contrast `ExamPlan` = the subset *this* class uses.
- No spaces in any sheet name — formulas and VBA reference them bare.

### 3.2 Named ranges — four prefixes, and they mean different things

**`cfg*` — a single configuration cell (47).**
One value, not a range. Most live on Settings (`cfgPassingScore`,
`cfgWeightMajor`, `cfgRetestClassDays`). The rest are **selector cells**
that drive a printable, and they live on the sheet they steer:

| Name | Lives on | Picks |
|---|---|---|
| `cfgProfileCadet` | CadetProfile!C5 | which cadet's profile |
| `cfgTranscriptCadet` | Transcript!C5 | which cadet's transcript |
| `cfgSignInDate` | SignIn!C5 | which day's roster |
| `cfgEvalChapter` | EvalSheet!C5 | which chapter's critique |
| `cfgPacketChapter` | ChapterPacket!C5 | which chapter's file page |
| `cfgGradeSheetExam` | ExamSheet!C5 | which exam's grade sheet |
| `cfgSpellPrintNum` / `cfgSpellPrintMode` | SpellingPrint!C5 / E5 | test # and Test-vs-Key |
| `cfgHandoutWeek` | WritingHandout!C5 | which week's assignments |
| `cfgPreviewAgency` | EmailPreview!C5 | which agency to preview |
| `cfgPolicyVersion`, `cfgBoardReviewed`, `cfgRulesAligned`, `cfgAlignReviewer` | AdvisoryBoard!D6–D9 | governance alignment |

> **Rule:** if you want to change what a printable shows, you change a
> `cfg*` cell — never the sheet body.

**`lst*` — a dropdown list (26).** All on the Lists sheet, one column
each. `lstCadetStatus`, `lstAttendanceEvent`, `lstReason`,
`lstCounselingType`, `lstMaterialsStatus`… To add a choice, type it at
the bottom of that column on Lists.

**`nr*` — a data column in a table (192).** The workhorse. Format is
`nr` + a short sheet tag + the field:

| Tag | Sheet | Examples |
|---|---|---|
| `nrES_` | ExamScores | `nrES_PID`, `nrES_Code`, `nrES_Rec`, `nrES_Final` |
| `nrAT_` | Attendance | `nrAT_ID`, `nrAT_Min`, `nrAT_Cleared` |
| `nrMK_` | Makeup | `nrMK_Link`, `nrMK_Credit`, `nrMK_Min` |
| `nrSK_` | Skills | `nrSK_Cat`, `nrSK_Res`, `nrSK_CoF` |
| `nrIN_` / `nrCO_` / `nrME_` / `nrMD_` | Incidents / Counseling / Memos / Medical | `nrIN_Report`, `nrME_Due` |
| `nrSCH_` | Schedule | `nrSCH_Date`, `nrSCH_Hrs`, `nrSCH_Act`, `nrSCH_ChNum` |
| `nrCH*` | ChapterMaster | `nrCHname`, `nrCHmin`, `nrCHdeliv`, `nrCHfileOK` |
| `nrCD*` | Control class days | `nrCDnum`, `nrCDdate`, `nrCDweek` |
| `nrSUB*` | ChapterMaster sub-class block | `nrSUBname`, `nrSUBparent` |
| `nrGR*` | sysGrades | `nrGRcurrent`, `nrGRrank`, `nrGRfinalExam` |
| `nrATT*` | sysAttendance | `nrATTclOwed`, `nrATTmakeupOK` |
| `nrFL*` / `nrCK*` / `nrAUD*` / `nrAW*` | sysFlags / sysChecks / sysAudit / sysAwards | `nrFLreasons`, `nrCKgradElig` |
| `nrSpell*` / `nrWR*` / `nrPT_*` / `nrCERT*` | Spelling / Writing / PT / Certifications | `nrSpellAvg`, `nrWRoverdue`, `nrPT_FinalPass` |

Note the two easily-confused ones: **`nrGRfinal`** is the weighted
*overall* grade; **`nrGRfinalExam`** is the *final exam* average. A
printout labelled "Final Exam" must use the latter. (Reading the wrong
one was a real bug caught in stress testing.)

**`rng*` — legacy compatibility ranges (28).** Same idea as `nr*`, kept
under the original V5.4 spellings so the roster/exam/agency lookups did
not have to be rewritten: `rngCadetPIDs`, `rngCadetNames`,
`rngAgencyIDs`, `rngEPcode`, `rngWMdue`, `rngSM_cat`. Treat as `nr*`.

### 3.3 Record IDs
Auto-generated, never typed. The letter tells you the source sheet:

| Prefix | Sheet | Example |
|---|---|---|
| `A###` | Attendance | `A012` |
| `M###` | Makeup | `M004` |
| `I###` | Incidents | `I007` |
| `C###` | Counseling | `C003` |
| `ME###` | Memos | `ME009` |
| `MD###` | Medical | `MD002` |
| `R###` | DismissalLog | `R001` |

These IDs are the glue: a Makeup row points at an `A###`, and a Memo
points at an `I###`, `A###`, or `C###`.

### 3.4 Column conventions inside a sheet
- **`Row Check`** — always the **last** column of a log. Reads `OK` or a
  short reason. Anything other than OK means that row is being ignored
  or is dangerous. Check it when a number looks wrong.
- **`… ?`** (question mark) — a Yes/No judgement (`Is PT?`, `Counts?`,
  `Cleared?`, `All Certs?`).
- **`Computed …` / `… Override` / plain name** — three-column pattern:
  the workbook computes, you may override, the plain column is the one
  everything downstream uses (see Memos `Due (computed)` / `Due
  Override` / `Due`).

---

## 4. How the pieces interact

### 4.1 The dependency spine

```
Settings (start date, weights, caps, thresholds)
      │
      ├─► Control ── holidays ─► CLASS-DAY CALENDAR (Day #, Date, Week #)
      │                                │
      │                                ├─► ExamScores "Retest Due By"  (+5 class days)
      │                                ├─► Memos "Due"                  (+3 class days)
      │                                ├─► WritingMaster assigned/due   (chapter week)
      │                                ├─► SignIn / DailyLog day numbers
      │                                └─► Dashboard "Today"
      │
      └─► Schedule (one row per time block)
                 │
                 ├─► ChapterMaster Delivered Hrs, First/Last Taught  ─► Audit, Addendum
                 ├─► Instructors "On Schedule?" / "Chapters Taught"  ─► Audit
                 ├─► EvalSheet + ChapterPacket headers
                 └─► WritingMaster (when a chapter is first taught)

Cadets (PID = the key)
      │
      ├─► every log (name → PID lookup)
      └─► row-mirrored to Spelling, Writing, PT, Certifications, StateExam, sys*

Logs ──► sys* engine ──► Dashboard / WatchList / GradChecklist / printables / email
```

**The class-day calendar is the backbone.** Control counts weekdays from
the start date, skipping computed holidays and any Extra Closure Dates
you add. Every deadline in the workbook is measured in *class days*, not
calendar days — that is why a retest due date skips a holiday weekend
automatically.

### 4.2 Cadet identity
You type a **name** in a log; the sheet resolves it to a **PID**; every
engine calculation is `SUMIFS`/`COUNTIFS` keyed on PID. Consequences:

- A misspelled name resolves to nothing and that row silently leaves the
  cadet's record. The **`Row Check`** column and the audit check *"Log
  rows with unrecognized cadet name"* both catch this.
- **Duplicate PIDs merge two cadets.** The PID cell turns red and the
  audit sheet counts it. Fix immediately.

### 4.3 Grades
`ExamScores` holds **one row per attempt**. Per exam, the engine picks
the *final* attempt and computes a **Recorded** score:

- Attempt 1 passed → recorded as-is.
- Attempt 2 passed → **recorded at 70** (`cfgRetakeRecordedCap`).
- Attempt 2 failed → dismissal review flag.
- A failed attempt 1 with no attempt 2 yet → `Retest Due By` = 5 class
  days out; `Retest Status` goes **OVERDUE** in red when it passes.

`sysGrades` then averages by category and weights them
40 / 30 / 10 / 20. **Passing requires ≥70 in each category** — the
weighted total alone is not enough. A category with no recorded scores
is *not* counted as passed; `sysChecks` "Exams Recorded" enforces that
so an unfinished cadet can never print as eligible.

### 4.4 Attendance ↔ Makeup (the per-event ledger)
This is the most-linked mechanism in the workbook.

1. Log the miss on **Attendance**: minutes (classroom) *or* sessions
   (PT), reason, whether it counts. You get `A012`.
2. Log the makeup on **Makeup** and pick **`A012`** in *Linked Event*.
3. Attendance `A012` shows **Made-Up**, **Balance**, and flips from
   **OPEN** to **CLEARED 09/14** when the balance reaches zero.

Guardrails on the Makeup row (`Row Check`): it refuses another cadet's
event (**WRONG CADET**), a non-existent ID (**NO SUCH EVENT**), a type
the caps don't credit (**TYPE NOT CREDITED**), and minutes booked
against a PT event (**UNIT MISMATCH**). Credit is also capped at the
event's own size, so over-crediting cannot mask an owed balance.

`sysAttendance` totals it against the caps: classroom minutes vs
`cfgClassroomCapMinutes` (5% of scheduled minutes) and PT sessions vs
`cfgPTCapSessions` (5). **Makeup Complete?** must be Yes to graduate.

### 4.5 Chapters, sub-classes, and hours
You schedule under **your** class names; TCOLE counts **its** chapters.
The sub-class block at the bottom of ChapterMaster maps them:

```
"Traffic Code" (50) ┐
"Crash Investigation" (12) ├─► parent chapter 22 (TCOLE min 74)
"TIM" (12) ┘
seven "Criminal Investigations - …" sub-classes ─► parent chapter 32
```

Schedule rows resolve to a parent `Ch #`; ChapterMaster sums delivered
hours by chapter number. The BPOC is reported at **exactly 736 hours** —
the **Addendum** sheet itemises every hour over a chapter minimum and
names the course to report it under (#2040 Arrest & Control, #2046
Driving, #2055 Firearms, #101 Addendum for everything else).

### 4.6 Instructors
`Instructors` scans the Schedule's instructor text and reports **On
Schedule?** and **Chapters Taught** per person — including co-teachers
buried in a multi-name cell. **Audit Ready?** requires a documented
qualification (TCOLE cert, or SME letter on file) **plus a bio**.
Anyone teaching without documentation turns red.

`InstructorBanks` is the per-topic pool: **Bank 1–10** = everyone
certified for that topic (survives New Academy Reset); **Teach 1–10** =
who you picked for *this* class. The Schedule's instructor dropdown for
a topic offers exactly that topic's picks; pick a second name and it
appends, pick an existing one and it is removed.

### 4.7 Flags vs gates — two different things
- **`sysFlags` (warnings)** — 14 threshold tests, every threshold a
  `cfg*` cell on Settings. Produces a **Flag Count** and a plain-English
  **Reasons** string. Surfaces on **WatchList** and the Dashboard.
  Flags never block anything.
- **`sysChecks` (the gate)** — the graduation decision: academics,
  classroom, PT sessions, skills assessed *and* qualified, incidents,
  writing, makeup complete, final PT passed, no open dismissal review.
  Produces **GraduationElig** and a **Blocking Issues** sentence.

Certifications are deliberately a **flag, not a gate** — a missing
certificate copy nags on the Dashboard rather than blocking the cadet.

### 4.8 Agency email
`EmailPreview` shows exactly what will go out for the exam set in
`cfgCurrentExamNum`. The macro drafts one Outlook message per agency
(the home agency gets everyone), **always as drafts, never sent**, and
logs the run to `EmailLog`.

Content per cadet: score, class average, retest note, spelling (omitted
when the spelling number trails the exam number), attendance tier and
writing status. Plus a **"since your last report"** digest — and this is
opt-in: only rows you marked **Report to Agency? = Yes** on Incidents,
Counseling, or Memos appear. Everything else stays an academy matter.
The cutoff comes from that agency's last EmailLog date.

### 4.9 The audit engine
`sysAudit` runs **27 live checks** — hours vs 736, chapter files
complete, special TCOLE requirements, instructor documentation, overdue
retests, makeup owed, certifications, governance alignment, plus the
data-integrity checks (`Row Check` failures, duplicate PIDs, unrecognized
names). The **Audit** sheet displays them, adds the program-requirement
checklist you tick, and the per-cadet enrollment-document grid.

---

## 5. Printed documentation

Everything prints from **PrintCenter** (buttons appear after the VBA
install). The pattern is always: **set a `cfg*` picker cell → press the
button**. Print areas are pre-set, so `Ctrl+P` on the sheet works too.

### 5.1 The daily record

| Printable | Set first | What it is | Where it goes |
|---|---|---|---|
| **Daily Report & Roster** (SignIn) | `cfgSignInDate` | One page: schedule for the day, AM roll-call block, sign-in grid (name + PID + AM/PM signature), PM changes block, class-leader and coordinator signature lines | Class leader fills and signs; **scan it** into that day's folder — the scan is the legal original, confirmed by TPD records management (recorded on the Audit sheet item 'Scanned documents established as legal originals') |
| **…week** | — | Next 5 class days, each pre-filled with its own date, day # and schedule | Hand the leader a week at a time |
| **…academy book** | — | Every class day in the academy, one run | Bind at the start of the academy |

You then log only the **exceptions** into Attendance/Incidents/Memos and
one 30-second **DailyLog** row, and tick *Leader Report Scanned?*.

### 5.2 The TCOLE training file (one folder per chapter)

| Printable | Set first | Purpose |
|---|---|---|
| **Chapter Packet** | `cfgPacketChapter` | The auditor's "show me chapter 23" page: hours vs minimum, every schedule block, every instructor who taught it with documentation status, the file checklist, special TCOLE requirement, linked exam stats. **Use it as the folder cover sheet** |
| **Exam Grade Sheet** | `cfgGradeSheetExam` | The IRG-required grade sheet per assessment: every cadet's raw and recorded score, pass/fail, retest note, class stats, proctor signature |
| **Course Critique** (EvalSheet) | `cfgEvalChapter` | The 11-question evaluation, course/instructor/dates auto-filled. The stack button prints one per active cadet |

Recommended folder, which the Chapter Packet doubles as the index for:

```
BPOC-2026-01/Ch23 - Intoxicated Driver (SFST)/
  00 ChapterPacket.pdf      ← cover: what this folder must contain
  01 Lesson Plan.pdf        ← SME-developed (a PowerPoint alone does not count)
  02 Instructor Bios+certs.pdf
  03 Sign-ins/              ← scans = the legal originals
  04 Assessment.pdf         ← LMS test printed to PDF
  05 Grade Sheet.pdf        ← ExamSheet print
  06 Evaluations/           ← scanned critiques
```

### 5.3 Cadet-facing and class-facing

| Printable | Set first | Use |
|---|---|---|
| **Spelling Test / Key** | `cfgSpellPrintNum`, `cfgSpellPrintMode` | Any of the 12 tests, blank or answer key |
| **Writing Handout** | `cfgHandoutWeek` | That week's assignments with full prompts and computed due dates |
| **Cadet Profile** | `cfgProfileCadet` | One page: grades, projections, attendance, open time, open memos, flags, recent incidents/counseling. **Your prep sheet for an agency call or a counseling session** |
| **Transcript** | `cfgTranscriptCadet` | End-of-academy record per cadet — categories, exam list, attendance, skills, PT, certifications, state exam, awards, signature line. Button can run the whole class |
| **Class Ranking** | — | Live ranking, unranked cadets excluded |
| **Graduation Checklist** | — | Every gate column per cadet plus Blocking Issues — the pre-ceremony sign-off |
| **Audit Packet** | — | The 27 live checks, program checklist, per-cadet enrollment documents |
| **Addendum** | — | Excess hours per class with the course number to report each under |
| **Schedule** | — | Full schedule listing, landscape, repeating headers |

---

## 6. Rhythm

**Daily** — print/hand out the roster; log exceptions, incidents, memos;
one DailyLog row; glance at the Dashboard.

**Per exam** — enter scores (one row per attempt), print the **Exam
Grade Sheet** into the chapter folder, set `cfgCurrentExamNum`, review
**EmailPreview**, mark anything that should reach agencies, draft the
emails.

**Per chapter** — collect critiques, tick the ChapterMaster file
columns, print the **Chapter Packet** as the folder cover.

**Weekly** — WatchList; clear OPEN missed time and outstanding memos.

**End of academy** — final PT and final exam; GradChecklist all-Yes with
Blocking Issues empty; print transcripts; confirm the Audit sheet; record
awards; Save-As `BPOC-2026-01 FINAL.xlsm`.

**New academy** — on a *copy*: **New Academy Reset** clears cadet and
daily data but keeps masters, instructor banks, and the board-meeting
list. It then runs **Academy Startup Review**, which asks for the latest
advisory-board meeting, whether rules changed, and the policy version —
and the Audit sheet stays red until you confirm the workbook was aligned.

---

## 7. When something looks wrong

1. **Check the `Row Check` column** on that log — a rejected row is
   being ignored on purpose, and it tells you why.
2. **Check the Audit sheet** — 27 checks including data integrity.
3. **Red PID or red cadet name** = duplicate. Two cadets are being
   merged. Fix first, ask questions later.
4. **A number that will not move** — the row's cadet name probably does
   not match the roster, so the row belongs to nobody.
5. **`#SPILL!`** — something was typed into a calculated block on a
   green sheet. Clear it; those areas are formula-owned.

---

## 8. Rebuilding from source

The workbook is **generated**, not hand-edited. To change it, change the
builder and rebuild:

```bash
python3 build/bpoc/build_bpoc.py     # writes workbooks/BPOC_Academy_Management_V6.xlsx
python3 build/bpoc/verify_build.py   # structural assertions - must pass
python3 build/bpoc/seed_bpoc7.py <v5.4.xlsm> <calendar.xlsx>   # optional: load real data
powershell -File tools/Install-BPOC-VBA.ps1                    # Windows+Excel: build the .xlsm
```

| File | Builds |
|---|---|
| `build/bpoc/sheets_config.py` | Settings, Lists, Agencies, Instructors, InstructorBanks, masters, Control, Schedule |
| `build/bpoc/sheets_inputs.py` | Every blue input sheet |
| `build/bpoc/sheets_engine.py` | Every `sys*` sheet |
| `build/bpoc/sheets_outputs.py` | Dashboard, printables, email sheets |
| `build/bpoc/data_*.py` | Seed data: chapters/hours, spelling words, writing prompts, lists |
| `build/bpoc/postprocess.py` | Stores modern functions with the `_xlfn.`/`_xlpm.` prefixes Excel needs |
| `build/bpoc/verify_build.py` | 141 structural checks; run after every change |
| `build/bpoc/lo_compat_for_test.py` | Test-only: lets LibreOffice evaluate the class-day chain |
| `src/vba/bpoc/*.bas` | Email, print, reset, buttons |

**Two rules if you edit the builder:**
1. Formulas reference columns by **literal letter**. Insert a column and
   you must update every cross-sheet reference, the VBA, and
   `verify_build.py`. This has been the single largest source of bugs.
2. Write plain function names in Python (`XLOOKUP`, `LET`) —
   `postprocess.py` adds the storage prefixes. Do **not** prefix
   `WORKDAY.INTL`/`NETWORKDAYS.INTL`; they are native and Excel shows
   `#NAME?` if prefixed.

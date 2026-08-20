"""Config & master sheets: StartHere, Settings, Lists, Agencies, Instructors,
ChapterMaster, ExamMaster, ExamPlan, SkillsMaster, SpellingMaster,
WritingMaster, Control (holidays + class-day calendar), Schedule.
"""
from openpyxl.utils import get_column_letter

from xlb import (
    HDR_ROW, DATA_ROW, CADET_LAST, ROWS_SCHEDULE, ROWS_CLASSDAYS,
    F_HDR, FILL_HDR, F_CALC, FILL_CALC, F_LABEL, F_BODY, F_SMALL, F_INPUT,
    FILL_INPUT, FILL_BAND, FILL_YELLOW, FILL_PAGE, FILL_STEEL, F_SECTION,
    A_LEFT, A_LEFT_WRAP, A_CENTER, BOX, DATE, TIME,
    header_row, fill_rows, dv_list, name_range, sheet_note, cf_yes_no,
    cf_formula, FILL_WARNBG, FILL_OKBG, FILL_AMBER, col_widths, section_bar,
    label, define, protect, unlock_range, page_setup_landscape,
    page_setup_portrait,
)
import data_chapters as DC
import data_lists as DL
import data_spelling as DS
import data_writing as DW


# --------------------------------------------------------------------------
def build_starthere(wb):
    ws = wb.create_sheet("StartHere")
    ws.sheet_view.showGridLines = False
    steps = [
        ("1. Settings", "Set academy class label, start/end dates, weights, "
                        "caps and flag thresholds. Blue cells are yours."),
        ("2. Lists / Agencies / Instructors", "Confirm dropdown lists, agency "
                        "contacts and the instructor roster (PID, certs, SME letters)."),
        ("3. ChapterMaster & Schedule", "Confirm chapter hours, then enter the "
                        "class schedule (date, times, chapter, instructor). "
                        "Delivered hours reconcile automatically."),
        ("4. ExamPlan / WritingMaster / SpellingMaster", "Pick this academy's "
                        "exams; writing due dates compute from the schedule."),
        ("5. Cadets", "Enter the roster (PID, names, agency). Everything keys "
                        "off PID."),
        ("6. Daily operation", "Log exam scores, spelling, attendance "
                        "exceptions, makeup, skills, writing, incidents, "
                        "counseling, PT and medical as they happen."),
        ("7. Watch the Dashboard / WatchList", "Flags, retest deadlines and "
                        "blocking issues surface automatically."),
        ("8. Print Center", "Sign-in sheets, evals, spelling tests, writing "
                        "handouts, transcripts and rosters — one place."),
        ("9. Agency emails", "EmailPreview + buttons draft per-agency Outlook "
                        "emails for review. Never auto-sent."),
        ("10. New academy", "Save-As, then use the New Academy reset button. "
                        "See the Coordinator Guide in docs/."),
    ]
    header_row(ws, ["Step", "What to do"])
    r = DATA_ROW
    for step, what in steps:
        ws.cell(row=r, column=2, value=step).font = F_LABEL
        c = ws.cell(row=r, column=3, value=what)
        c.font = F_BODY
        c.alignment = A_LEFT_WRAP
        ws.row_dimensions[r].height = 30
        r += 1
    col_widths(ws, {"A": 3, "B": 34, "C": 90})
    sheet_note(ws, "Workbook map: blue = you type it, gray = calculated, "
                   "sys* sheets are the locked engine.")
    return ws


# --------------------------------------------------------------------------
SETTINGS = [
    # (label, value, note, name, fmt)
    ("Academy Class", "BPOC-2026-01", "Label shown on reports", "cfgAcademyClass", None),
    ("Start Date", "2026-05-11", "First day of academy", "cfgStartDate", DATE),
    ("End Date", "2026-11-06", "Scheduled graduation date", "cfgEndDate", DATE),
    ("Total Scheduled Minutes", 145800, "Academy length in minutes", "cfgTotalScheduledMinutes", "#,##0"),
    ("Passing Score", 70, "Minimum passing score for exams", "cfgPassingScore", None),
    ("Pass Threshold Score", 70, "Category-average threshold", "cfgThresholdScore", None),
    ("Retake Recorded Cap", 70, "Recorded score after a passed retest", "cfgRetakeRecordedCap", None),
    ("Threshold After Exam #", 4, "Category avg enforced after this many exams", "cfgThresholdAfterExam", None),
    ("Retest Within (class days)", 5, "Policy 300.5 retest window", "cfgRetestClassDays", None),
    ("Weight: Major", 0.4, "Major exams weight", "cfgWeightMajor", "0%"),
    ("Weight: Minor", 0.3, "Minor exams weight", "cfgWeightMinor", "0%"),
    ("Weight: Spelling", 0.1, "Spelling weight", "cfgWeightSpelling", "0%"),
    ("Weight: Final", 0.2, "Final exam weight", "cfgWeightFinal", "0%"),
    ("Spelling Intervention Avg", 75, "Below this spelling avg = early intervention (300.4.B)", "cfgSpellInterventionAvg", None),
    ("Classroom Cap %", 0.05, "Attendance cap as % of total minutes", "cfgClassroomCapPct", "0.0%"),
    ("PT Cap (sessions)", 5, "Max PT sessions missed", "cfgPTCapSessions", None),
    ("Attendance Advisory %", 0.6, "Tier 1 warning at this % of cap", "cfgAttendanceAdvisoryPct", "0%"),
    ("Attendance Critical %", 0.8, "Tier 2 warning at this % of cap", "cfgAttendanceCriticalPct", "0%"),
    ("Home Agency (gets all cadets)", "TPD", "Agency code whose email includes every cadet", "cfgHomeAgency", None),
    ("Current Exam # (for emails)", 1, "Sequence # of the exam to email", "cfgCurrentExamNum", None),
    ("Current Spelling # (for emails)", 1, "Highest spelling test # administered", "cfgCurrentSpellingNum", None),
    ("Writing Due Time", "1700", "Due time shown on handouts", "cfgWritingDueTime", None),
    ("PT Final Min Points", 0, "Minimum points to pass final PT (set when rubric arrives)", "cfgPTFinalMinPoints", None),
    # flag engine thresholds
    ("Flag: consecutive exam fails", 2, "Consecutive graded-exam scores below passing", "cfgFlagConsecFails", None),
    ("Flag: grade drop (points)", 10, "Drop in recorded score vs cadet's prior exam", "cfgFlagGradeDrop", None),
    ("Flag: category margin", 5, "Category avg within this many points of 70", "cfgFlagCategoryMargin", None),
    ("Flag: attendance % of cap", 0.6, "Classroom or PT usage at/over this % of cap", "cfgFlagAttendancePct", "0%"),
    ("Flag: open negative incidents", 2, "Open negative incidents at/above this count", "cfgFlagOpenIncidents", None),
    ("Flag: overdue writing", 1, "Overdue writing assignments at/above this count", "cfgFlagOverdueWriting", None),
]

CALC_SETTINGS = [
    ("Classroom Cap (minutes)", "ROUND(cfgTotalScheduledMinutes*cfgClassroomCapPct,0)",
     "CALCULATED", "cfgClassroomCapMinutes", "#,##0"),
    ("Weights total (must = 100%)", "cfgWeightMajor+cfgWeightMinor+cfgWeightSpelling+cfgWeightFinal",
     "CALCULATED", "cfgWeightsTotal", "0%"),
]


def build_settings(wb):
    ws = wb.create_sheet("Settings")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Setting", "Value", "Notes", "Internal name",
                    "Detected from data", "Check"])
    r = DATA_ROW
    for lab, val, note, nm, fmt in SETTINGS:
        ws.cell(row=r, column=2, value=lab).font = F_LABEL
        c = ws.cell(row=r, column=3, value=val)
        c.fill = FILL_INPUT
        c.font = F_INPUT
        c.border = BOX
        if fmt:
            c.number_format = fmt
        if nm in ("cfgStartDate", "cfgEndDate"):
            from datetime import datetime
            c.value = datetime.strptime(val, "%Y-%m-%d")
            c.number_format = DATE
        ws.cell(row=r, column=4, value=note).font = F_SMALL
        ws.cell(row=r, column=5, value=nm).font = F_SMALL
        define(wb, nm, "Settings", f"$C${r}")
        r += 1
    for lab, fx, note, nm, fmt in CALC_SETTINGS:
        ws.cell(row=r, column=2, value=lab).font = F_LABEL
        c = ws.cell(row=r, column=3, value="=" + fx)
        c.fill = FILL_CALC
        c.font = F_CALC
        c.border = BOX
        if fmt:
            c.number_format = fmt
        ws.cell(row=r, column=4, value=note).font = F_SMALL
        ws.cell(row=r, column=5, value=nm).font = F_SMALL
        define(wb, nm, "Settings", f"$C${r}")
        r += 1
    # detection helpers for the two "Current #" settings
    for i, (lab, _, _, nm, _) in enumerate(SETTINGS):
        row = DATA_ROW + i
        if nm == "cfgCurrentExamNum":
            ws.cell(row=row, column=6, value=(
                "=IFERROR(MAX(FILTER(nrES_Seq,(nrES_Rec<>\"\")*(nrES_Seq<>\"\"))),0)"
            )).font = F_CALC
            ws.cell(row=row, column=7, value=(
                f'=IF($F${row}=0,"No exam scores posted yet",'
                f'IF($C${row}>$F${row},"CHECK - set to Test "&$C${row}&" but data only has Test "&$F${row},'
                f'IF($F${row}>$C${row},"CHECK - data shows Test "&$F${row},"OK")))'
            )).font = F_SMALL
        if nm == "cfgCurrentSpellingNum":
            ws.cell(row=row, column=6, value=(
                "=MAX(SUMPRODUCT(MAX((nrSpellTaken>0)*COLUMN(nrSpellTaken)))-COLUMN(Spelling!$D$1)+1,0)"
            )).font = F_CALC
            ws.cell(row=row, column=7, value=(
                f'=IF($C${row}<cfgCurrentExamNum,'
                f'"CHECK - spelling will be OMITTED this run",'
                f'IF($F${row}>$C${row},"CHECK - data shows Spelling "&$F${row},"OK"))'
            )).font = F_SMALL
    weights_row = r
    ws.cell(row=r + 1, column=2, value="Weights check:").font = F_LABEL
    ws.cell(row=r + 1, column=3, value=(
        '=IF(ROUND(cfgWeightsTotal,4)=1,"OK","FIX: weights must total 100%")'
    )).font = F_CALC
    r += 3

    # ---- PT final-assessment points rubric (values arrive later) ----
    section_bar(ws, r, 2, 9, "Final PT Assessment — Points Rubric "
                             "(enter the approved chart when received; "
                             "engine uses it as soon as points are present)")
    r += 1
    rub_hdr = r
    header_row(ws, ["Event", "Standard (baseline)", "Level 1 pts",
                    "Level 2 pts", "Level 3 pts", "Level 4 pts", "Level 5 pts",
                    "Level thresholds (notes)"], row=r)
    r += 1
    rub_first = r
    for ev, std, unit, hib in DL.PT_EVENTS:
        ws.cell(row=r, column=2, value=ev).font = F_LABEL
        ws.cell(row=r, column=3, value=std).font = F_SMALL
        for c in range(4, 9):
            cc = ws.cell(row=r, column=c)
            cc.fill = FILL_YELLOW
            cc.font = F_INPUT
            cc.border = BOX
        ws.cell(row=r, column=9).fill = FILL_YELLOW
        ws.cell(row=r, column=9).font = F_INPUT
        ws.cell(row=r, column=9).border = BOX
        r += 1
    define(wb, "nrPTRubricEvents", "Settings", f"$B${rub_first}:$B${r-1}")
    define(wb, "nrPTRubricPoints", "Settings", f"$D${rub_first}:$H${r-1}")
    col_widths(ws, {"A": 3, "B": 34, "C": 22, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12, "I": 30})
    sheet_note(ws, "Blue = edit per academy. Yellow = pending data (PT rubric "
                   "from the other coordinator). Gray = calculated. Internal "
                   "names must not change.")
    return ws


# --------------------------------------------------------------------------
def build_lists(wb):
    ws = wb.create_sheet("Lists")
    ws.sheet_view.showGridLines = False
    keys = list(DL.LISTS.keys())
    header_row(ws, keys)
    for i, k in enumerate(keys):
        colL = get_column_letter(2 + i)
        vals = DL.LISTS[k]
        for j, v in enumerate(vals):
            c = ws.cell(row=DATA_ROW + j, column=2 + i, value=v)
            c.fill = FILL_INPUT
            c.font = F_INPUT
            c.border = BOX
        nm = ("lst" + "".join(w.capitalize() for w in
              k.replace("/", " ").replace("?", "").replace("#", "Num")
              .replace("-", " ").split()))
        define(wb, nm, "Lists", f"${colL}${DATA_ROW}:${colL}${DATA_ROW+len(vals)-1}")
        ws.column_dimensions[colL].width = max(16, len(k) + 2)
    ws.column_dimensions["A"].width = 3
    sheet_note(ws, "These feed every dropdown. Add items at the bottom of a "
                   "column, then extend the named range (Formulas > Name "
                   "Manager) if you exceed the seeded rows.")
    return ws


# --------------------------------------------------------------------------
def build_agencies(wb):
    ws = wb.create_sheet("Agencies")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["AgencyID", "AgencyName", "Primary Contact", "Email",
                    "Phone", "Mailing Address", "ActiveYN", "Last Email Sent"])
    last = DATA_ROW + 44
    for i, (aid, aname, contact, email, phone, addr, act) in enumerate(DL.AGENCIES):
        r = DATA_ROW + i
        for col, v in ((2, aid), (3, aname), (4, contact), (5, email),
                       (6, phone), (7, addr), (8, act)):
            c = ws.cell(row=r, column=col, value=v)
            c.fill = FILL_INPUT
            c.font = F_INPUT
            c.border = BOX
    fill_rows(ws, DATA_ROW, last, {
        "B": (None, "in"), "C": (None, "in"), "D": (None, "in"),
        "E": (None, "in"), "F": (None, "in"), "G": (None, "in"),
        "H": (None, "in"),
        "I": ('IF($B{r}="","",IFERROR(MAXIFS(nrELdate,nrELagency,$B{r}),""))', "fx"),
    })
    for r in range(DATA_ROW, last + 1):
        ws[f"I{r}"].number_format = DATE
    dv_list(ws, "=lstYesNo", [f"H{DATA_ROW}:H{last}"])
    define(wb, "rngAgencyIDs", "Agencies", f"$B${DATA_ROW}:$B${last}")
    define(wb, "rngAgencyNames", "Agencies", f"$C${DATA_ROW}:$C${last}")
    define(wb, "rngAgencyEmails", "Agencies", f"$E${DATA_ROW}:$E${last}")
    define(wb, "rngAgencyActive", "Agencies", f"$H${DATA_ROW}:$H${last}")
    define(wb, "nrAgencyLastEmail", "Agencies", f"$I${DATA_ROW}:$I${last}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24, "D": 20, "E": 46, "F": 14,
                    "G": 32, "H": 9, "I": 14})
    sheet_note(ws, "AgencyID is the short code used on Cadets. 'Last Email "
                   "Sent' is calculated from EmailLog and drives the "
                   "since-last-email discipline digest.")
    return ws


# --------------------------------------------------------------------------
def build_instructors(wb):
    ws = wb.create_sheet("Instructors")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Instructor Name", "Agency / Unit", "TCOLE PID",
                    "Instructor Type", "Certificates (list)",
                    "Cert Expiration", "SME Letter", "Bio On File?",
                    "Notes", "Audit Ready?", "On Schedule?",
                    "Chapters Taught"])
    last = DATA_ROW + 89
    for i, nm in enumerate(DL.INSTRUCTORS):
        ws.cell(row=DATA_ROW + i, column=2, value=nm)
    for j, nm in enumerate(DL.GUEST_ENTITIES):
        r = DATA_ROW + len(DL.INSTRUCTORS) + j
        ws.cell(row=r, column=2, value=nm)
        ws.cell(row=r, column=5, value="Guest/Outside")
    fill_rows(ws, DATA_ROW, last, {
        "B": (None, "in"), "C": (None, "in"), "D": (None, "in"),
        "E": (None, "in"), "F": (None, "in"), "G": (None, "in"),
        "H": (None, "in"), "I": (None, "in"), "J": (None, "in"),
        # IRG: every instructor/co-teacher teaching LOs needs a documented
        # qualification (TCOLE cert or written SME approval) AND a bio.
        "K": ('IF($B{r}="","",IF($E{r}="Guest/Outside","N/A",'
              'IF(AND(OR($E{r}="TCOLE Instructor",'
              'AND(LEFT($E{r},3)="SME",$H{r}="On File")),'
              'OR($I{r}="On File",$I{r}="N/A")),"Yes","No")))', "fx"),
        "L": ('IF($B{r}="","",IF(SUMPRODUCT(--ISNUMBER(SEARCH($B{r},'
              'nrSCH_Instr)))>0,"Yes",""))', "fx"),
        "M": ('IF($B{r}="","",IFERROR(TEXTJOIN(", ",TRUE,SORT(UNIQUE('
              'FILTER(nrSCH_ChNum,ISNUMBER(SEARCH($B{r},nrSCH_Instr))*'
              '(nrSCH_ChNum<>""))))),""))', "fx"),
    })
    for r in range(DATA_ROW, last + 1):
        ws[f"G{r}"].number_format = DATE
    dv_list(ws, "=lstInstructorType", [f"E{DATA_ROW}:E{last}"])
    dv_list(ws, "=lstMaterialsStatus",
            [f"H{DATA_ROW}:H{last}", f"I{DATA_ROW}:I{last}"])
    cf_yes_no(ws, f"K{DATA_ROW}:K{last}")
    cf_formula(ws, f"L{DATA_ROW}:L{last}",
               f'AND($L{DATA_ROW}="Yes",$K{DATA_ROW}="No")', FILL_WARNBG)
    define(wb, "nrInstrNames", "Instructors", f"$B${DATA_ROW}:$B${last}")
    define(wb, "nrInstrReady", "Instructors", f"$K${DATA_ROW}:$K${last}")
    define(wb, "nrInstrOnSched", "Instructors", f"$L${DATA_ROW}:$L${last}")
    col_widths(ws, {"A": 3, "B": 24, "C": 18, "D": 12, "E": 20, "F": 30,
                    "G": 13, "H": 12, "I": 12, "J": 26, "K": 11, "L": 12,
                    "M": 30})
    sheet_note(ws, "IRG: EVERY instructor or co-teacher who teaches a "
                   "learning objective needs a bio + documented "
                   "qualification (TCOLE instructor cert, or SME letter "
                   "approved in writing). 'On Schedule?' scans every "
                   "schedule block's instructor text — co-teachers in "
                   "multi-name entries are matched individually. Red = "
                   "teaching without documentation. Guest/Outside entries "
                   "(proctors, venues) are exempt.")
    return ws


# --------------------------------------------------------------------------
BANK_SLOTS = 10      # certified instructors per topic
SEL_SLOTS = 8        # picked to teach this academy


def build_instructorbanks(wb):
    """Per-topic instructor banks: the certified pool persists across
    academies; the 'Teaching this academy' picks are cleared on reset and
    drive the Schedule's instructor dropdown for that topic's rows."""
    from openpyxl.utils import get_column_letter as gcl
    ws = wb.create_sheet("InstructorBanks")
    ws.sheet_view.showGridLines = False
    hdrs = (["Topic / Class"]
            + [f"Bank {i+1}" for i in range(BANK_SLOTS)]
            + [f"Teach {i+1}" for i in range(SEL_SLOTS)]
            + ["# Selected"])
    header_row(ws, hdrs)
    first = DATA_ROW
    topics = ([name for _m, _c, name, _mi, _tp in DC.CHAPTERS]
              + [name for name, _p, _t in DC.SUBTOPICS]
              + [a for a in DC.ACTIVITIES if not a.startswith("Test ")
                 and a != "Lunch"])
    r = first
    for t in topics:
        ws.cell(row=r, column=2, value=t)
        r += 1
    last = r + 9        # spare rows
    bank_first_c, bank_last_c = 3, 2 + BANK_SLOTS               # C..L
    sel_first_c, sel_last_c = 3 + BANK_SLOTS, 2 + BANK_SLOTS + SEL_SLOTS  # M..T
    cnt_c = sel_last_c + 1                                       # U
    cols = {"B": (None, "in")}
    for c in range(bank_first_c, sel_last_c + 1):
        cols[gcl(c)] = (None, "in")
    cols[gcl(cnt_c)] = (
        'IF($B{r}="","",COUNTA(%s{r}:%s{r}))'
        % (gcl(sel_first_c), gcl(sel_last_c)), "fx")
    fill_rows(ws, first, last, cols)
    dv_list(ws, "=nrInstrNames",
            [f"{gcl(bank_first_c)}{first}:{gcl(bank_last_c)}{last}"])
    # academy picks must come from THAT ROW's certified bank
    dv_list(ws, f"=INDEX(${gcl(bank_first_c)}${first}:${gcl(bank_last_c)}"
                f"${last},MATCH($B{first},$B${first}:$B${last},0),0)",
            [f"{gcl(sel_first_c)}{first}:{gcl(sel_last_c)}{last}"],
            enforce=False)
    define(wb, "nrBankTopics", "InstructorBanks", f"$B${first}:$B${last}")
    define(wb, "nrBankGrid", "InstructorBanks",
           f"${gcl(bank_first_c)}${first}:${gcl(bank_last_c)}${last}")
    define(wb, "nrBankSel", "InstructorBanks",
           f"${gcl(sel_first_c)}${first}:${gcl(sel_last_c)}${last}")
    col_widths(ws, {"A": 3, "B": 46})
    for c in range(bank_first_c, cnt_c):
        ws.column_dimensions[gcl(c)].width = 16
    ws.column_dimensions[gcl(cnt_c)].width = 10
    sheet_note(ws, "Bank = everyone certified/approved to teach the topic "
                   "(kept between academies). Teach = who you picked for "
                   "THIS academy — the Schedule's instructor dropdown for a "
                   "topic offers exactly these names. Topics without picks "
                   "fall back to the full roster. New Academy Reset clears "
                   "Teach columns only.")
    return ws


# --------------------------------------------------------------------------
def build_chaptermaster(wb):
    ws = wb.create_sheet("ChapterMaster")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Module", "Ch #", "Chapter Name", "TCOLE Min Hrs",
                    "Planned Hrs", "Delivered Hrs", "vs TCOLE Min",
                    "First Taught", "Last Taught", "Default Instructor",
                    # per-chapter TRAINING FILE items (IRG mandatory list)
                    "Lesson Plan (SME)", "Instructor Bio", "Sign-In Sheets",
                    "Assessment", "Grade Sheet", "Evals Collected",
                    "Handouts/PPT (opt.)",
                    "Special TCOLE Requirement", "Special Req Met?",
                    "File Complete?"])
    first = DATA_ROW
    r = first
    for mod, ch, name, minh, tpdh in DC.CHAPTERS:
        ws.cell(row=r, column=2, value=mod)
        ws.cell(row=r, column=3, value=ch)
        ws.cell(row=r, column=4, value=name)
        ws.cell(row=r, column=5, value=minh)
        ws.cell(row=r, column=6, value=tpdh)
        if ch in DC.SPECIAL_REQS:
            ws.cell(row=r, column=19, value=DC.SPECIAL_REQS[ch])
        r += 1
    last = r - 1
    fill_rows(ws, first, last, {
        "B": (None, "in"), "C": (None, "in"), "D": (None, "in"),
        "E": (None, "in"), "F": (None, "in"),
        # rollup by chapter NUMBER: schedule rows resolve sub-classes
        # (Traffic Code, Crash, TIM, Crim-Inv subs) to their parent chapter
        "G": ('IF($D{r}="","",ROUND(SUMIFS(nrSCH_Hrs,nrSCH_ChNum,$C{r}),2))', "fx"),
        "H": ('IF($D{r}="","",IF($G{r}=0,"",$G{r}-$E{r}))', "fx"),
        "I": ('IF($D{r}="","",IFERROR(IF(MINIFS(nrSCH_Date,nrSCH_ChNum,$C{r})=0,'
              '"",MINIFS(nrSCH_Date,nrSCH_ChNum,$C{r})),""))', "fx"),
        "J": ('IF($D{r}="","",IFERROR(IF(MAXIFS(nrSCH_Date,nrSCH_ChNum,$C{r})=0,'
              '"",MAXIFS(nrSCH_Date,nrSCH_ChNum,$C{r})),""))', "fx"),
        "K": (None, "in"),
        "L": (None, "in"), "M": (None, "in"), "N": (None, "in"),
        "O": (None, "in"), "P": (None, "in"), "Q": (None, "in"),
        "R": (None, "in"), "S": (None, "in"), "T": (None, "in"),
        "U": ('IF($D{r}="","",IF(AND(COUNTIF(L{r}:P{r},"On File")'
              '+COUNTIF(L{r}:P{r},"N/A")=5,'
              'OR($Q{r}="On File",$Q{r}="N/A"),'
              'OR($S{r}="",$T{r}="Yes",$T{r}="N/A")),"Yes","No"))', "fx"),
    })
    for r2 in range(first, last + 1):
        ws[f"I{r2}"].number_format = DATE
        ws[f"J{r2}"].number_format = DATE
        ws[f"S{r2}"].alignment = A_LEFT_WRAP
    dv_list(ws, "=nrInstrNames", [f"K{first}:K{last}"])
    dv_list(ws, "=lstMaterialsStatus",
            [f"L{first}:L{last}", f"M{first}:M{last}", f"N{first}:N{last}",
             f"O{first}:O{last}", f"P{first}:P{last}", f"Q{first}:Q{last}",
             f"R{first}:R{last}"])
    dv_list(ws, '"Yes,No,N/A"', [f"T{first}:T{last}"])
    cf_formula(ws, f"H{first}:H{last}",
               f'AND($H{first}<>"",$H{first}<0)', FILL_WARNBG)
    cf_yes_no(ws, f"U{first}:U{last}")
    define(wb, "nrCHmod", "ChapterMaster", f"$B${first}:$B${last}")
    define(wb, "nrCHnum", "ChapterMaster", f"$C${first}:$C${last}")
    define(wb, "nrCHname", "ChapterMaster", f"$D${first}:$D${last}")
    define(wb, "nrCHmin", "ChapterMaster", f"$E${first}:$E${last}")
    define(wb, "nrCHplan", "ChapterMaster", f"$F${first}:$F${last}")
    define(wb, "nrCHdeliv", "ChapterMaster", f"$G${first}:$G${last}")
    define(wb, "nrCHfirst", "ChapterMaster", f"$I${first}:$I${last}")
    define(wb, "nrCHinstr", "ChapterMaster", f"$K${first}:$K${last}")
    define(wb, "nrCHlesson", "ChapterMaster", f"$L${first}:$L${last}")
    define(wb, "nrCHbio", "ChapterMaster", f"$M${first}:$M${last}")
    define(wb, "nrCHsignin", "ChapterMaster", f"$N${first}:$N${last}")
    define(wb, "nrCHassess", "ChapterMaster", f"$O${first}:$O${last}")
    define(wb, "nrCHgrade", "ChapterMaster", f"$P${first}:$P${last}")
    define(wb, "nrCHevals", "ChapterMaster", f"$Q${first}:$Q${last}")
    define(wb, "nrCHspecial", "ChapterMaster", f"$S${first}:$S${last}")
    define(wb, "nrCHspecialMet", "ChapterMaster", f"$T${first}:$T${last}")
    define(wb, "nrCHfileOK", "ChapterMaster", f"$U${first}:$U${last}")
    # totals row
    tr = last + 2
    ws.cell(row=tr, column=4, value="Totals:").font = F_LABEL
    ws.cell(row=tr, column=5, value=f"=SUM(E{first}:E{last})").font = F_CALC
    ws.cell(row=tr, column=6, value=f"=SUM(F{first}:F{last})").font = F_CALC
    ws.cell(row=tr, column=7, value=f"=SUM(G{first}:G{last})").font = F_CALC
    ws.cell(row=tr + 1, column=4, value="Required TCOLE hours (exact):").font = F_LABEL
    ws.cell(row=tr + 1, column=5, value=DC.REQUIRED_TCOLE_HOURS).font = F_CALC
    ws.cell(row=tr + 1, column=6, value=(
        "Report the BPOC at exactly 736; excess goes to Addendum course #101"
    )).font = F_SMALL
    define(wb, "nrCHtotalDelivered", "ChapterMaster", f"$G${tr}")

    # ---- TPD sub-classes (scheduled separately, roll up to a chapter) ----
    sr = tr + 3
    section_bar(ws, sr, 2, 9, "TPD sub-classes — schedule by THESE names; "
                              "hours roll up to the parent TCOLE chapter")
    sr += 1
    header_row(ws, ["Sub-class", None, None, "Parent Ch #", "TPD Target Hrs",
                    "Delivered Hrs", "vs Target"], row=sr)
    ws.merge_cells(start_row=sr, start_column=2, end_row=sr, end_column=4)
    sr += 1
    sub_first = sr
    for name, parent, target in DC.SUBTOPICS:
        ws.merge_cells(start_row=sr, start_column=2, end_row=sr, end_column=4)
        c = ws.cell(row=sr, column=2, value=name)
        c.fill = FILL_INPUT
        c.font = F_INPUT
        c.border = BOX
        p = ws.cell(row=sr, column=5, value=parent)
        p.fill = FILL_INPUT
        p.font = F_INPUT
        p.border = BOX
        t = ws.cell(row=sr, column=6, value=target)
        t.fill = FILL_INPUT
        t.font = F_INPUT
        t.border = BOX
        d = ws.cell(row=sr, column=7, value=(
            f'=IF($B{sr}="","",ROUND(SUMIFS(nrSCH_Hrs,nrSCH_Act,$B{sr}),2))'))
        d.font = F_CALC
        d.fill = FILL_CALC
        d.border = BOX
        v = ws.cell(row=sr, column=8, value=(
            f'=IF(OR($B{sr}="",$G{sr}=0),"",ROUND($G{sr}-$F{sr},2))'))
        v.font = F_CALC
        v.border = BOX
        sr += 1
    sub_last = sr + 4          # a few blank rows for future sub-classes
    for rr in range(sr, sub_last + 1):
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        for col in (2, 5, 6):
            cc = ws.cell(row=rr, column=col)
            cc.fill = FILL_INPUT
            cc.font = F_INPUT
            cc.border = BOX
        ws.cell(row=rr, column=7, value=(
            f'=IF($B{rr}="","",ROUND(SUMIFS(nrSCH_Hrs,nrSCH_Act,$B{rr}),2))'
        )).font = F_CALC
        ws.cell(row=rr, column=8, value=(
            f'=IF(OR($B{rr}="",$G{rr}=0),"",ROUND($G{rr}-$F{rr},2))'
        )).font = F_CALC
    dv_list(ws, "=nrCHnum", [f"E{sub_first}:E{sub_last}"])
    cf_formula(ws, f"H{sub_first}:H{sub_last}",
               f'AND($H{sub_first}<>"",$H{sub_first}<0)', FILL_WARNBG)
    define(wb, "nrSUBname", "ChapterMaster", f"$B${sub_first}:$B${sub_last}")
    define(wb, "nrSUBparent", "ChapterMaster", f"$E${sub_first}:$E${sub_last}")
    define(wb, "nrSUBtarget", "ChapterMaster", f"$F${sub_first}:$F${sub_last}")

    col_widths(ws, {"A": 3, "B": 8, "C": 6, "D": 44, "E": 10, "F": 10,
                    "G": 11, "H": 10, "I": 12, "J": 12, "K": 20, "L": 14,
                    "M": 13, "N": 13, "O": 12, "P": 12, "Q": 13, "R": 15,
                    "S": 46, "T": 13, "U": 12})
    sheet_note(ws, "Per the IRG, every chapter's training file needs: SME "
                   "lesson plan (a PowerPoint alone does NOT count), "
                   "instructor bio, original sign-in sheets w/ PID, "
                   "assessment + grade sheet, and course evaluation. "
                   "'File Complete?' turns Yes when all are On File/NA and "
                   "any special TCOLE requirement is met.")
    return ws


# --------------------------------------------------------------------------
def build_exammaster(wb):
    ws = wb.create_sheet("ExamMaster")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["ExamCode", "Exam Name", "Default Type", "Default Passing",
                    "Default Seq"])
    r = DATA_ROW
    for code, name, typ, passing, seq in DL.EXAMS:
        for col, v in ((2, code), (3, name), (4, typ), (5, passing), (6, seq)):
            c = ws.cell(row=r, column=col, value=v)
            c.fill = FILL_INPUT
            c.font = F_INPUT
            c.border = BOX
        r += 1
    last = r + 6  # room to add exams
    fill_rows(ws, r, last, {c: (None, "in") for c in "BCDEF"})
    dv_list(ws, "=lstExamType", [f"D{DATA_ROW}:D{last}"])
    define(wb, "rngEMcode", "ExamMaster", f"$B${DATA_ROW}:$B${last}")
    define(wb, "rngEMname", "ExamMaster", f"$C${DATA_ROW}:$C${last}")
    define(wb, "rngEMtype", "ExamMaster", f"$D${DATA_ROW}:$D${last}")
    define(wb, "rngEMpass", "ExamMaster", f"$E${DATA_ROW}:$E${last}")
    define(wb, "rngEMseq", "ExamMaster", f"$F${DATA_ROW}:$F${last}")
    col_widths(ws, {"A": 3, "B": 11, "C": 52, "D": 13, "E": 14, "F": 12})
    sheet_note(ws, "Library of all possible exams. ExamPlan picks which ones "
                   "THIS academy uses.")
    return ws


def build_examplan(wb):
    ws = wb.create_sheet("ExamPlan")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Use?", "ExamCode", "Exam Name", "Type", "Passing", "Seq",
                    "Linked Ch #", "Exam Date"])
    first, last = DATA_ROW, DATA_ROW + 24
    exam_ch = {"E01": "1", "E02": "7", "E03": "8", "E04": "9", "E05": "10",
               "E06": "15", "E07": "17", "E08": "20", "E09": "22",
               "E10": "23", "E11": "25", "E12": "26", "E13": "28",
               "E14": "29", "E15": "32", "E16": "34", "E17": "40",
               "FIN": ""}
    r = first
    for code, name, typ, passing, seq in DL.EXAMS:
        ws.cell(row=r, column=2, value="Yes")
        ws.cell(row=r, column=3, value=code)
        ws.cell(row=r, column=8, value=exam_ch.get(code, ""))
        r += 1
    fill_rows(ws, first, last, {
        "B": (None, "in"), "C": (None, "in"),
        "D": ('IF($C{r}="","",IFERROR(INDEX(rngEMname,MATCH($C{r},rngEMcode,0)),"?"))', "fx"),
        "E": ('IF($C{r}="","",IFERROR(INDEX(rngEMtype,MATCH($C{r},rngEMcode,0)),"?"))', "fx"),
        "F": ('IF($C{r}="","",IFERROR(INDEX(rngEMpass,MATCH($C{r},rngEMcode,0)),cfgPassingScore))', "fx"),
        "G": ('IF($C{r}="","",IFERROR(INDEX(rngEMseq,MATCH($C{r},rngEMcode,0)),""))', "fx"),
        "H": (None, "in"),
        "I": ('IF(OR($C{r}="",$H{r}=""),"",IFERROR(IF(MAXIFS(nrSCH_Date,'
              'nrSCH_ChNum,$H{r})=0,"",MAXIFS(nrSCH_Date,nrSCH_ChNum,$H{r})),""))', "fx"),
    })
    for r2 in range(first, last + 1):
        ws[f"I{r2}"].number_format = DATE
    dv_list(ws, "=lstYesNo", [f"B{first}:B{last}"])
    dv_list(ws, "=rngEMcode", [f"C{first}:C{last}"])
    dv_list(ws, "=nrCHnum", [f"H{first}:H{last}"])
    define(wb, "rngEPuse", "ExamPlan", f"$B${first}:$B${last}")
    define(wb, "rngEPcode", "ExamPlan", f"$C${first}:$C${last}")
    define(wb, "rngEPname", "ExamPlan", f"$D${first}:$D${last}")
    define(wb, "rngEPtype", "ExamPlan", f"$E${first}:$E${last}")
    define(wb, "rngEPpass", "ExamPlan", f"$F${first}:$F${last}")
    define(wb, "rngEPseq", "ExamPlan", f"$G${first}:$G${last}")
    define(wb, "rngEPch", "ExamPlan", f"$H${first}:$H${last}")
    define(wb, "rngEPdate", "ExamPlan", f"$I${first}:$I${last}")
    col_widths(ws, {"A": 3, "B": 7, "C": 11, "D": 52, "E": 10, "F": 10,
                    "G": 7, "H": 11, "I": 12})
    sheet_note(ws, "Linked Ch # ties each exam to its chapter; 'Exam Date' "
                   "estimates from the last scheduled day of that chapter "
                   "(informational — retest deadlines run off actual score "
                   "dates).")
    return ws


def build_skillsmaster(wb):
    ws = wb.create_sheet("SkillsMaster")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Category", "Max Attempts", "Scoring Mode", "Passing Score"])
    first = DATA_ROW
    r = first
    for cat, mx, mode, passing in DL.SKILLS:
        ws.cell(row=r, column=2, value=cat)
        ws.cell(row=r, column=3, value=mx)
        ws.cell(row=r, column=4, value=mode)
        if passing is not None:
            ws.cell(row=r, column=5, value=passing)
        r += 1
    last = r + 3
    fill_rows(ws, first, last, {c: (None, "in") for c in "BCDE"})
    dv_list(ws, "=lstScoringMode", [f"D{first}:D{last}"])
    define(wb, "rngSM_cat", "SkillsMaster", f"$B${first}:$B${last}")
    define(wb, "rngSM_max", "SkillsMaster", f"$C${first}:$C${last}")
    define(wb, "rngSM_mode", "SkillsMaster", f"$D${first}:$D${last}")
    define(wb, "rngSM_pass", "SkillsMaster", f"$E${first}:$E${last}")
    col_widths(ws, {"A": 3, "B": 18, "C": 14, "D": 14, "E": 14})
    sheet_note(ws, "Attempt limits per policy 300.7/600: exhausting attempts "
                   "= separation. Firearms scored; others pass/fail.")
    return ws


# --------------------------------------------------------------------------
def build_spellingmaster(wb):
    ws = wb.create_sheet("SpellingMaster")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Word #"] + [f"Test {n}" for n in range(1, 13)])
    first = DATA_ROW
    for w in range(25):
        ws.cell(row=first + w, column=2, value=w + 1).font = F_LABEL
        for t in range(1, 13):
            c = ws.cell(row=first + w, column=2 + t,
                        value=DS.TESTS[t][w])
            c.fill = FILL_INPUT
            c.font = F_INPUT
            c.border = BOX
    last = first + 24
    define(wb, "nrSpellWords", "SpellingMaster", f"$C${first}:$N${last}")
    r = last + 2
    ws.cell(row=r, column=2, value="Each word is worth 4 points "
            "(25 words x 4 = 100).").font = F_SMALL
    col_widths(ws, {"A": 3, "B": 8})
    for t in range(1, 13):
        ws.column_dimensions[get_column_letter(2 + t)].width = 16
    sheet_note(ws, "Word lists print from Print Center (test sheet or key). "
                   "Edit words here to change future academies.")
    return ws


# --------------------------------------------------------------------------
def build_writingmaster(wb):
    ws = wb.create_sheet("WritingMaster")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["#", "Assignment Title", "Linked Ch #", "Assign Delay (wks)",
                    "Due Offset (wks)", "Computed Assigned", "Computed Due",
                    "Override Assigned", "Override Due", "Assigned", "Due",
                    "Prompt"])
    first = DATA_ROW
    r = first
    for num, title, ch, delay, off, prompt in DW.ASSIGNMENTS:
        ws.cell(row=r, column=2, value=num)
        ws.cell(row=r, column=3, value=title)
        ws.cell(row=r, column=4, value=ch)
        ws.cell(row=r, column=5, value=delay)
        ws.cell(row=r, column=6, value=off)
        pc = ws.cell(row=r, column=13, value=prompt)
        pc.alignment = A_LEFT_WRAP
        r += 1
    last = r - 1
    # Computed Assigned: first class day of (chapter-first-taught week + delay)
    fill_rows(ws, first, last, {
        "B": (None, "in"), "C": (None, "in"), "D": (None, "in"),
        "E": (None, "in"), "F": (None, "in"),
        "G": ('IF($D{r}="","",LET(f,IFERROR(XLOOKUP($D{r},nrCHnum,nrCHfirst),""),'
              'IF(OR(f="",f=0),"",LET(wkstart,f-WEEKDAY(f,2)+1+$E{r}*7,'
              'IFERROR(MINIFS(nrCDdate,nrCDdate,">="&wkstart),"")))))', "fx"),
        "H": ('IF($G{r}="","",LET(a,$G{r},ws2,a-WEEKDAY(a,2)+1+$F{r}*7,'
              'd,MAXIFS(nrCDdate,nrCDdate,">="&ws2,nrCDdate,"<"&(ws2+7)),'
              'IF(d=0,"",d)))', "fx"),
        "I": (None, "in"), "J": (None, "in"),
        "K": ('IF($B{r}="","",IF($I{r}<>"",$I{r},$G{r}))', "fx"),
        "L": ('IF($B{r}="","",IF($J{r}<>"",$J{r},$H{r}))', "fx"),
        "M": (None, "in"),
    })
    for r2 in range(first, last + 1):
        for cl in "GHIJKL":
            ws[f"{cl}{r2}"].number_format = DATE
        ws[f"M{r2}"].alignment = A_LEFT_WRAP
        ws.row_dimensions[r2].height = 42
    dv_list(ws, "=nrCHnum", [f"D{first}:D{last}"])
    define(wb, "rngWMnum", "WritingMaster", f"$B${first}:$B${last}")
    define(wb, "rngWMtitle", "WritingMaster", f"$C${first}:$C${last}")
    define(wb, "rngWMassigned", "WritingMaster", f"$K${first}:$K${last}")
    define(wb, "rngWMdue", "WritingMaster", f"$L${first}:$L${last}")
    define(wb, "rngWMprompt", "WritingMaster", f"$M${first}:$M${last}")
    nr = last + 2
    ws.cell(row=nr, column=2, value="Requirements:").font = F_LABEL
    c = ws.cell(row=nr, column=3, value=DW.REQUIREMENTS_NOTE)
    c.font = F_SMALL
    c.alignment = A_LEFT_WRAP
    ws.merge_cells(start_row=nr, start_column=3, end_row=nr, end_column=12)
    col_widths(ws, {"A": 3, "B": 5, "C": 38, "D": 10, "E": 12, "F": 12,
                    "G": 14, "H": 14, "I": 14, "J": 14, "K": 12, "L": 12,
                    "M": 80})
    sheet_note(ws, "Dates COMPUTE from the Schedule: assigned = first class "
                   "day of the week the linked chapter is first taught (+ "
                   "delay); due = last class day of the due week, 1700. "
                   "Override columns always win.")
    return ws


# --------------------------------------------------------------------------
def build_control(wb):
    """Holidays + generated class-day calendar."""
    ws = wb.create_sheet("Control")
    ws.sheet_view.showGridLines = False
    # holidays B:D
    header_row(ws, ["Holiday", "Observed (Yr 1)", "Observed (Yr 2)", None,
                    "Extra Closure Dates", None, "Day #", "Class Date",
                    "Week #", "In Session?"])
    first = DATA_ROW
    r = first
    for name, tmpl in DL.HOLIDAYS:
        ws.cell(row=r, column=2, value=name).font = F_LABEL
        c1 = ws.cell(row=r, column=3,
                     value="=" + tmpl.format(Y="YEAR(cfgStartDate)"))
        c2 = ws.cell(row=r, column=4,
                     value="=" + tmpl.format(Y="YEAR(cfgStartDate)+1"))
        for c in (c1, c2):
            c.number_format = DATE
            c.font = F_CALC
            c.fill = FILL_CALC
            c.border = BOX
        r += 1
    hol_last = r - 1
    define(wb, "nrHolidays1", "Control", f"$C${first}:$C${hol_last}")
    define(wb, "nrHolidays2", "Control", f"$D${first}:$D${hol_last}")
    # extra closures (manual)
    extra_last = first + 14
    fill_rows(ws, first, extra_last, {"F": (None, "in")})
    for r2 in range(first, extra_last + 1):
        ws[f"F{r2}"].number_format = DATE
    define(wb, "nrExtraClosures", "Control", f"$F${first}:$F${extra_last}")
    # class-day calendar H:K
    cd_last = first + ROWS_CLASSDAYS - 1
    fill_rows(ws, first, cd_last, {
        "H": (('v', None), "fx"),
        "I": ('IF($H{r}="","",WORKDAY.INTL(cfgStartDate-1,$H{r},"0000011",'
              'nrAllClosures))', "fx"),
        "J": ('IF($H{r}="","",INT(($I{r}-(cfgStartDate-WEEKDAY(cfgStartDate,2)+1))/7)+1)', "fx"),
        "K": ('IF($H{r}="","",IF($I{r}<=cfgEndDate,"Yes","No"))', "fx"),
    })
    for i, r2 in enumerate(range(first, cd_last + 1)):
        ws[f"H{r2}"].value = i + 1
        ws[f"I{r2}"].number_format = DATE
    # combined closures name (holidays yr1+yr2+extra) via helper column M
    ws.cell(row=HDR_ROW, column=13, value="AllClosures").font = F_SMALL
    n_h = hol_last - first + 1
    for i in range(n_h):
        ws.cell(row=first + i, column=13, value=f"=C{first+i}").number_format = DATE
        ws.cell(row=first + n_h + i, column=13,
                value=f"=D{first+i}").number_format = DATE
    for i in range(15):
        ws.cell(row=first + 2 * n_h + i, column=13,
                value=f'=IF(F{first+i}="",DATE(1900,1,1),F{first+i})'
                ).number_format = DATE
    all_last = first + 2 * n_h + 14
    define(wb, "nrAllClosures", "Control", f"$M${first}:$M${all_last}")
    define(wb, "nrCDnum", "Control", f"$H${first}:$H${cd_last}")
    define(wb, "nrCDdate", "Control", f"$I${first}:$I${cd_last}")
    define(wb, "nrCDweek", "Control", f"$J${first}:$J${cd_last}")
    define(wb, "nrCDinsession", "Control", f"$K${first}:$K${cd_last}")
    col_widths(ws, {"A": 3, "B": 26, "C": 15, "D": 15, "E": 2, "F": 16,
                    "G": 2, "H": 8, "I": 13, "J": 8, "K": 11, "L": 2, "M": 12})
    sheet_note(ws, "Class days = Mon-Fri from Start Date, skipping observed "
                   "holidays and any Extra Closure Dates you add. Retest "
                   "deadlines, writing dates and sign-in sheets all key off "
                   "this calendar.")
    return ws


# --------------------------------------------------------------------------
def build_schedule(wb):
    ws = wb.create_sheet("Schedule")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Date", "Day", "Start", "End", "Hours",
                    "Chapter / Activity", "Ch #", "Instructor", "Location",
                    "Week #", "Day #", "Notes", "Instructor OK?"])
    first, last = DATA_ROW, DATA_ROW + ROWS_SCHEDULE - 1
    fill_rows(ws, first, last, {
        "B": (None, "in"),
        "C": ('IF($B{r}="","",TEXT($B{r},"ddd"))', "fx"),
        "D": (None, "in"), "E": (None, "in"),
        "F": ('IF(OR($B{r}="",$D{r}="",$E{r}=""),"",ROUND(($E{r}-$D{r})*24,2))', "fx"),
        "G": (None, "in"),
        "H": ('IF($G{r}="","",IFERROR(XLOOKUP($G{r},nrCHname,nrCHnum),'
              'IFERROR(XLOOKUP($G{r},nrSUBname,nrSUBparent),"")))', "fx"),
        "I": (None, "in"), "J": (None, "in"),
        "K": ('IF($B{r}="","",IFERROR(XLOOKUP($B{r},nrCDdate,nrCDweek),""))', "fx"),
        "L": ('IF($B{r}="","",IFERROR(XLOOKUP($B{r},nrCDdate,nrCDnum),""))', "fx"),
        "M": (None, "in"),
        "N": ('IF(OR($B{r}="",$I{r}=""),"",'
              'IF(SUMPRODUCT(--ISNUMBER(SEARCH(nrInstrNames,$I{r})))>0,"OK",'
              '"UNRECOGNIZED"))', "fx"),
    })
    for r2 in range(first, last + 1):
        ws[f"B{r2}"].number_format = DATE
        ws[f"D{r2}"].number_format = TIME
        ws[f"E{r2}"].number_format = TIME
    dv_list(ws, "=nrScheduleItems", [f"G{first}:G{last}"])
    # instructor dropdown depends on the row's topic: offers that topic's
    # "Teaching this academy" picks from InstructorBanks; topics without a
    # bank row fall back to the full roster. Warning-only so multi-name
    # combined entries (built by the pick-to-append macro) stay valid.
    dv_list(ws, f"=IF(IFERROR(MATCH($G{first},nrBankTopics,0),0)=0,"
                f"nrInstrNames,INDEX(nrBankSel,"
                f"IFERROR(MATCH($G{first},nrBankTopics,0),1),0))",
            [f"I{first}:I{last}"], enforce=False)
    define(wb, "nrSCH_Date", "Schedule", f"$B${first}:$B${last}")
    define(wb, "nrSCH_Start", "Schedule", f"$D${first}:$D${last}")
    define(wb, "nrSCH_End", "Schedule", f"$E${first}:$E${last}")
    define(wb, "nrSCH_Hrs", "Schedule", f"$F${first}:$F${last}")
    define(wb, "nrSCH_Act", "Schedule", f"$G${first}:$G${last}")
    define(wb, "nrSCH_ChNum", "Schedule", f"$H${first}:$H${last}")
    define(wb, "nrSCH_Instr", "Schedule", f"$I${first}:$I${last}")
    define(wb, "nrSCH_Loc", "Schedule", f"$J${first}:$J${last}")
    define(wb, "nrSCH_InstrOK", "Schedule", f"$N${first}:$N${last}")
    cf_formula(ws, f"N{first}:N{last}", f'$N{first}="UNRECOGNIZED"',
               FILL_WARNBG)
    col_widths(ws, {"A": 3, "B": 12, "C": 6, "D": 9, "E": 9, "F": 8,
                    "G": 46, "H": 6, "I": 24, "J": 16, "K": 8, "L": 7,
                    "M": 30, "N": 13})
    page_setup_landscape(ws, print_area=f"B{HDR_ROW}:M{first+400}",
                         repeat_rows=f"{HDR_ROW}:{HDR_ROW}")
    sheet_note(ws, "One row per time block (a day usually has several). "
                   "Hours, chapter reconciliation, sign-in sheets, writing "
                   "dates and eval headers all read from here. Printable "
                   "as-is (File > Print).")
    return ws


def build_schedule_items_helper(wb):
    """Hidden helper: chapter names + activities as one dropdown list."""
    ws = wb.create_sheet("sysListsHelper")
    ws.sheet_view.showGridLines = False
    r = DATA_ROW
    ws.cell(row=HDR_ROW, column=2, value="Schedule dropdown items").font = F_LABEL
    n_ch = len(DC.CHAPTERS)
    for i in range(n_ch):
        ws.cell(row=r + i, column=2,
                value=f"=IF(ChapterMaster!D{DATA_ROW+i}=\"\",\"\","
                      f"ChapterMaster!D{DATA_ROW+i})")
    n_sub = len(DC.SUBTOPICS)
    for j, (name, _p, _t) in enumerate(DC.SUBTOPICS):
        ws.cell(row=r + n_ch + j, column=2, value=name)
    for j, act in enumerate(DC.ACTIVITIES):
        ws.cell(row=r + n_ch + n_sub + j, column=2, value=act)
    last = r + n_ch + n_sub + len(DC.ACTIVITIES) - 1
    define(wb, "nrScheduleItems", "sysListsHelper", f"$B${r}:$B${last}")
    ws.sheet_state = "hidden"
    return ws


def build_all_config(wb):
    build_starthere(wb)
    build_settings(wb)
    build_lists(wb)
    build_agencies(wb)
    build_instructors(wb)
    build_instructorbanks(wb)
    build_chaptermaster(wb)
    build_exammaster(wb)
    build_examplan(wb)
    build_skillsmaster(wb)
    build_spellingmaster(wb)
    build_writingmaster(wb)
    build_control(wb)
    build_schedule_items_helper(wb)
    build_schedule(wb)

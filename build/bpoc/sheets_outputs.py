"""Output & printable sheets: Dashboard, ScoresGrid, Ranking, WatchList,
CadetProfile, Transcript, GradChecklist, DismissalLog, Audit, PrintCenter,
SignIn, EvalSheet, SpellingPrint, WritingHandout, EmailPreview, EmailLog,
NamedRanges registry.
"""
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference

from xlb import (
    HDR_ROW, DATA_ROW, CADETS, CADET_LAST, F_HDR, FILL_HDR, F_CALC, FILL_CALC,
    F_LABEL, F_SMALL, F_BODY, F_INPUT, FILL_INPUT, FILL_BAND, FILL_YELLOW,
    F_KPI, F_TITLE, F_SECTION, FILL_NAVY, FILL_STEEL, A_LEFT, A_LEFT_WRAP,
    A_RIGHT, A_CENTER, BOX, UNDER, DATE, header_row, fill_rows, dv_list,
    sheet_note, cf_yes_no, cf_formula, FILL_WARNBG, FILL_OKBG, FILL_AMBER,
    col_widths, define, section_bar, label, protect, unlock_range,
    page_setup_portrait, page_setup_landscape,
)
import data_lists as DL
import data_writing as DW

FIRST, LAST = DATA_ROW, CADET_LAST


def _kpi(ws, row, col, title, formula, wide=2):
    ws.merge_cells(start_row=row, start_column=col, end_row=row,
                   end_column=col + wide - 1)
    t = ws.cell(row=row, column=col, value=title)
    t.font = F_SMALL
    t.alignment = A_LEFT
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1,
                   end_column=col + wide - 1)
    v = ws.cell(row=row + 1, column=col, value="=" + formula)
    v.font = F_KPI
    v.alignment = A_LEFT
    for c in range(col, col + wide):
        ws.cell(row=row, column=c).fill = FILL_BAND
        ws.cell(row=row + 1, column=c).fill = FILL_BAND
    return v


# --------------------------------------------------------------------------
def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value='=cfgAcademyClass&" — Dashboard"').font = F_KPI
    ws.cell(row=r, column=8, value='="As of "&TEXT(TODAY(),"mm/dd/yyyy")'
            ).font = F_SMALL
    g = ws.cell(row=r, column=10, value=(
        '=HYPERLINK("#InputGuide!B5","Data Entry Guide →")'))
    g.font = F_LABEL
    r += 2
    # key dates (final/state/graduation pull from the Schedule when entered)
    _kpi(ws, r, 2, "Academy start",
         'TEXT(cfgStartDate,"mm/dd/yyyy")')
    _kpi(ws, r, 4, "Final Exam date",
         'LET(d,MAXIFS(nrSCH_Date,nrSCH_Act,"Final Test"),'
         'IF(d=0,"(not on schedule)",TEXT(d,"mm/dd/yyyy")))')
    _kpi(ws, r, 6, "State Exam date",
         'LET(d,MAXIFS(nrSCH_Date,nrSCH_Act,"State Test"),'
         'IF(d=0,"(not on schedule)",TEXT(d,"mm/dd/yyyy")))')
    _kpi(ws, r, 8, "Graduation / end",
         'LET(d,MAXIFS(nrSCH_Date,nrSCH_Act,"Graduation"),'
         'IF(d=0,TEXT(cfgEndDate,"mm/dd/yyyy"),TEXT(d,"mm/dd/yyyy")))')
    _kpi(ws, r, 10, "Active cadets",
         'SUMPRODUCT((nrCadetStatus="Active")*1)')
    r += 3
    section_bar(ws, r, 2, 11, "Today")
    r += 1
    ws.cell(row=r, column=8, value=(
        '=IFERROR("Training Day #"&XLOOKUP(TODAY(),nrCDdate,nrCDnum)&'
        '"  (Week "&XLOOKUP(TODAY(),nrCDdate,nrCDweek)&")",'
        '"(not a class day)")')).font = F_LABEL
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(FILTER(HSTACK(TEXT(nrSCH_Start,"h:mm AM/PM"),'
        'TEXT(nrSCH_End,"h:mm AM/PM"),nrSCH_Act,nrSCH_Instr,nrSCH_Loc),'
        'nrSCH_Date=TODAY()),7),"— no schedule entered for today —")'))
    r += 7
    ws.cell(row=r, column=2, value=(
        '=IF(COUNTIF(nrDL_Date,TODAY())=0,'
        '"DailyLog: no entry for today yet","DailyLog: entered ✓")'
    )).font = F_SMALL
    r += 2
    c = ws.cell(row=r, column=2, value=(
        '="TODAY — "&UPPER(TEXT(TODAY(),"dddd, mmmm d"))&'
        'IFERROR(" — TRAINING DAY #"&XLOOKUP(TODAY(),nrCDdate,nrCDnum)&'
        '", WEEK "&XLOOKUP(TODAY(),nrCDdate,nrCDweek),"")'))
    c.font = F_SECTION
    for col in range(2, 12):
        ws.cell(row=r, column=col).fill = FILL_STEEL
    ws.row_dimensions[r].height = 16
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(FILTER(HSTACK(TEXT(nrSCH_Start,"h:mm AM/PM"),'
        'TEXT(nrSCH_End,"h:mm AM/PM"),nrSCH_Act,nrSCH_Instr,nrSCH_Loc),'
        'nrSCH_Date=TODAY()),8),"No class scheduled today")'))
    r += 8
    _kpi(ws, r, 2, "Separated / total enrolled",
         'SUMPRODUCT((nrCadetStatus="Separated")*1)&" / "&'
         'SUMPRODUCT((nrCadetStatus<>"")*1)')
    _kpi(ws, r, 4, "Class current avg",
         # AVERAGE(IF(...)) is a legacy CSE array formula; stored as an
         # ordinary formula it silently evaluates to 0. AVERAGEIF needs no
         # array entry and is what the rest of the sheet already uses.
         'IFERROR(ROUND(AVERAGEIF(nrCadetStatus,"Active",nrGRcurrent),1),"—")')
    _kpi(ws, r, 6, "Graduation eligible",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrCKgradElig="Yes"))')
    _kpi(ws, r, 8, "Flagged cadets",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrFLcount>0))')
    _kpi(ws, r, 10, "Retests overdue / undated",
         'COUNTIF(nrES_RetStat,"OVERDUE")+COUNTIF(nrES_RetStat,"CHECK DATE")')
    r += 3
    _kpi(ws, r, 2, "Spelling interventions",
         'COUNTIF(nrSpellFlag,"INTERVENTION")')
    _kpi(ws, r, 4, "Makeup owed (cadets)",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrATTmakeupOK="No"))')
    _kpi(ws, r, 6, "Delivered hours",
         'ROUND(N(nrCHtotalDelivered),0)&" / "&cfgRequiredHours')
    _kpi(ws, r, 8, "Audit checks failing",
         'SUMPRODUCT((nrAUDstatus<>"OK")*(nrAUDstatus<>""))')
    _kpi(ws, r, 10, "Days to graduation",
         'MAX(0,cfgEndDate-TODAY())')
    _kpi(ws, r, 12, "Cert copies to collect",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrCERTmissing<>""))')
    r += 3
    section_bar(ws, r, 2, 11, "Watch list — highest flag counts first "
                              "(full list on WatchList)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(rngCadetNames,nrCadetAgency,'
        'nrFLcount,nrFLreasons),(nrFLcount>0)*(nrCadetStatus="Active")),'
        'FILTER(nrFLcount,(nrFLcount>0)*(nrCadetStatus="Active")),-1),11),'
        '"No flags — clear")'))
    watch_top = r
    r += 11
    section_bar(ws, r, 2, 11, "Certification reminders — copies to collect "
                              "from cadets (Certifications sheet)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(FILTER(HSTACK(rngCadetNames,nrCadetAgency,'
        'nrCERTmissing),(nrCERTmissing<>"")*(nrCadetStatus="Active")),9),'
        '"All certification copies collected")'))
    r += 9
    section_bar(ws, r, 2, 11, "Open missed-time events — not yet made up "
                              "(Attendance ↔ Makeup by EventID)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(nrAT_ID,Attendance!$D$6:$D$805,'
        'TEXT(nrAT_Date,"mm/dd"),Attendance!$G$6:$G$805,'
        'Attendance!$H$6:$H$805,nrAT_Balance),(nrAT_Cleared="OPEN")),'
        'FILTER(nrAT_Date,(nrAT_Cleared="OPEN")),-1),9),'
        '"All missed time cleared")'))
    r += 9
    section_bar(ws, r, 2, 11, "Outstanding memos (pending / overdue)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(Memos!$B$6:$B$305,nrME_Cadet,'
        'TEXT(nrME_Assigned,"mm/dd"),nrME_Ref,nrME_Subject,'
        'TEXT(nrME_Due,"mm/dd"),nrME_Status),'
        '(nrME_Cadet<>"")*(nrME_Received="")),'
        'FILTER(nrME_Due,(nrME_Cadet<>"")*(nrME_Received="")),1),8),'
        '"No memos outstanding")'))
    r += 8
    section_bar(ws, r, 2, 11, "Class average by exam")
    chart_anchor_row = r + 1
    r += 16
    section_bar(ws, r, 2, 11, "Spelling class average by test")
    chart2_anchor = r + 1

    # charts read the ScoresGrid class-average row and Spelling stats
    ch = LineChart()
    ch.title = "Class average by exam (recorded, final attempts)"
    ch.height = 7
    ch.width = 22
    data = Reference(wb["ScoresGrid"], min_col=4, max_col=28,
                     min_row=CADET_LAST + 2, max_row=CADET_LAST + 2)
    cats = Reference(wb["ScoresGrid"], min_col=4, max_col=28, min_row=HDR_ROW)
    # the data is one ROW (the class-average row): without from_rows openpyxl
    # emits one single-point series per column (25 of them), so the line has
    # nothing to connect and every marker lands on category 1
    ch.add_data(data, titles_from_data=False, from_rows=True)
    ch.set_categories(cats)
    ch.legend = None
    ch.y_axis.scaling.min = 50
    ch.y_axis.scaling.max = 100
    ws.add_chart(ch, f"B{chart_anchor_row}")

    sp = wb["Spelling"]
    sr = CADET_LAST + 2
    ch2 = BarChart()
    ch2.title = "Spelling class average by test"
    ch2.height = 7
    ch2.width = 22
    data2 = Reference(sp, min_col=4, max_col=15, min_row=sr + 1, max_row=sr + 1)
    cats2 = Reference(sp, min_col=4, max_col=15, min_row=HDR_ROW)
    ch2.add_data(data2, titles_from_data=False, from_rows=True)
    ch2.set_categories(cats2)
    ch2.legend = None
    ch2.y_axis.scaling.min = 0
    ch2.y_axis.scaling.max = 100
    ws.add_chart(ch2, f"B{chart2_anchor}")

    col_widths(ws, {"A": 3, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14,
                    "G": 14, "H": 14, "I": 14, "J": 14, "K": 40})
    return ws


# --------------------------------------------------------------------------
def build_scoresgrid(wb):
    ws = wb.create_sheet("ScoresGrid")
    ws.sheet_view.showGridLines = False
    n_exams = 25
    hdrs = ["PID", "Cadet Name"]
    for i in range(n_exams):
        hdrs.append(f'=IF(INDEX(rngEPcode,{i+1})="","",INDEX(rngEPcode,{i+1}))')
    hdrs += ["Current Grade", "Rank"]
    # header formulas need manual write (header_row writes plain values)
    header_row(ws, ["PID", "Cadet Name"] + [None] * n_exams +
               ["Current Grade", "Rank"])
    for i in range(n_exams):
        c = ws.cell(row=HDR_ROW, column=4 + i,
                    value=f'=IF(INDEX(rngEPcode,{i+1})="","",INDEX(rngEPcode,{i+1}))')
    cols = {
        "B": ('IF(Cadets!$B{r}="","",Cadets!$B{r})', "fx"),
        "C": ('IF($B{r}="","",Cadets!$F{r})', "fx"),
    }
    for i in range(n_exams):
        cl = get_column_letter(4 + i)
        # blank (not 0) when the cadet has no recorded final attempt:
        # SUMIFS returns 0 on no match, so guard with COUNTIFS first
        cols[cl] = (
            'IF(OR($B{r}="",%s$%d=""),"",IF(COUNTIFS(nrES_PID,$B{r},'
            'nrES_Code,%s$%d,nrES_Final,"Yes",nrES_Rec,">=0")=0,"",'
            'SUMIFS(nrES_Rec,nrES_PID,$B{r},nrES_Code,%s$%d,'
            'nrES_Final,"Yes")))'
            % (cl, HDR_ROW, cl, HDR_ROW, cl, HDR_ROW), "fx")
    gcol = get_column_letter(4 + n_exams)
    rcol = get_column_letter(5 + n_exams)
    cols[gcol] = ('IF($B{r}="","",sysGrades!$M{r})', "fx")
    cols[rcol] = ('IF($B{r}="","",IFERROR(sysGrades!$U{r},""))', "fx")
    fill_rows(ws, FIRST, LAST, cols)
    # class average row
    ar = LAST + 2
    ws.cell(row=ar, column=3, value="Class average:").font = F_LABEL
    for i in range(n_exams):
        cl = get_column_letter(4 + i)
        # average recorded final attempts straight from the log so untaken
        # exams / cadets with no attempt can never drag the average down
        ws[f"{cl}{ar}"] = (f'=IF({cl}{HDR_ROW}="","",IFERROR(ROUND('
                           f'AVERAGEIFS(nrES_Rec,nrES_Code,{cl}{HDR_ROW},'
                           f'nrES_Final,"Yes"),1),""))')
        ws[f"{cl}{ar}"].font = F_CALC
    define(wb, "nrSGclassavg", "ScoresGrid",
           f"$D${ar}:${get_column_letter(3+n_exams)}${ar}")
    # sub-70 highlighting
    lastcl = get_column_letter(3 + n_exams)
    cf_formula(ws, f"D{FIRST}:{lastcl}{LAST}",
               f'AND(D{FIRST}<>"",ISNUMBER(D{FIRST}),D{FIRST}<cfgPassingScore)',
               FILL_WARNBG)
    col_widths(ws, {"A": 3, "B": 10, "C": 24})
    for i in range(n_exams):
        ws.column_dimensions[get_column_letter(4 + i)].width = 7
    sheet_note(ws, "Recorded score of each exam's final attempt (retest cap "
                   "applied). Red = below 70.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_ranking(wb):
    ws = wb.create_sheet("Ranking")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Rank", "Cadet", "Agency", "Current Grade", "Major Avg",
                    "Minor Avg", "Spelling Avg", "Final", "Grad Elig?"])
    ws.cell(row=FIRST, column=2, value=(
        '=IFERROR(SORT(FILTER(HSTACK(nrGRrank,rngCadetNames,nrCadetAgency,'
        'nrGRcurrent,nrGRmajavg,sysGrades!$J$%d:$J$%d,sysGrades!$K$%d:$K$%d,'
        'nrGRfinal,nrCKgradElig),ISNUMBER(nrGRrank)),1),"No eligible cadets yet")'
        % (FIRST, LAST, FIRST, LAST)))
    col_widths(ws, {"A": 3, "B": 7, "C": 24, "D": 20, "E": 13, "F": 11,
                    "G": 11, "H": 12, "I": 9, "J": 11})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:J{LAST+5}",
                        repeat_rows=f"{HDR_ROW}:{HDR_ROW}")
    sheet_note(ws, "Live class ranking (dismissal-review, failed-out and "
                   "separated cadets are unranked per policy).")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_watchlist(wb):
    ws = wb.create_sheet("WatchList")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Cadet", "Agency", "Flags", "Reasons", "Current Grade",
                    "Attendance Cl%", "Overdue Writing", "Open Counseling"])
    ws.cell(row=FIRST, column=2, value=(
        '=IFERROR(SORTBY(FILTER(HSTACK(rngCadetNames,nrCadetAgency,nrFLcount,'
        'nrFLreasons,nrGRcurrent,TEXT(nrATTclPct,"0%%"),nrWRoverdue,'
        'sysIncidents!$M$%d:$M$%d),(nrFLcount>0)*(nrCadetStatus="Active")),'
        'FILTER(nrFLcount,(nrFLcount>0)*(nrCadetStatus="Active")),-1),'
        '"No flagged cadets")' % (FIRST, LAST)))
    col_widths(ws, {"A": 3, "B": 24, "C": 18, "D": 7, "E": 64, "F": 13,
                    "G": 13, "H": 14, "I": 15})
    page_setup_landscape(ws, print_area=f"B{HDR_ROW}:I{LAST+5}",
                         repeat_rows=f"{HDR_ROW}:{HDR_ROW}")
    sheet_note(ws, "Every threshold is adjustable on Settings. Document your "
                   "response on the Counseling log — that's the early-"
                   "intervention record policy 300.4.B expects.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def _profile_label(ws, r, c, text, val_formula, wide=2):
    ws.cell(row=r, column=c, value=text).font = F_LABEL
    ws.merge_cells(start_row=r, start_column=c + 1, end_row=r,
                   end_column=c + wide)
    v = ws.cell(row=r, column=c + 1, value="=" + val_formula)
    v.font = F_BODY
    return v


def build_cadetprofile(wb):
    ws = wb.create_sheet("CadetProfile")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Cadet:").font = F_LABEL
    sel = ws.cell(row=r, column=3)
    sel.fill = FILL_INPUT
    sel.font = F_INPUT
    sel.border = BOX
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    dv_list(ws, "=rngCadetNames", [f"C{r}"])
    define(wb, "cfgProfileCadet", "CadetProfile", f"$C${r}")
    ws.cell(row=r, column=7, value='=cfgAcademyClass&" — Cadet Profile — "'
            '&TEXT(TODAY(),"mm/dd/yyyy")').font = F_SMALL
    P = 'XLOOKUP(cfgProfileCadet,rngCadetNames,'
    r += 2
    section_bar(ws, r, 2, 10, "Identity & standing")
    r += 1
    _profile_label(ws, r, 2, "PID", P + 'rngCadetPIDs,"")')
    _profile_label(ws, r, 5, "Agency", P + 'nrCadetAgency,"")')
    _profile_label(ws, r, 8, "Status", P + 'nrCadetStatus,"")')
    r += 1
    _profile_label(ws, r, 2, "Current grade", P + 'nrGRcurrent,"")')
    _profile_label(ws, r, 5, "Rank", 'IFERROR(' + P + 'nrGRrank),"unranked")')
    _profile_label(ws, r, 8, "Grad eligible", P + 'nrCKgradElig,"")')
    r += 1
    _profile_label(ws, r, 2, "Blocking issues", P + 'nrCKblocking,"")', wide=8)
    r += 2
    section_bar(ws, r, 2, 10, "Academics (pass = 70 in each category)")
    r += 1
    # NOTE: "Final exam" is the FINAL-EXAM average (sysGrades L = nrGRfinalExam),
    # not the weighted composite (sysGrades N = nrGRfinal, shown as "Current
    # grade" above) — this block is the 70-in-each-category test.
    for lab2, rng in (("Major avg", "nrGRmajavg"),
                      ("Minor avg", f"sysGrades!$J${FIRST}:$J${LAST}"),
                      ("Spelling avg", f"sysGrades!$K${FIRST}:$K${LAST}"),
                      ("Final exam avg", "nrGRfinalExam")):
        _profile_label(ws, r, 2, lab2, P + rng + ',"—")')
        r += 1
    r -= 4
    _profile_label(ws, r, 5, "Needs on remaining Major",
                   P + f'sysGrades!$W${FIRST}:$W${LAST},"—")')
    _profile_label(ws, r + 1, 5, "Needs on remaining Minor",
                   P + f'sysGrades!$X${FIRST}:$X${LAST},"—")')
    _profile_label(ws, r + 2, 5, "Retakes used",
                   P + 'nrGRretakes,"—")')
    _profile_label(ws, r + 3, 5, "Consecutive fails",
                   P + 'nrGRconsec,"—")')
    _profile_label(ws, r, 8, "Spelling flag",
                   P + 'nrSpellFlag,"—")')
    _profile_label(ws, r + 1, 8, "Last vs prev score",
                   P + f'sysGrades!$Z${FIRST}:$Z${LAST},"—")&" / "&'
                   + P + f'sysGrades!$AA${FIRST}:$AA${LAST},"—")')
    r += 5
    section_bar(ws, r, 2, 10, "Attendance / PT / Skills / Writing")
    r += 1
    _profile_label(ws, r, 2, "Classroom missed (net min)",
                   P + f'sysAttendance!$G${FIRST}:$G${LAST},"—")')
    _profile_label(ws, r, 5, "Tier", P + 'nrATTclTier,"—")')
    _profile_label(ws, r, 8, "Makeup owed (min)", P + 'nrATTclOwed,"—")')
    r += 1
    _profile_label(ws, r, 2, "PT sessions missed (net)",
                   P + f'sysAttendance!$O${FIRST}:$O${LAST},"—")')
    _profile_label(ws, r, 5, "PT tier", P + 'nrATTptTier,"—")')
    _profile_label(ws, r, 8, "Final PT", P + 'nrPT_FinalPass,"—")')
    r += 1
    _profile_label(ws, r, 2, "Skills status",
                   'IF(' + P + 'nrSKfailedout,"—")="Yes","FAILED OUT",'
                   'IF(' + P + 'nrSKelig,"—")="Yes","On track","Remediation"))')
    _profile_label(ws, r, 5, "Firearms avg", P + 'nrSKfirearmsAvg,"—")')
    _profile_label(ws, r, 8, "Writing current",
                   P + 'nrWRcurrent,"—")&" ("&' + P +
                   'nrWRoverdue,"0")&" overdue)"')
    r += 2
    section_bar(ws, r, 2, 10, "Flags")
    r += 1
    _profile_label(ws, r, 2, "Flag reasons", P + 'nrFLreasons,"none")', wide=8)
    r += 2
    section_bar(ws, r, 2, 10, "Recent incidents & counseling (10 latest)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(TEXT(nrIN_Date,"mm/dd/yyyy"),'
        'nrIN_Dir,nrIN_Sev,'
        'nrIN_Desc),nrIN_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,'
        'rngCadetPIDs,"")),FILTER(nrIN_Date,nrIN_PID=XLOOKUP(cfgProfileCadet,'
        'rngCadetNames,rngCadetPIDs,"")),-1),10),"none")'))
    r += 11
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(TEXT(nrCO_Date,"mm/dd/yyyy"),'
        'nrCO_Type,nrCO_Desc),'
        'nrCO_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,rngCadetPIDs,"")),'
        'FILTER(nrCO_Date,nrCO_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,'
        'rngCadetPIDs,"")),-1),10),"none")'))
    r += 11
    section_bar(ws, r, 2, 10, "Attendance exceptions & makeup (10 latest)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(TEXT(nrAT_Date,"mm/dd/yyyy"),'
        'nrAT_Type,'
        'Attendance!$H$6:$H$805,nrAT_Min,nrAT_Sess,nrAT_Excused),'
        'nrAT_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,rngCadetPIDs,"")),'
        'FILTER(nrAT_Date,nrAT_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,'
        'rngCadetPIDs,"")),-1),10),"none")'))
    r += 11
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(TEXT(Makeup!$C$6:$C$505,'
        '"mm/dd/yyyy"),nrMK_Type,'
        'nrMK_Min,nrMK_Sess,nrMK_Credit),'
        'nrMK_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,rngCadetPIDs,"")),'
        'FILTER(Makeup!$C$6:$C$505,nrMK_PID=XLOOKUP(cfgProfileCadet,'
        'rngCadetNames,rngCadetPIDs,"")),-1),10),"no makeup entries")'))
    r += 11
    section_bar(ws, r, 2, 10, "OPEN missed time — 6 most recent "
                              "(event / date / type / reason / balance owed)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(nrAT_ID,TEXT(nrAT_Date,"mm/dd"),'
        'Attendance!$G$6:$G$805,Attendance!$H$6:$H$805,nrAT_Balance),'
        '(nrAT_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,rngCadetPIDs,""))*'
        '(nrAT_Cleared="OPEN")),FILTER(nrAT_Date,'
        '(nrAT_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,rngCadetPIDs,""))*'
        '(nrAT_Cleared="OPEN")),-1),6),"all missed time cleared")'))
    r += 6
    section_bar(ws, r, 2, 10, "OPEN memos — 7 shown (id / assigned / subject "
                              "/ due / status)")
    r += 1
    # TAKE(...,7) matches the 7 rows reserved below and the B5:J86 print
    # area: uncapped, an 8th open memo spilled off the bottom of the page
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(FILTER(HSTACK(Memos!$B$6:$B$305,'
        'TEXT(nrME_Assigned,"mm/dd"),nrME_Subject,TEXT(nrME_Due,"mm/dd"),'
        'nrME_Status),(nrME_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,'
        'rngCadetPIDs,""))*(nrME_Cadet<>"")*(nrME_Received="")),7),'
        '"none outstanding")'))
    r += 6
    col_widths(ws, {"A": 3, "B": 24, "C": 14, "D": 14, "E": 16, "F": 14,
                    "G": 14, "H": 18, "I": 14, "J": 14})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:J{r}")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")   # C5 is the only picker
    sheet_note(ws, "Pick a cadet, print (Print Center button or File > "
                   "Print). Everything on one page for agency calls and "
                   "review boards.")
    return ws


# --------------------------------------------------------------------------
def build_transcript(wb):
    ws = wb.create_sheet("Transcript")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Cadet:").font = F_LABEL
    sel = ws.cell(row=r, column=3)
    sel.fill = FILL_INPUT
    sel.font = F_INPUT
    sel.border = BOX
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    dv_list(ws, "=rngCadetNames", [f"C{r}"])
    define(wb, "cfgTranscriptCadet", "Transcript", f"$C${r}")
    P = 'XLOOKUP(cfgTranscriptCadet,rngCadetNames,'
    ws.cell(row=r, column=7, value=(
        '="TYLER POLICE ACADEMY — "&cfgAcademyClass&" — OFFICIAL TRANSCRIPT"'
    )).font = F_LABEL
    r += 2
    _profile_label(ws, r, 2, "PID", P + 'rngCadetPIDs,"")')
    _profile_label(ws, r, 5, "Agency", P + 'nrCadetAgency,"")')
    _profile_label(ws, r, 8, "Status", P + 'nrCadetStatus,"")')
    r += 1
    _profile_label(ws, r, 2, "Academy dates",
                   'TEXT(cfgStartDate,"mm/dd/yyyy")&" – "&TEXT(cfgEndDate,"mm/dd/yyyy")')
    _profile_label(ws, r, 5, "Final grade", P + 'nrGRfinal,"—")')
    _profile_label(ws, r, 8, "Class rank", 'IFERROR(' + P + 'nrGRrank),"—")')
    r += 2
    section_bar(ws, r, 2, 10, "Category averages (70 required in each)")
    r += 1
    # the four CATEGORY averages the 70-in-each rule is tested against;
    # the Final Exam row reads sysGrades L (final-exam average), NOT the
    # weighted composite in N which is already printed as "Final grade"
    for lab2, rng, w in (("Major (40%)", "nrGRmajavg", None),
                         ("Minor (30%)", f"sysGrades!$J${FIRST}:$J${LAST}", None),
                         ("Spelling (10%)", f"sysGrades!$K${FIRST}:$K${LAST}", None),
                         ("Final Exam (20%)", "nrGRfinalExam", None)):
        _profile_label(ws, r, 2, lab2, P + rng + ',"—")')
        r += 1
    r += 1
    section_bar(ws, r, 2, 10, "Exam record (recorded scores, final attempts)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(FILTER(HSTACK(ExamScores!$G$%d:$G$%d,ExamScores!$H$%d:$H$%d,'
        'ExamScores!$K$%d:$K$%d,ExamScores!$M$%d:$M$%d,'
        'TEXT(ExamScores!$S$%d:$S$%d,"mm/dd/yyyy")),'
        '(nrES_PID=%srngCadetPIDs,""))*(nrES_Final="Yes")*(nrES_Rec<>"")),"—")'
        % (FIRST, FIRST + 1499, FIRST, FIRST + 1499, FIRST, FIRST + 1499,
           FIRST, FIRST + 1499, FIRST, FIRST + 1499, P)))
    # cap to the reserved window so a long log can never #SPILL! into the
    # section bar below (25 exams max on the plan)
    v17 = ws.cell(row=r, column=2).value
    assert v17.startswith('=IFERROR(FILTER(') and v17.endswith('),"—")')
    ws.cell(row=r, column=2).value = (
        '=IFERROR(TAKE(' + v17[len('=IFERROR('):-len(',"—")')] + ',26),"—")')
    r += 26
    section_bar(ws, r, 2, 10, "Attendance / Skills / PT / Writing / Conduct")
    r += 1
    _profile_label(ws, r, 2, "Classroom minutes missed (net)",
                   P + f'sysAttendance!$G${FIRST}:$G${LAST},"0")&" of "&'
                   'cfgClassroomCapMinutes&" cap"', wide=4)
    r += 1
    _profile_label(ws, r, 2, "PT sessions missed (net)",
                   P + f'sysAttendance!$O${FIRST}:$O${LAST},"0")&" of "&'
                   'cfgPTCapSessions&" cap"', wide=4)
    r += 1
    _profile_label(ws, r, 2, "Makeup complete", P + 'nrATTmakeupOK,"—")', wide=4)
    r += 1
    _profile_label(ws, r, 2, "Skills", 'IF(' + P + 'nrSKfailedout,"—")="Yes",'
                   '"FAILED OUT","Qualified/on record — see Skills log")', wide=4)
    r += 1
    _profile_label(ws, r, 2, "Firearms average", P + 'nrSKfirearmsAvg,"—")', wide=4)
    r += 1
    _profile_label(ws, r, 2, "Final PT assessment",
                   P + 'nrPT_FinalPass,"—")&" ("&' + P +
                   'nrPT_FinalPts,"—")&" pts)"', wide=4)
    r += 1
    _profile_label(ws, r, 2, "Writing assignments complete",
                   P + 'nrWRcomplete,"0")&" of 40"', wide=4)
    r += 1
    _profile_label(ws, r, 2, "Incidents (neg/pos)",
                   P + f'sysIncidents!$F${FIRST}:$F${LAST},"0")&" / "&'
                   + P + f'sysIncidents!$G${FIRST}:$G${LAST},"0")', wide=4)
    r += 1
    _profile_label(ws, r, 2, "Awards",
                   'IFERROR(TEXTJOIN(", ",TRUE,IF(nrAWfinal=cfgTranscriptCadet,'
                   'nrAWnames,"")),"")', wide=4)
    r += 1
    _profile_label(ws, r, 2, "Required certifications",
                   'LET(m,' + P + 'nrCERTmissing,""),'
                   'IF(m="","All on file (TIM, SFST, TCIC, CPR/AED, ALERRT, '
                   'ICS)","Outstanding: "&m))', wide=4)
    r += 1
    _profile_label(ws, r, 2, "State licensing exam",
                   P + 'nrSEstatus,"Not yet attempted")', wide=4)
    r += 2
    ws.cell(row=r, column=2, value="Training Coordinator:").font = F_LABEL
    ws.cell(row=r, column=5, value="_______________________").font = F_BODY
    ws.cell(row=r, column=7, value="Date:").font = F_LABEL
    ws.cell(row=r, column=8, value="____________").font = F_BODY
    r += 2
    ws.cell(row=r, column=2, value=DL.ACADEMY_ADDRESS).font = F_SMALL
    col_widths(ws, {"A": 3, "B": 26, "C": 14, "D": 14, "E": 16, "F": 14,
                    "G": 16, "H": 14, "I": 12, "J": 12})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:J{r}")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")   # C5 is the only picker
    sheet_note(ws, "The end-of-academy deliverable per cadet — print one per "
                   "cadet for the agency packet.")
    return ws


# --------------------------------------------------------------------------
def build_gradchecklist(wb):
    ws = wb.create_sheet("GradChecklist")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Cadet", "Agency", "Academic", "Classroom",
                    "PT Sessions", "Skills", "Incidents", "Writing",
                    "Makeup", "Final PT", "No Dismiss Rev", "Certs",
                    "ELIGIBLE", "Blocking Issues"])
    cols = {
        "B": ('IF(Cadets!$B{r}="","",Cadets!$B{r})', "fx"),
        "C": ('IF($B{r}="","",Cadets!$F{r})', "fx"),
        "D": ('IF($B{r}="","",Cadets!$H{r})', "fx"),
        "E": ('IF($B{r}="","",sysChecks!$E{r})', "fx"),
        "F": ('IF($B{r}="","",sysChecks!$F{r})', "fx"),
        "G": ('IF($B{r}="","",sysChecks!$G{r})', "fx"),
        "H": ('IF($B{r}="","",sysChecks!$H{r})', "fx"),
        "I": ('IF($B{r}="","",sysChecks!$I{r})', "fx"),
        "J": ('IF($B{r}="","",sysChecks!$J{r})', "fx"),
        "K": ('IF($B{r}="","",sysChecks!$K{r})', "fx"),
        "L": ('IF($B{r}="","",sysChecks!$L{r})', "fx"),
        "M": ('IF($B{r}="","",IF(sysChecks!$M{r}="Yes","No","Yes"))', "fx"),
        "N": ('IF($B{r}="","",sysChecks!$Q{r})', "fx"),
        "O": ('IF($B{r}="","",sysChecks!$N{r})', "fx"),
        "P": ('IF($B{r}="","",sysChecks!$O{r})', "fx"),
    }
    fill_rows(ws, FIRST, LAST, cols)
    cf_yes_no(ws, f"E{FIRST}:O{LAST}")
    # anything that is not a plain "Yes" is a block — "Pending" / "Not taken"
    # (final PT never assessed, rubric never configured) must read as red,
    # not as an uncoloured neutral value on the graduation gate
    cf_formula(ws, f"E{FIRST}:O{LAST}",
               f'AND(E{FIRST}<>"",E{FIRST}<>"Yes")', FILL_WARNBG)
    col_widths(ws, {"A": 3, "B": 10, "C": 24, "D": 18, "P": 54})
    for cl in "EFGHIJKLMNO":
        ws.column_dimensions[cl].width = 10
    page_setup_landscape(ws, print_area=f"B{HDR_ROW}:P{LAST}",
                         repeat_rows=f"{HDR_ROW}:{HDR_ROW}")
    sheet_note(ws, "The final gate before the ceremony — columns Academic "
                   "through Certs must ALL read Yes and Blocking Issues must "
                   "be empty. Anything else blocks and is spelled out in "
                   "Blocking Issues: 'No', 'Pending' (Final PT points rubric "
                   "not entered on Settings) and 'Not taken' (final PT never "
                   "assessed) are all blocks, not neutral values.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_dismissallog(wb):
    ws = wb.create_sheet("DismissalLog")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["ReviewID", "Cadet Name", "PID", "Review Type",
                    "Trigger", "Opened", "Board/Reviewer", "Decision Date",
                    "Outcome", "Asst. Chief Approval", "Approval Date",
                    "Agency Notified", "Docs Ref", "Notes"])
    first, last = DATA_ROW, DATA_ROW + 99
    fill_rows(ws, first, last, {
        "B": ('IF($C{r}="","","R"&TEXT(ROW()-%d,"000"))' % HDR_ROW, "fx"),
        "C": (None, "in"),
        "D": ('IF($C{r}="","",IFERROR(INDEX(rngCadetPIDs,MATCH($C{r},rngCadetNames,0)),"?"))', "fx"),
        "E": (None, "in"), "F": (None, "in"), "G": (None, "in"),
        "H": (None, "in"), "I": (None, "in"), "J": (None, "in"),
        "K": (None, "in"), "L": (None, "in"), "M": (None, "in"),
        "N": (None, "in"), "O": (None, "in"),
    })
    for r in range(first, last + 1):
        for cl in ("G", "I", "L"):
            ws[f"{cl}{r}"].number_format = DATE
    dv_list(ws, "=rngCadetNames", [f"C{first}:C{last}"])
    dv_list(ws, "=lstReviewType", [f"E{first}:E{last}"])
    dv_list(ws, "=lstReviewOutcome", [f"J{first}:J{last}"])
    dv_list(ws, "=lstYesNo", [f"K{first}:K{last}", f"M{first}:M{last}"])
    col_widths(ws, {"A": 3, "B": 9, "C": 22, "D": 9, "E": 16, "F": 28,
                    "G": 11, "H": 20, "I": 12, "J": 15, "K": 16, "L": 12,
                    "M": 13, "N": 14, "O": 30})
    sheet_note(ws, "Formal record of every dismissal/academic/conduct review: "
                   "trigger, decision, Assistant-Chief approval (policy "
                   "600.2.E) and agency notification. This is the file "
                   "TCOLE/legal asks for.")
    return ws


# --------------------------------------------------------------------------
def build_audit(wb):
    ws = wb.create_sheet("Audit")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value=(
        '="TCOLE AUDIT READINESS — "&cfgAcademyClass&" — "&TEXT(TODAY(),"mm/dd/yyyy")'
    )).font = F_KPI
    r += 2
    section_bar(ws, r, 2, 8, "Program checks (live from the engine)")
    r += 1
    hdr2 = r
    header_row(ws, ["Check", "Value", "Target", "Status", "Detail"], row=r)
    r += 1
    # driven off the engine's own list — a hard-coded count meant a check
    # added to sysAudit was counted by the Dashboard "Audit checks failing"
    # tile but never printed on the packet handed to TCOLE, so the tile went
    # red with no visible cause
    from sheets_engine import AUDIT_CHECKS
    n_checks = len(AUDIT_CHECKS)
    for i in range(n_checks):
        sr = FIRST + i
        ws.cell(row=r, column=2, value=f"=sysAudit!$B${sr}").font = F_BODY
        ws.cell(row=r, column=3, value=f"=sysAudit!$C${sr}").font = F_CALC
        ws.cell(row=r, column=4, value=f"=sysAudit!$D${sr}").font = F_CALC
        s = ws.cell(row=r, column=5, value=f"=sysAudit!$E${sr}")
        s.font = F_LABEL
        ws.cell(row=r, column=6, value=f"=sysAudit!$F${sr}").font = F_SMALL
        r += 1
    cf_formula(ws, f"E{hdr2+1}:E{r-1}", f'AND($E{hdr2+1}<>"OK",$E{hdr2+1}<>"")',
               FILL_WARNBG)
    cf_formula(ws, f"E{hdr2+1}:E{r-1}", f'$E{hdr2+1}="OK"', FILL_OKBG)
    r += 1
    section_bar(ws, r, 2, 8, "Requirement sources — what is TCOLE rule vs "
                             "Academy policy")
    r += 1
    for txt in (
        "TCOLE Rule (Ch. 215/218/219): course #1000736 minimum hours per "
        "chapter, licensed/SME instructors, course records retention, "
        "exam security, student enrollment standards (L2/L3, medical), "
        "attendance records.",
        "Academy Policy (May 2026 manual): weighted grading (40/30/10/20), "
        "70-in-each-category rule, 12 spelling tests across the academy + "
        "75% intervention, "
        "40 writing assignments, PT program and caps, awards.",
    ):
        c = ws.cell(row=r, column=2, value="• " + txt)
        c.font = F_SMALL
        c.alignment = A_LEFT_WRAP
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 26
        r += 1
    r += 1
    section_bar(ws, r, 2, 8, "Program-level TCOLE requirements (mark Yes "
                             "when satisfied)")
    r += 1
    header_row(ws, ["Requirement", None, None, None, "Met?", "Notes"], row=r)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    r += 1
    prg_first = r
    import data_chapters as DC
    for item, detail in DC.PROGRAM_REQS:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c = ws.cell(row=r, column=2, value=item + " — " + detail)
        c.font = F_BODY
        c.alignment = A_LEFT_WRAP
        m = ws.cell(row=r, column=6)
        m.fill = FILL_INPUT
        m.font = F_INPUT
        m.border = BOX
        n = ws.cell(row=r, column=7)
        n.fill = FILL_INPUT
        n.font = F_INPUT
        n.border = BOX
        ws.row_dimensions[r].height = 26
        r += 1
    prg_last = r - 1
    dv_list(ws, '"Yes,No,N/A"', [f"F{prg_first}:F{prg_last}"])
    define(wb, "nrPRGitems", "Audit", f"$B${prg_first}:$B${prg_last}")
    define(wb, "nrPRGmet", "Audit", f"$F${prg_first}:$F${prg_last}")
    cf_formula(ws, f"F{prg_first}:F{prg_last}",
               f'AND($B{prg_first}<>"",$F{prg_first}<>"Yes",$F{prg_first}<>"N/A")',
               FILL_WARNBG)
    r += 1
    section_bar(ws, r, 2, 8, "Per-cadet enrollment documents (mark Yes when "
                             "on file)")
    r += 1
    doc_hdr = r
    header_row(ws, ["Cadet", "Enroll App", "TCLEDDS L1", "Medical (L2)",
                    "Psych (L3)", "Background", "Photo ID/DL", "Rules Ack",
                    "All Docs?"], row=r)
    r += 1
    doc_first = r
    doc_last = r + CADETS - 1
    # (no `cols` dict here: the grid is written cell-by-cell below. A dead
    # copy used to sit at this spot still describing the PRE-shift layout —
    # AllDocs in column I counting C:H = 6 — which is exactly the stale
    # column-letter bug this workbook keeps regressing on.)
    for rr in range(doc_first, doc_last + 1):
        src = FIRST + (rr - doc_first)
        ws.cell(row=rr, column=2,
                value=f'=IF(Cadets!$B{src}="","",Cadets!$F{src})').font = F_CALC
        for c in range(3, 10):
            cc = ws.cell(row=rr, column=c)
            cc.fill = FILL_INPUT
            cc.font = F_INPUT
            cc.border = BOX
        ws.cell(row=rr, column=10, value=(
            f'=IF($B{rr}="","",IF(COUNTIF(C{rr}:I{rr},"Yes")=7,"Yes","No"))'
        )).font = F_CALC
    dv_list(ws, "=lstYesNo", [f"C{doc_first}:I{doc_last}"])
    cf_yes_no(ws, f"J{doc_first}:J{doc_last}")
    col_widths(ws, {"A": 3, "B": 30, "C": 11, "D": 11, "E": 12, "F": 11,
                    "G": 12, "H": 12, "I": 10})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:J{doc_last}",
                        repeat_rows=f"{doc_hdr}:{doc_hdr}")
    sheet_note(ws, "Top: live program checks. Bottom: per-cadet enrollment "
                   "file checklist. Chapter-level records live on "
                   "ChapterMaster; instructor credentials on Instructors.")
    return ws


# --------------------------------------------------------------------------
def build_addendum(wb):
    """Printable excess-hours report: per-chapter delivered vs TCOLE minimum,
    with the correct separate-reporting course number for each excess."""
    import data_chapters as DC
    ws = wb.create_sheet("Addendum")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value=(
        '="TYLER POLICE ACADEMY — "&cfgAcademyClass&'
        '" — EXCESS HOURS REPORTING (ADDENDUM) — "&TEXT(TODAY(),"mm/dd/yyyy")'
    )).font = F_KPI
    r += 1
    c = ws.cell(row=r, column=2, value=(
        "The BPOC is reported to TCOLE at exactly 736 hours. Hours delivered "
        "beyond a chapter's TCOLE minimum are reported separately: Arrest & "
        "Control, Driving and Firearms under their own course numbers; all "
        "other excess under the Addendum to BPOC (#101). Verify course "
        "numbers against current TCOLE reporting guidance before submission."))
    c.font = F_SMALL
    c.alignment = A_LEFT_WRAP
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 40
    r += 2
    hdr = r
    header_row(ws, ["Ch #", "Chapter / Class", "TCOLE Min Hrs",
                    "Delivered Hrs", "Excess Hrs", "Report Under"], row=r)
    r += 1
    first = r
    n = len(DC.CHAPTERS)
    for i in range(n):
        rr = first + i
        src = DATA_ROW + i          # ChapterMaster data row
        ch = DC.CHAPTERS[i][1]
        ws.cell(row=rr, column=2, value=f"=ChapterMaster!$C${src}").font = F_BODY
        ws.cell(row=rr, column=3, value=f"=ChapterMaster!$D${src}").font = F_BODY
        ws.cell(row=rr, column=4, value=f"=ChapterMaster!$E${src}").font = F_CALC
        ws.cell(row=rr, column=5, value=(
            f'=IF(ChapterMaster!$G{src}=0,"",ChapterMaster!$G{src})'
        )).font = F_CALC
        ws.cell(row=rr, column=6, value=(
            f'=IF(OR(ChapterMaster!$G{src}="",ChapterMaster!$G{src}=0),"",'
            f'IF(ChapterMaster!$G{src}>ChapterMaster!$E{src},'
            f'ROUND(ChapterMaster!$G{src}-ChapterMaster!$E{src},2),""))'
        )).font = F_LABEL
        report_as = DC.SEPARATE_REPORT.get(ch, DC.ADDENDUM_COURSE)
        ws.cell(row=rr, column=7, value=(
            f'=IF($F{rr}="","","{report_as}")')).font = F_BODY
        for ccol in range(2, 8):
            ws.cell(row=rr, column=ccol).border = BOX
    last = first + n - 1
    cf_formula(ws, f"F{first}:F{last}",
               f'AND($F{first}<>"",$F{first}>0)', FILL_AMBER)
    r = last + 2
    ws.cell(row=r, column=3, value="Total excess to report:").font = F_LABEL
    tot = ws.cell(row=r, column=6, value=f"=SUM(F{first}:F{last})")
    tot.font = F_KPI
    ws.cell(row=r + 1, column=3,
            value="— of which under #101 (Addendum):").font = F_LABEL
    ws.cell(row=r + 1, column=6, value=(
        f'=SUMPRODUCT(($F{first}:$F{last}<>"")*'
        f'($G{first}:$G{last}="{DC.ADDENDUM_COURSE}")*'
        f'IFERROR($F{first}:$F{last}+0,0))')).font = F_CALC
    ws.cell(row=r + 2, column=3,
            value="— of which under separate course #s:").font = F_LABEL
    ws.cell(row=r + 2, column=6, value=(
        f'=SUMPRODUCT(($F{first}:$F{last}<>"")*'
        f'($G{first}:$G{last}<>"{DC.ADDENDUM_COURSE}")*'
        f'($G{first}:$G{last}<>"")*IFERROR($F{first}:$F{last}+0,0))'
    )).font = F_CALC
    r += 4
    ws.cell(row=r, column=2, value="Training Coordinator:").font = F_LABEL
    ws.cell(row=r, column=4, value="_______________________").font = F_BODY
    ws.cell(row=r, column=6, value="Date:").font = F_LABEL
    ws.cell(row=r, column=7, value="____________").font = F_BODY
    r += 2
    ws.cell(row=r, column=2, value=DL.ACADEMY_ADDRESS).font = F_SMALL
    col_widths(ws, {"A": 3, "B": 7, "C": 46, "D": 12, "E": 12, "F": 11,
                    "G": 30, "H": 4})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:G{r}",
                        repeat_rows=f"{hdr}:{hdr}")
    protect(ws)
    sheet_note(ws, "Fills from ChapterMaster/Schedule automatically — only "
                   "chapters with logged hours show; excess rows highlight. "
                   "Print for the TCOLE reporting packet.")
    return ws


# --------------------------------------------------------------------------
def build_chapterpacket(wb):
    """The auditor's 'show me your file for chapter X' page: pick a chapter,
    everything about it appears — hours, blocks, instructors, file status,
    special requirements, linked exam stats."""
    ws = wb.create_sheet("ChapterPacket")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Chapter:").font = F_LABEL
    sel = ws.cell(row=r, column=3, value="1")
    sel.fill = FILL_INPUT
    sel.font = F_INPUT
    sel.border = BOX
    dv_list(ws, "=nrCHnum", [f"C{r}"])
    define(wb, "cfgPacketChapter", "ChapterPacket", f"$C${r}")
    C = "cfgPacketChapter"
    ws.cell(row=r, column=5, value=(
        f'="TCOLE TRAINING FILE — "&cfgAcademyClass&" — Ch. "&{C}&" "'
        f'&IFERROR(XLOOKUP({C},nrCHnum,nrCHname),"")')).font = F_KPI
    r += 2
    X = f'IFERROR(XLOOKUP({C},nrCHnum,'
    _profile_label(ws, r, 2, "Module", X + 'nrCHmod),"")')
    _profile_label(ws, r, 5, "TCOLE min hrs", X + 'nrCHmin),"")')
    _profile_label(ws, r, 8, "Delivered hrs", X + 'nrCHdeliv),"")')
    r += 1
    _profile_label(ws, r, 2, "First taught",
                   'IFERROR(TEXT(XLOOKUP(' + C + ',nrCHnum,nrCHfirst),'
                   '"mm/dd/yyyy"),"")')
    _profile_label(ws, r, 5, "Last taught",
                   'IFERROR(TEXT(MAXIFS(nrSCH_Date,nrSCH_ChNum,' + C + '),'
                   '"mm/dd/yyyy"),"")')
    _profile_label(ws, r, 8, "vs TCOLE min",
                   'LET(d,IFERROR(XLOOKUP(' + C + ',nrCHnum,nrCHdeliv),0),'
                   'm,IFERROR(XLOOKUP(' + C + ',nrCHnum,nrCHmin),0),'
                   'IF(d=0,"no hours logged",TEXT(d-m,"+0.##;-0.##;0")))')
    r += 2
    section_bar(ws, r, 2, 9, "Training-file contents (required by the IRG)")
    r += 1
    for lab2, rng in (("Lesson plan (SME)", "nrCHlesson"),
                      ("Instructor bio(s)", "nrCHbio"),
                      ("Sign-in sheets (w/ PID)", "nrCHsignin"),
                      ("Assessment", "nrCHassess"),
                      ("Grade sheet", "nrCHgrade"),
                      ("Course evaluations", "nrCHevals"),
                      ("Handouts/PPT (optional)", "nrCHhandoutOpt")):
        _profile_label(ws, r, 2, lab2, X + rng + '),"")')
        r += 1
    r -= 7
    _profile_label(ws, r, 6, "Special TCOLE requirement",
                   X + 'nrCHspecial),"")&""', wide=4)
    _profile_label(ws, r + 1, 6, "Special requirement met",
                   X + 'nrCHspecialMet),"—")')
    _profile_label(ws, r + 2, 6, "File complete?", X + 'nrCHfileOK),"No")')
    lex = ('IFERROR(INDEX(rngEPcode,MATCH(' + C + ',rngEPch,0)),"")')
    _profile_label(ws, r + 3, 6, "Linked exam",
                   'IF(' + lex + '="","(none)",' + lex +
                   '&" — "&IFERROR(INDEX(rngEPname,MATCH(' + C +
                   ',rngEPch,0)),""))', wide=4)
    _profile_label(ws, r + 4, 6, "Exam class avg / low / fails",
                   'LET(c,' + lex + ',IF(c="","—",'
                   'IFERROR(ROUND(AVERAGEIFS(nrES_Rec,nrES_Code,c,'
                   'nrES_Final,"Yes"),1),"—")&" / "&'
                   'IFERROR(MINIFS(nrES_Rec,nrES_Code,c,nrES_Final,"Yes"),"—")'
                   '&" / "&COUNTIFS(nrES_Code,c,nrES_Final,"Yes",'
                   'nrES_Rec,"<"&cfgPassingScore)))', wide=4)
    r += 8
    section_bar(ws, r, 2, 9, "Instructors who taught this chapter "
                             "(from the schedule)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(FILTER(HSTACK(nrInstrNames,nrInstrReady,'
        'nrInstrChTaught),'
        'ISNUMBER(SEARCH(", "&' + C + '&",",", "&nrInstrChTaught&","))),12),'
        '"— none on the schedule yet —")'))
    r += 12
    section_bar(ws, r, 2, 9, "Schedule blocks delivered")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(FILTER(HSTACK(TEXT(nrSCH_Date,"mm/dd/yyyy"),'
        'TEXT(nrSCH_Start,"h:mm AM/PM"),TEXT(nrSCH_End,"h:mm AM/PM"),'
        'nrSCH_Act,nrSCH_Instr),nrSCH_ChNum=' + C + '),32),'
        '"— no blocks scheduled —")'))
    r += 32
    ws.cell(row=r, column=2, value="Training Coordinator:").font = F_LABEL
    ws.cell(row=r, column=4, value="_______________________").font = F_BODY
    ws.cell(row=r, column=6, value="Date:").font = F_LABEL
    ws.cell(row=r, column=7, value="____________").font = F_BODY
    col_widths(ws, {"A": 3, "B": 24, "C": 16, "D": 16, "E": 14, "F": 22,
                    "G": 16, "H": 16, "I": 22})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:I{r}")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")
    sheet_note(ws, "Pick a chapter — the whole training file's status on one "
                   "printable page. Pair with the ExamSheet grade sheet and "
                   "the EvalSheet critique for the paper folder.")
    return ws


# --------------------------------------------------------------------------
def build_examsheet(wb):
    """IRG-required grade sheet per assessment: pick an exam, print the
    class's scores with pass/fail, retests, stats and a proctor signature."""
    ws = wb.create_sheet("ExamSheet")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Exam seq #:").font = F_LABEL
    sel = ws.cell(row=r, column=3, value=1)
    sel.fill = FILL_INPUT
    sel.font = F_INPUT
    sel.border = BOX
    define(wb, "cfgGradeSheetExam", "ExamSheet", f"$C${r}")
    E = "cfgGradeSheetExam"
    r += 2
    ws.cell(row=r, column=2, value=(
        '="TYLER POLICE ACADEMY — "&cfgAcademyClass&" — GRADE SHEET — "'
        '&IFERROR(INDEX(rngEPname,MATCH(' + E + ',IFERROR(rngEPseq+0,-1),0)),'
        '"(set exam #)")')).font = F_KPI
    r += 1
    ws.cell(row=r, column=2, value=(
        '="Exam code: "&IFERROR(INDEX(rngEPcode,MATCH(' + E +
        ',IFERROR(rngEPseq+0,-1),0)),"?")&"   Administered: "&'
        'IFERROR(TEXT(MAXIFS(nrSCH_Date,nrSCH_Act,"Test "&' + E + '),'
        '"mm/dd/yyyy"),"")&"   Passing score: "&cfgPassingScore&'
        '"   (passed retests record at "&cfgRetakeRecordedCap&")"'
    )).font = F_SMALL
    r += 2
    hdr = r
    header_row(ws, ["#", "Cadet", "PID", "Agency", "Raw Score",
                    "Recorded", "Pass?", "Retest"], row=r)
    r += 1
    first = r
    code = ('IFERROR(INDEX(rngEPcode,MATCH(' + E +
            ',IFERROR(rngEPseq+0,-1),0)),"")')
    for i in range(CADETS):
        rr = first + i
        src = FIRST + i
        ws.cell(row=rr, column=2, value=(
            f'=IF(Cadets!$B{src}="","",ROW()-{first-1})')).font = F_CALC
        ws.cell(row=rr, column=3, value=(
            f'=IF(Cadets!$B{src}="","",Cadets!$F{src})')).font = F_BODY
        ws.cell(row=rr, column=4, value=(
            f'=IF($C{rr}="","",Cadets!$B{src})')).font = F_CALC
        ws.cell(row=rr, column=5, value=(
            f'=IF($C{rr}="","",Cadets!$H{src})')).font = F_CALC
        ws.cell(row=rr, column=6, value=(
            f'=IF($C{rr}="","",LET(c,{code},IF(COUNTIFS(nrES_PID,'
            f'Cadets!$B{src},nrES_Code,c,nrES_Att,1)=0,"",'
            f'SUMIFS(nrES_Raw,nrES_PID,Cadets!$B{src},nrES_Code,c,'
            f'nrES_Att,1))))')).font = F_CALC
        ws.cell(row=rr, column=7, value=(
            f'=IF($C{rr}="","",LET(c,{code},IF(COUNTIFS(nrES_PID,'
            f'Cadets!$B{src},nrES_Code,c,nrES_Final,"Yes",nrES_Rec,"<>")=0,'
            f'"",SUMIFS(nrES_Rec,nrES_PID,Cadets!$B{src},nrES_Code,c,'
            f'nrES_Final,"Yes"))))')).font = F_CALC
        ws.cell(row=rr, column=8, value=(
            f'=IF(OR($C{rr}="",$G{rr}=""),"",'
            f'IF($G{rr}>=cfgPassingScore,"Pass","FAIL"))')).font = F_CALC
        ws.cell(row=rr, column=9, value=(
            f'=IF($C{rr}="","",LET(c,{code},IF(COUNTIFS(nrES_PID,'
            f'Cadets!$B{src},nrES_Code,c,nrES_Att,2)>0,"Retested","")))'
        )).font = F_CALC
        for ccol in range(2, 10):
            ws.cell(row=rr, column=ccol).border = BOX
    last = first + CADETS - 1
    cf_formula(ws, f"H{first}:H{last}", f'$H{first}="FAIL"', FILL_WARNBG)
    r = last + 1
    ws.cell(row=r, column=3, value="Class average / low / high / fails:"
            ).font = F_LABEL
    ws.cell(row=r, column=6, value=(
        '=LET(c,' + code + ',IF(c="","",'
        'IFERROR(ROUND(AVERAGEIFS(nrES_Rec,nrES_Code,c,nrES_Final,"Yes"),1),"—")'
        '&" / "&IFERROR(MINIFS(nrES_Rec,nrES_Code,c,nrES_Final,"Yes"),"—")'
        '&" / "&IFERROR(MAXIFS(nrES_Rec,nrES_Code,c,nrES_Final,"Yes"),"—")'
        '&" / "&COUNTIFS(nrES_Code,c,nrES_Final,"Yes",'
        'nrES_Rec,"<"&cfgPassingScore)))')).font = F_CALC
    r += 2
    ws.cell(row=r, column=2, value="Proctor / Instructor: ____________________"
            "____     Training Coordinator: ________________________"
            ).font = F_BODY
    r += 1
    ws.cell(row=r, column=2, value=DL.ACADEMY_ADDRESS).font = F_SMALL
    col_widths(ws, {"A": 3, "B": 5, "C": 28, "D": 10, "E": 18, "F": 11,
                    "G": 11, "H": 9, "I": 10})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW+2}:I{r}",
                        repeat_rows=f"{hdr}:{hdr}")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")
    sheet_note(ws, "The IRG requires a grade sheet for each assessment in "
                   "the training file — print one per exam and file it with "
                   "the chapter packet.")
    return ws


# --------------------------------------------------------------------------
def build_signin(wb):
    ws = wb.create_sheet("SignIn")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Sign-in date:").font = F_LABEL
    d = ws.cell(row=r, column=3)
    d.fill = FILL_INPUT
    d.font = F_INPUT
    d.border = BOX
    d.number_format = DATE
    define(wb, "cfgSignInDate", "SignIn", f"$C${r}")
    ws.cell(row=r, column=5, value=(
        '=IFERROR("Training Day #"&XLOOKUP(cfgSignInDate,nrCDdate,nrCDnum)'
        '&"  (Week "&XLOOKUP(cfgSignInDate,nrCDdate,nrCDweek)&")","")'
    )).font = F_LABEL
    r += 2
    ws.cell(row=r, column=2, value=(
        '="TYLER POLICE ACADEMY — "&cfgAcademyClass&" — DAILY REPORT & '
        'ATTENDANCE ROSTER — "&TEXT(cfgSignInDate,"dddd, mmmm d, yyyy")'
    )).font = F_KPI
    r += 2
    section_bar(ws, r, 2, 9, "Instruction scheduled this date")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(FILTER(HSTACK(TEXT(nrSCH_Start,"h:mm AM/PM"),'
        'TEXT(nrSCH_End,"h:mm AM/PM"),nrSCH_Act,nrSCH_Instr,nrSCH_Loc),'
        'nrSCH_Date=cfgSignInDate),9),'
        '"— no schedule entered for this date —")'))
    r += 9
    section_bar(ws, r, 2, 9, "AM roll call")
    r += 1
    ws.cell(row=r, column=2, value=(
        "Present: ______ of ______     Absent/Late (name & reason): "
        "____________________________________________")).font = F_BODY
    ws.row_dimensions[r].height = 20
    r += 1
    ws.cell(row=r, column=2, value=(
        "PT: [ ] Full participation   [ ] Modified: ______________   "
        "[ ] Missed: ______________")).font = F_BODY
    ws.row_dimensions[r].height = 20
    r += 1
    section_bar(ws, r, 2, 9, "Cadet sign-in")
    r += 1
    header_row(ws, ["#", "Cadet", "PID", "Agency", "AM Signature",
                    "PM Signature", "Remarks"], row=r)
    r += 1
    first = r
    for i in range(CADETS):
        rr = first + i
        src = FIRST + i
        # gate the row number on the NAME (C), not merely on enrolment:
        # a separated cadet used to get a numbered but BLANK signature line
        # on the printed roster. COUNTIF keeps the numbering contiguous.
        ws.cell(row=rr, column=2, value=(
            f'=IF(C{rr}="","",COUNTIF($C${first}:C{rr},"?*"))')).font = F_CALC
        ws.cell(row=rr, column=3, value=(
            f'=IF(Cadets!$B{src}="","",IF(Cadets!$I{src}<>"Active","",'
            f'Cadets!$F{src}))')).font = F_BODY
        ws.cell(row=rr, column=4, value=(
            f'=IF(C{rr}="","",Cadets!$B{src})')).font = F_BODY
        ws.cell(row=rr, column=5, value=(
            f'=IF(C{rr}="","",Cadets!$H{src})')).font = F_BODY
        for c in range(2, 9):
            ws.cell(row=rr, column=c).border = BOX
        ws.row_dimensions[rr].height = 20
    last = first + CADETS - 1
    r = last + 1
    section_bar(ws, r, 2, 9, "PM report — changes through the day")
    r += 1
    ws.cell(row=r, column=2, value=(
        "Early departures (name / time out / reason — e.g. agency recall): "
        "______________________________________________")).font = F_BODY
    ws.row_dimensions[r].height = 20
    r += 1
    ws.cell(row=r, column=2, value=(
        "____________________________________________________________"
        "____________________________________________")).font = F_BODY
    ws.row_dimensions[r].height = 20
    r += 1
    ws.cell(row=r, column=2, value=(
        "Incidents / injuries: _______________________________________"
        "____________________________________________")).font = F_BODY
    ws.row_dimensions[r].height = 20
    r += 1
    ws.cell(row=r, column=2, value=(
        "Memos turned in: ____________________________   Remarks: "
        "_____________________________________________")).font = F_BODY
    ws.row_dimensions[r].height = 20
    r += 1
    ws.cell(row=r, column=2, value=(
        "Class Leader signature: ______________________________     "
        "Coordinator: ______________________________")).font = F_BODY
    ws.row_dimensions[r].height = 22
    r += 1
    ws.cell(row=r, column=2, value="Instructor verification: ____________________"
            "____________     Time: ____________").font = F_BODY
    r += 1
    ws.cell(row=r, column=2, value=DL.ACADEMY_ADDRESS).font = F_SMALL
    col_widths(ws, {"A": 3, "B": 6, "C": 28, "D": 10, "E": 18, "F": 21,
                    "G": 21, "H": 22, "I": 12})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:H{r}")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")
    sheet_note(ws, "The one-page daily report & roster: AM roll call, "
                   "sign-in, PM changes, signatures. Print on demand; the "
                   "signed original is scanned into the file (scans = "
                   "originals per records policy) and the day is logged as "
                   "one DailyLog row. Exceptions still go on Attendance.")
    return ws


# --------------------------------------------------------------------------
def build_evalsheet(wb):
    ws = wb.create_sheet("EvalSheet")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Chapter:").font = F_LABEL
    sel = ws.cell(row=r, column=3)
    sel.fill = FILL_INPUT
    sel.font = F_INPUT
    sel.border = BOX
    dv_list(ws, "=nrCHnum", [f"C{r}"])
    define(wb, "cfgEvalChapter", "EvalSheet", f"$C${r}")
    ws.cell(row=r, column=5, value='="(prints for: "&IFERROR(XLOOKUP('
            'cfgEvalChapter,nrCHnum,nrCHname),"pick a chapter")&")"'
            ).font = F_SMALL
    r += 2
    ws.cell(row=r, column=2, value="TYLER POLICE ACADEMY").font = F_KPI
    r += 1
    ws.cell(row=r, column=2, value=DL.ACADEMY_ADDRESS).font = F_SMALL
    r += 2
    ws.cell(row=r, column=2, value="CLASS EVALUATION AND ASSESSMENT").font = F_LABEL
    r += 1
    ws.cell(row=r, column=2, value=DL.EVAL_SCALE).font = F_SMALL
    r += 2
    _profile_label(ws, r, 2, "Course",
                   'cfgAcademyClass&" — Ch. "&cfgEvalChapter&" "&'
                   'IFERROR(XLOOKUP(cfgEvalChapter,nrCHnum,nrCHname),"")', wide=5)
    r += 1
    _profile_label(ws, r, 2, "Instructor",
                   'IFERROR(XLOOKUP(cfgEvalChapter,nrCHnum,nrCHinstr),"")', wide=3)
    _profile_label(ws, r, 6, "Date(s)",
                   'IFERROR(TEXT(XLOOKUP(cfgEvalChapter,nrCHnum,nrCHfirst),'
                   '"mm/dd")&" – "&TEXT(MAXIFS(nrSCH_Date,nrSCH_ChNum,'
                   'cfgEvalChapter),"mm/dd/yyyy"),"")', wide=2)
    r += 2
    header_row(ws, ["#", "Question", "1", "2", "3", "4", "5"], row=r)
    r += 1
    for i, q in enumerate(DL.EVAL_QUESTIONS):
        ws.cell(row=r, column=2, value=i + 1).font = F_BODY
        qc = ws.cell(row=r, column=3, value=q)
        qc.font = F_BODY
        qc.alignment = A_LEFT_WRAP
        for c in range(4, 9):
            ws.cell(row=r, column=c).border = BOX
        ws.cell(row=r, column=2).border = BOX
        qc.border = BOX
        ws.row_dimensions[r].height = 26
        r += 1
    r += 1
    ws.cell(row=r, column=2, value="ADDITIONAL COMMENTS").font = F_LABEL
    r += 1
    for i in range(6):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.cell(row=r, column=2).border = UNDER
        ws.row_dimensions[r].height = 20
        r += 1
    c = ws.cell(row=r, column=2, value=DL.EVAL_FOOTER)
    c.font = F_SMALL
    c.alignment = A_LEFT_WRAP
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 28
    col_widths(ws, {"A": 3, "B": 5, "C": 62, "D": 6, "E": 6, "F": 6,
                    "G": 6, "H": 6})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW+2}:H{r}")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")
    sheet_note(ws, "Modernized replica of the existing critique form. Pick a "
                   "chapter — course, instructor and dates fill from the "
                   "schedule. Print one per cadet (Print Center does stacks).")
    return ws


# --------------------------------------------------------------------------
def build_spellingprint(wb):
    ws = wb.create_sheet("SpellingPrint")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Test #:").font = F_LABEL
    s1 = ws.cell(row=r, column=3, value=1)
    s1.fill = FILL_INPUT
    s1.font = F_INPUT
    s1.border = BOX
    define(wb, "cfgSpellPrintNum", "SpellingPrint", f"$C${r}")
    ws.cell(row=r, column=4, value="Mode:").font = F_LABEL
    s2 = ws.cell(row=r, column=5, value="Test")
    s2.fill = FILL_INPUT
    s2.font = F_INPUT
    s2.border = BOX
    dvm = dv_list(ws, '"Test,Key"', [f"E{r}"])
    define(wb, "cfgSpellPrintMode", "SpellingPrint", f"$E${r}")
    r += 2
    ws.cell(row=r, column=2, value="TYLER POLICE DEPARTMENT ACADEMY").font = F_KPI
    r += 1
    ws.cell(row=r, column=2, value=(
        '="Spelling Test #"&cfgSpellPrintNum&IF(cfgSpellPrintMode="Key",'
        '" — KEY","")')).font = F_LABEL
    r += 2
    ws.cell(row=r, column=2, value="Name: ______________________    "
            "Date: ____________    Score: ________").font = F_BODY
    r += 2
    first = r
    for i in range(13):
        rr = first + i
        ws.cell(row=rr, column=2, value=f"{i+1}.").font = F_BODY
        w1 = ws.cell(row=rr, column=3, value=(
            f'=IF(cfgSpellPrintMode="Key",INDEX(nrSpellWords,{i+1},'
            f'cfgSpellPrintNum),"_______________________")'))
        w1.font = F_BODY
        if i + 14 <= 25:
            ws.cell(row=rr, column=5, value=f"{i+14}.").font = F_BODY
            w2 = ws.cell(row=rr, column=6, value=(
                f'=IF(cfgSpellPrintMode="Key",INDEX(nrSpellWords,{i+14},'
                f'cfgSpellPrintNum),"_______________________")'))
            w2.font = F_BODY
        ws.row_dimensions[rr].height = 24
    r = first + 14
    ws.cell(row=r, column=2, value="Each word is worth 4 points.").font = F_SMALL
    col_widths(ws, {"A": 3, "B": 5, "C": 30, "D": 4, "E": 5, "F": 30})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW+2}:F{r}")
    protect(ws)
    # two real pickers here (C5 = test #, E5 = Test/Key mode); D5 holds the
    # "Mode:" label and stays locked so it cannot be erased
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")
    unlock_range(ws, f"E{HDR_ROW}:E{HDR_ROW}")
    sheet_note(ws, "Pick test # and Test/Key mode; print. Words come from "
                   "SpellingMaster.")
    return ws


# --------------------------------------------------------------------------
def build_writinghandout(wb):
    ws = wb.create_sheet("WritingHandout")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Week #:").font = F_LABEL
    s1 = ws.cell(row=r, column=3, value=1)
    s1.fill = FILL_INPUT
    s1.font = F_INPUT
    s1.border = BOX
    define(wb, "cfgHandoutWeek", "WritingHandout", f"$C${r}")
    r += 2
    ws.cell(row=r, column=2, value=(
        '="TYLER POLICE ACADEMY — "&cfgAcademyClass&" — WRITING ASSIGNMENTS '
        '— WEEK "&cfgHandoutWeek')).font = F_KPI
    r += 1
    c = ws.cell(row=r, column=2, value=DW.REQUIREMENTS_NOTE)
    c.font = F_SMALL
    c.alignment = A_LEFT_WRAP
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 28
    r += 2
    header_row(ws, ["#", "Title", "Assigned", "Due (1700)", "Prompt"], row=r)
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(FILTER(HSTACK(rngWMnum,rngWMtitle,TEXT(rngWMassigned,'
        '"mm/dd"),TEXT(rngWMdue,"mm/dd"),rngWMprompt),'
        'IFERROR(XLOOKUP(rngWMassigned,nrCDdate,nrCDweek),0)=cfgHandoutWeek),'
        '"— no assignments start this week —")'))
    for rr in range(r, r + 8):
        ws.row_dimensions[rr].height = 60
        for ccol in range(2, 9):
            ws.cell(row=rr, column=ccol).alignment = A_LEFT_WRAP
    col_widths(ws, {"A": 3, "B": 5, "C": 32, "D": 10, "E": 10, "F": 80})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW+2}:F{r+8}")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")
    sheet_note(ws, "Pick an academy week; assignments whose computed assigned "
                   "date falls in that week appear with their prompts.")
    return ws


# --------------------------------------------------------------------------
def build_emailpreview(wb):
    ws = wb.create_sheet("EmailPreview")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Agency:").font = F_LABEL
    sel = ws.cell(row=r, column=3, value="TPD")
    sel.fill = FILL_INPUT
    sel.font = F_INPUT
    sel.border = BOX
    dv_list(ws, "=rngAgencyIDs", [f"C{r}"])
    define(wb, "cfgPreviewAgency", "EmailPreview", f"$C${r}")
    ws.cell(row=r, column=5, value=(
        '="Exam #"&cfgCurrentExamNum&" — "&IFERROR(INDEX(rngEPname,'
        'MATCH(cfgCurrentExamNum,IFERROR(rngEPseq+0,-1),0)),"?")&'
        '" | Spelling through #"&cfgCurrentSpellingNum&" | Last emailed: "&'
        'IFERROR(TEXT(MAXIFS(nrELdate,nrELagency,cfgPreviewAgency),'
        '"mm/dd/yyyy"),"never")')).font = F_SMALL
    r += 2
    section_bar(ws, r, 2, 12, "Cadet results (as the email will report them)")
    r += 1
    header_row(ws, ["Cadet", "Score", "Class Avg", "Retake?", "Spelling Avg",
                    "Spelling Flag", "Attendance", "Writing Current",
                    "Open Flags"], row=r)
    r += 1
    grid_first = r
    for i in range(CADETS):
        rr = grid_first + i
        src = FIRST + i
        pred = (f'IF(OR(Cadets!$B{src}="",Cadets!$I{src}<>"Active"),FALSE,'
                f'OR(cfgPreviewAgency=cfgHomeAgency,'
                f'Cadets!$G{src}=cfgPreviewAgency))')
        ws.cell(row=rr, column=2, value=(
            f'=IF({pred},Cadets!$F{src},"")')).font = F_CALC
        ws.cell(row=rr, column=3, value=(
            f'=IF($B{rr}="","",IF(COUNTIFS(nrES_PID,Cadets!$B{src},'
            f'nrES_Seq,cfgCurrentExamNum,nrES_Final,"Yes",nrES_Rec,">=0")=0,'
            f'"",SUMIFS(nrES_Rec,nrES_PID,Cadets!$B{src},'
            f'nrES_Seq,cfgCurrentExamNum,nrES_Final,"Yes")))')).font = F_CALC
        ws.cell(row=rr, column=4, value=(
            f'=IF($B{rr}="","",IFERROR(ROUND(AVERAGEIFS(nrES_Rec,'
            f'nrES_Seq,cfgCurrentExamNum,nrES_Final,"Yes"),1),""))')).font = F_CALC
        ws.cell(row=rr, column=5, value=(
            f'=IF($B{rr}="","",IF(COUNTIFS(nrES_PID,Cadets!$B{src},'
            f'nrES_Seq,cfgCurrentExamNum,nrES_Att,2)>0,"RETEST (cap 70)",""))'
        )).font = F_CALC
        ws.cell(row=rr, column=6, value=(
            f'=IF($B{rr}="","",IF(cfgCurrentSpellingNum<cfgCurrentExamNum,'
            f'"(omitted)",Spelling!$P{src}))')).font = F_CALC
        ws.cell(row=rr, column=7, value=(
            f'=IF($B{rr}="","",IF(Spelling!$R{src}="INTERVENTION",'
            f'"BELOW 75 — intervention",""))')).font = F_CALC
        ws.cell(row=rr, column=8, value=(
            f'=IF($B{rr}="","",sysAttendance!$J{src}&" / PT "&'
            f'sysAttendance!$R{src})')).font = F_CALC
        ws.cell(row=rr, column=9, value=(
            f'=IF($B{rr}="","",Writing!$AT{src}&" ("&Writing!$AS{src}&'
            f'" overdue)")')).font = F_CALC
        ws.cell(row=rr, column=10, value=(
            f'=IF($B{rr}="","",sysFlags!$S{src})')).font = F_CALC
    grid_last = grid_first + CADETS - 1
    define(wb, "nrEPVgrid", "EmailPreview",
           f"$B${grid_first}:$J${grid_last}")
    r = grid_last + 2
    section_bar(ws, r, 2, 12,
                "Marked-for-reporting items since last email to this agency "
                "(incidents / counseling / memos you flagged Yes)")
    r += 1
    define(wb, "nrEPVsinceRow", "EmailPreview", f"$B${r+1}")
    header_row(ws, ["Date", "Cadet", "Type", "Severity/Kind", "Description"],
               row=r)
    r += 1
    # INT() + >= below: EmailLog!B is stamped with Now (date AND time) while
    # every log dates its rows date-only, so a plain > against the raw
    # timestamp dropped anything dated on a previous run's calendar day —
    # permanently, from this preview and from the draft. The VBA sender
    # (modAgencyEmail.LastEmailDate) applies the identical Int()/>= rule;
    # these two must stay in lockstep or the preview stops matching the draft.
    since = ('LET(cutoff,INT(IFERROR(MAXIFS(nrELdate,nrELagency,'
             'cfgPreviewAgency),0)),')
    # only rows YOU marked for agency reporting are included (Incidents
    # "Report to Agency?", Counseling "Agency Notified?", Memos "Report to
    # Agency?") — everything else stays an academy teaching moment
    ws.cell(row=r, column=2, value=(
        '=IFERROR(' + since +
        'inc,FILTER(HSTACK(nrIN_Date,IFERROR(XLOOKUP(nrIN_PID,rngCadetPIDs,'
        'rngCadetNames),""),IF(SEQUENCE(ROWS(nrIN_Date)),"Incident"),'
        'nrIN_Sev,nrIN_Desc),(nrIN_Date>=cutoff)*(nrIN_Report="Yes")*'
        'IFERROR((XLOOKUP(nrIN_PID,rngCadetPIDs,nrCadetAgencyID)='
        'cfgPreviewAgency)+(cfgPreviewAgency=cfgHomeAgency),0),'
        '{"","","","",""}),'
        'cns,FILTER(HSTACK(nrCO_Date,IFERROR(XLOOKUP(nrCO_PID,rngCadetPIDs,'
        'rngCadetNames),""),IF(SEQUENCE(ROWS(nrCO_Date)),"Counseling"),'
        'nrCO_Type,nrCO_Desc),(nrCO_Date>=cutoff)*(nrCO_Report="Yes")*'
        'IFERROR((XLOOKUP(nrCO_PID,rngCadetPIDs,nrCadetAgencyID)='
        'cfgPreviewAgency)+(cfgPreviewAgency=cfgHomeAgency),0),'
        '{"","","","",""}),'
        'mem,FILTER(HSTACK(nrME_Assigned,IFERROR(XLOOKUP(nrME_PID,'
        'rngCadetPIDs,rngCadetNames),""),IF(SEQUENCE(ROWS(nrME_Assigned)),'
        '"Memo"),nrME_Status,nrME_Subject),(nrME_Assigned>=cutoff)*'
        '(nrME_Report="Yes")*IFERROR((XLOOKUP(nrME_PID,rngCadetPIDs,'
        'nrCadetAgencyID)=cfgPreviewAgency)+(cfgPreviewAgency='
        'cfgHomeAgency),0),{"","","","",""}),'
        'all,VSTACK(inc,cns,mem),'
        'SORT(FILTER(all,INDEX(all,0,1)<>"",'
        '"— nothing marked for this agency since last email —"),1)),'
        '"— nothing marked for this agency since last email —")'))
    r += 16
    col_widths(ws, {"A": 3, "B": 12, "C": 24, "D": 11, "E": 11, "F": 16,
                    "G": 12, "H": 20, "I": 18, "J": 60})
    sheet_note(ws, "Preview of what the Outlook draft will contain — the "
                   "same day-granular 'since last email' cutoff the sender "
                   "uses. The buttons on THIS sheet and on the Dashboard "
                   "build the drafts — for review, never auto-sent.")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")
    return ws


# --------------------------------------------------------------------------
def build_emaillog(wb):
    ws = wb.create_sheet("EmailLog")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Date", "AgencyID", "Exam #", "Spelling #", "Cadets",
                    "Included Discipline Since", "Drafted By", "Notes"])
    first, last = DATA_ROW, DATA_ROW + 499
    fill_rows(ws, first, last, {c: (None, "in") for c in "BCDEFGHI"})
    for r in range(first, last + 1):
        ws[f"B{r}"].number_format = DATE
    define(wb, "nrELdate", "EmailLog", f"$B${first}:$B${last}")
    define(wb, "nrELagency", "EmailLog", f"$C${first}:$C${last}")
    col_widths(ws, {"A": 3, "B": 12, "C": 10, "D": 8, "E": 10, "F": 9,
                    "G": 22, "H": 16, "I": 30})
    sheet_note(ws, "Appended automatically by the email macro (one row per "
                   "agency draft). 'Last Email Sent' on Agencies and the "
                   "since-last-email digest read from here.")
    return ws


# --------------------------------------------------------------------------
def build_printcenter(wb):
    ws = wb.create_sheet("PrintCenter")
    ws.sheet_view.showGridLines = False
    r = HDR_ROW
    ws.cell(row=r, column=2, value="Print Center").font = F_KPI
    r += 2
    rows = [
        ("Daily report & roster (one day)", "Set the date on SignIn; blank "
         "date prints a generic form the leader dates by hand", "SignIn",
         "btnPrintSignIn"),
        ("Daily report & roster (week)", "Prints the next 5 class days, "
         "each pre-filled with its own date, day # and schedule", "SignIn",
         "btnPrintSignInWeek"),
        ("Academy book (all days)", "Every class day's form in one run — "
         "the class leader's bound book for the whole academy (print to "
         "PDF for the print shop)", "SignIn", "btnPrintSignInAcademy"),
        ("Course critique stack", "Set chapter on EvalSheet; prints one per "
         "active cadet", "EvalSheet", "btnPrintEvals"),
        ("Spelling test / key", "Set test # and mode on SpellingPrint",
         "SpellingPrint", "btnPrintSpelling"),
        ("Writing handout (week)", "Set week # on WritingHandout",
         "WritingHandout", "btnPrintWriting"),
        ("Cadet profile", "Set cadet on CadetProfile", "CadetProfile",
         "btnPrintProfile"),
        ("Transcript", "Set cadet on Transcript (button can run all cadets)",
         "Transcript", "btnPrintTranscript"),
        ("Class ranking", "Prints as-is", "Ranking", "btnPrintRanking"),
        ("Graduation checklist", "Prints as-is", "GradChecklist",
         "btnPrintGradCheck"),
        ("Audit packet", "Program checks + enrollment docs", "Audit",
         "btnPrintAudit"),
        ("Chapter packet", "Set chapter on ChapterPacket — the full "
         "training-file page for the auditor", "ChapterPacket",
         "btnPrintChapterPacket"),
        ("Exam grade sheet", "Set exam # on ExamSheet — the IRG-required "
         "grade sheet per assessment", "ExamSheet", "btnPrintGradeSheet"),
        ("Addendum (excess hours)", "Per-class excess vs TCOLE minimum with "
         "reporting course #s (#101 / #2040 / #2046 / #2055)", "Addendum",
         "btnPrintAddendum"),
        ("Schedule", "Full schedule listing", "Schedule", "btnPrintSchedule"),
    ]
    header_row(ws, ["What", "How", "Sheet", "Macro (button)"], row=r)
    r += 1
    for what, how, sheet, macro in rows:
        ws.cell(row=r, column=2, value=what).font = F_LABEL
        ws.cell(row=r, column=3, value=how).font = F_BODY
        ws.cell(row=r, column=4, value=sheet).font = F_BODY
        ws.cell(row=r, column=5, value=macro).font = F_SMALL
        ws.row_dimensions[r].height = 20
        r += 1
    r += 1
    c = ws.cell(row=r, column=2, value=(
        "Buttons are added by the BPOC VBA install — run "
        "tools/Install-BPOC-VBA.ps1 (see docs/BPOC-Coordinator-Guide.md, "
        "'Rebuilding from source'). Each button prints its sheet's defined "
        "print area to the default printer; Ctrl+P on the sheet does the "
        "same thing manually."))
    c.font = F_SMALL
    c.alignment = A_LEFT_WRAP
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 30
    # F blank gutter; G+H carry the install script's button strip, one
    # button per table row (rows 8..22) — they need real width or the
    # captions are clipped to a few characters
    col_widths(ws, {"A": 3, "B": 26, "C": 52, "D": 16, "E": 20,
                    "F": 3, "G": 20, "H": 20})
    return ws


# --------------------------------------------------------------------------
INPUT_GUIDE = [
    ("Getting set up", None),
    ("Settings", "Academy label, start/end dates, grading weights, caps, "
     "retest window, flag thresholds, PT points rubric (yellow block)."),
    ("Lists", "Dropdown choices used across the workbook — add items here."),
    ("Agencies", "Sending agencies: codes, contacts, email addresses."),
    ("Instructors", "Roster with TCOLE PID, certs, SME letters, bio-on-file "
     "status. 'On Schedule?' shows who actually teaches — every co-teacher "
     "needs documentation."),
    ("InstructorBanks", "Per-topic banks: the certified pool (kept between "
     "academies) and who you picked to teach THIS academy — those picks "
     "become the Schedule's instructor dropdown for that topic."),
    ("ChapterMaster", "Chapter hours + the per-chapter TCOLE training-file "
     "record (lesson plan, bio, sign-ins, assessment, grade sheet, eval, "
     "special requirements)."),
    ("Control", "Extra closure dates (holidays compute automatically)."),
    ("Schedule", "The class schedule: one row per time block — date, times, "
     "chapter/activity, instructor. Drives hours, due dates, sign-ins."),
    ("ExamPlan", "Which exams this academy uses, their sequence and linked "
     "chapters."),
    ("ExamMaster", "The exam library (edit names/types/defaults)."),
    ("SkillsMaster", "Attempt limits and scoring mode per skill category."),
    ("SpellingMaster", "The 12 spelling word lists (printable via Print "
     "Center)."),
    ("WritingMaster", "The 40 assignments: prompts, linked chapters, "
     "date overrides."),
    ("Daily operation", None),
    ("Cadets", "The roster: PID, names, agency, status, enroll/separation."),
    ("ExamScores", "One row per exam attempt: cadet, exam, attempt #, raw "
     "score, date. Retest deadlines compute."),
    ("Spelling", "Spelling test scores per cadet per test."),
    ("Attendance", "Exception log: missed/modified time in minutes (or PT "
     "sessions), reason, documentation, excused."),
    ("Makeup", "Makeup credit: minutes/sessions, documentation, holds."),
    ("Skills", "Skills attempts: category, result, score, firearms course "
     "of fire."),
    ("Writing", "Type X when an assignment is received (auto-capitalizes)."),
    ("Incidents", "Positive/negative incidents with severity and resolution."),
    ("Counseling", "Every intervention: tutoring, counseling, agency "
     "notification, performance plans. 'Agency Notified?' = Yes puts it in "
     "the next email digest."),
    ("Memos", "Deficiency memos: assign, link to the I/A/C record, due "
     "auto-computes, mark received. 'Report to Agency?' is your call."),
    ("DailyLog", "One row per training day — the digital daily report. "
     "Counters compute; mark when the leader's signed report is scanned."),
    ("PT", "Baseline and final raw values per event; final points once the "
     "rubric arrives."),
    ("Medical", "Injuries, restrictions, clearances and expirations."),
    ("Certifications", "Per-cadet TIM/SFST/TCIC/CPR/ALERRT/ICS completion "
     "dates and certificate copies collected."),
    ("StateExam", "TCOLE licensing exam attempts and results (3 max)."),
    ("DismissalLog", "Formal reviews: trigger, decision, approvals."),
    ("AdvisoryBoard", "Governance: this academy's alignment record "
     "(policy version, minutes reviewed, workbook aligned) + the running "
     "board-meeting list with minutes folder locations. The Startup Review "
     "button fills it."),
    ("End of academy / as needed", None),
    ("Audit", "Program-requirement checklist answers + per-cadet enrollment "
     "documents grid."),
    ("EmailPreview", "Pick an agency to preview its email; the button "
     "builds Outlook drafts."),
    ("sysAwards", "Award override cells — your pick always wins."),
    ("PrintCenter", "Every printable in one place (buttons after VBA "
     "install)."),
]


def build_inputguide(wb):
    ws = wb.create_sheet("InputGuide")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Go to…", "What you enter there"])
    r = DATA_ROW
    for name, desc in INPUT_GUIDE:
        if desc is None:
            section_bar(ws, r, 2, 8, name)
            r += 1
            continue
        c = ws.cell(row=r, column=2,
                    value=f'=HYPERLINK("#{name}!B6","{name}")')
        c.font = F_LABEL
        d = ws.cell(row=r, column=3, value=desc)
        d.font = F_BODY
        d.alignment = A_LEFT_WRAP
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 26
        r += 1
    col_widths(ws, {"A": 3, "B": 18, "C": 90})
    sheet_note(ws, "Only pages that take input are listed — gray sys* tabs "
                   "are the locked calculation engine; green tabs are "
                   "outputs. White boxes with a blue border are yours to type "
                   "(what you type shows in blue); gray cells calculate.")
    return ws


def add_home_links(wb):
    """'◄ Dashboard' hyperlink in B1 of every visible sheet."""
    from openpyxl.styles import Font as _Font
    link_font = _Font(name="Arial", size=9, bold=True, color="1F3B5C",
                      underline="single")
    for ws in wb.worksheets:
        if ws.title in ("Dashboard", "sysListsHelper"):
            continue
        c = ws["B1"]
        c.value = '=HYPERLINK("#Dashboard!B5","◄ Dashboard")'
        c.font = link_font
        from openpyxl.styles import Protection as _Prot
        c.protection = _Prot(locked=False)


def gray_separated_rows(wb):
    """Gray out (and strike) rows of non-active cadets on cadet grids."""
    from openpyxl.styles import Font as _Font, PatternFill as _Fill
    from openpyxl.formatting.rule import FormulaRule as _FR
    gray_font = _Font(color="9AA5B1", strike=True)
    gray_fill = _Fill("solid", fgColor="EDEFF2")
    targets = {
        "Writing": "B6:AT55", "Spelling": "B6:R55", "PT": "B6:AC55",
        "Certifications": "B6:U55", "ScoresGrid": "B6:AC55",
        "GradChecklist": "B6:P55", "StateExam": "B6:L55",
        "Cadets": "B6:M55",
    }
    for name, rng in targets.items():
        ws = wb[name]
        ws.conditional_formatting.add(rng, _FR(
            formula=['AND(Cadets!$B6<>"",Cadets!$I6<>"Active")'],
            font=gray_font, fill=gray_fill))


def build_namedranges(wb):
    ws = wb.create_sheet("NamedRanges")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Name", "Refers To"])
    r = DATA_ROW
    for name, dn in sorted(wb.defined_names.items()):
        ws.cell(row=r, column=2, value=name).font = F_BODY
        ws.cell(row=r, column=3, value=dn.attr_text).font = F_SMALL
        r += 1
    col_widths(ws, {"A": 3, "B": 30, "C": 46})
    sheet_note(ws, "Auto-generated registry of every named range at build "
                   "time. Do not rename ranges without updating formulas.")
    protect(ws)
    return ws


def build_all_outputs(wb):
    build_scoresgrid(wb)
    build_ranking(wb)
    build_watchlist(wb)
    build_cadetprofile(wb)
    build_transcript(wb)
    build_gradchecklist(wb)
    build_dismissallog(wb)
    build_audit(wb)
    build_addendum(wb)
    build_chapterpacket(wb)
    build_examsheet(wb)
    build_signin(wb)
    build_evalsheet(wb)
    build_spellingprint(wb)
    build_writinghandout(wb)
    build_emailpreview(wb)
    build_emaillog(wb)
    build_printcenter(wb)
    build_inputguide(wb)
    build_dashboard(wb)     # after ScoresGrid/Spelling exist (charts)
    gray_separated_rows(wb)
    # NamedRanges BEFORE add_home_links: created after, it was the only
    # visible sheet with no '◄ Dashboard' link — and it is protected with
    # nothing unlocked, so there was no way back off it at all.
    build_namedranges(wb)   # captures every name defined above
    add_home_links(wb)

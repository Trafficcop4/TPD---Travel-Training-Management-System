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
    r += 2
    _kpi(ws, r, 2, "Active cadets",
         'SUMPRODUCT((nrCadetStatus="Active")*1)')
    _kpi(ws, r, 4, "Class current avg",
         'IFERROR(ROUND(AVERAGE(IF(nrCadetStatus="Active",nrGRcurrent)),1),"—")')
    _kpi(ws, r, 6, "Graduation eligible",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrCKgradElig="Yes"))')
    _kpi(ws, r, 8, "Flagged cadets",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrFLcount>0))')
    _kpi(ws, r, 10, "Overdue retests",
         'COUNTIF(nrES_RetStat,"OVERDUE")')
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
    r += 3
    section_bar(ws, r, 2, 11, "Watch list — highest flag counts first "
                              "(full list on WatchList)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(SORTBY(FILTER(HSTACK(rngCadetNames,nrCadetAgency,nrFLcount,'
        'nrFLreasons),(nrFLcount>0)*(nrCadetStatus="Active")),'
        'FILTER(nrFLcount,(nrFLcount>0)*(nrCadetStatus="Active")),-1),'
        '"No flags — clear")'))
    watch_top = r
    r += 11
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
    ch.add_data(data, titles_from_data=False)
    ch.set_categories(cats)
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
    ch2.add_data(data2, titles_from_data=False)
    ch2.set_categories(cats2)
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
        cols[cl] = (
            'IF(OR($B{r}="",%s$%d=""),"",IFERROR(SUMIFS(nrES_Rec,nrES_PID,'
            '$B{r},nrES_Code,%s$%d,nrES_Final,"Yes"),""))'
            % (cl, HDR_ROW, cl, HDR_ROW), "fx")
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
        ws[f"{cl}{ar}"] = (f'=IF({cl}{HDR_ROW}="","",IFERROR(ROUND(AVERAGEIF('
                           f'{cl}{FIRST}:{cl}{LAST},"<>"),1),""))')
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
    for lab2, rng in (("Major avg", "nrGRmajavg"),
                      ("Minor avg", f"sysGrades!$J${FIRST}:$J${LAST}"),
                      ("Spelling avg", f"sysGrades!$K${FIRST}:$K${LAST}"),
                      ("Final", "nrGRfinal")):
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
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(nrIN_Date,nrIN_Dir,nrIN_Sev,'
        'nrIN_Desc),nrIN_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,'
        'rngCadetPIDs,"")),FILTER(nrIN_Date,nrIN_PID=XLOOKUP(cfgProfileCadet,'
        'rngCadetNames,rngCadetPIDs,"")),-1),10),"none")'))
    r += 11
    ws.cell(row=r, column=2, value=(
        '=IFERROR(TAKE(SORTBY(FILTER(HSTACK(nrCO_Date,nrCO_Type,nrCO_Desc),'
        'nrCO_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,rngCadetPIDs,"")),'
        'FILTER(nrCO_Date,nrCO_PID=XLOOKUP(cfgProfileCadet,rngCadetNames,'
        'rngCadetPIDs,"")),-1),10),"none")'))
    r += 11
    col_widths(ws, {"A": 3, "B": 24, "C": 14, "D": 14, "E": 16, "F": 14,
                    "G": 14, "H": 18, "I": 14, "J": 14})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:J{r}")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:E{HDR_ROW}")
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
    for lab2, rng, w in (("Major (40%)", "nrGRmajavg", None),
                         ("Minor (30%)", f"sysGrades!$J${FIRST}:$J${LAST}", None),
                         ("Spelling (10%)", f"sysGrades!$K${FIRST}:$K${LAST}", None),
                         ("Final Exam (20%)", "nrGRfinal", None)):
        _profile_label(ws, r, 2, lab2, P + rng + ',"—")')
        r += 1
    r += 1
    section_bar(ws, r, 2, 10, "Exam record (recorded scores, final attempts)")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(FILTER(HSTACK(ExamScores!$G$%d:$G$%d,ExamScores!$H$%d:$H$%d,'
        'ExamScores!$K$%d:$K$%d,ExamScores!$M$%d:$M$%d,ExamScores!$S$%d:$S$%d),'
        '(nrES_PID=%srngCadetPIDs,""))*(nrES_Final="Yes")*(nrES_Rec<>"")),"—")'
        % (FIRST, FIRST + 1499, FIRST, FIRST + 1499, FIRST, FIRST + 1499,
           FIRST, FIRST + 1499, FIRST, FIRST + 1499, P)))
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
    unlock_range(ws, f"C{HDR_ROW}:E{HDR_ROW}")
    sheet_note(ws, "The end-of-academy deliverable per cadet — print one per "
                   "cadet for the agency packet.")
    return ws


# --------------------------------------------------------------------------
def build_gradchecklist(wb):
    ws = wb.create_sheet("GradChecklist")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Cadet", "Agency", "Academic", "Classroom",
                    "PT Sessions", "Skills", "Incidents", "Writing",
                    "Makeup", "Final PT", "No Dismiss Rev", "ELIGIBLE",
                    "Blocking Issues"])
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
        "N": ('IF($B{r}="","",sysChecks!$N{r})', "fx"),
        "O": ('IF($B{r}="","",sysChecks!$O{r})', "fx"),
    }
    fill_rows(ws, FIRST, LAST, cols)
    cf_yes_no(ws, f"E{FIRST}:N{LAST}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24, "D": 18, "O": 44})
    for cl in "EFGHIJKLMN":
        ws.column_dimensions[cl].width = 10
    page_setup_landscape(ws, print_area=f"B{HDR_ROW}:O{LAST}",
                         repeat_rows=f"{HDR_ROW}:{HDR_ROW}")
    sheet_note(ws, "The final gate before the ceremony — every column must "
                   "be Yes.")
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
    n_checks = 11
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
        "70-in-each-category rule, weekly spelling tests + 75% intervention, "
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
                    "Psych (L3)", "Background", "Photo ID/DL", "All Docs?"],
               row=r)
    r += 1
    doc_first = r
    doc_last = r + CADETS - 1
    cols = {
        "B": ('IF(Cadets!$B{rr}="","",Cadets!$F{rr})', "fx"),
        "I": ('IF($B{r}="","",IF(COUNTIF(C{r}:H{r},"Yes")=6,"Yes","No"))', "fx"),
    }
    for rr in range(doc_first, doc_last + 1):
        src = FIRST + (rr - doc_first)
        ws.cell(row=rr, column=2,
                value=f'=IF(Cadets!$B{src}="","",Cadets!$F{src})').font = F_CALC
        for c in range(3, 9):
            cc = ws.cell(row=rr, column=c)
            cc.fill = FILL_INPUT
            cc.font = F_INPUT
            cc.border = BOX
        ws.cell(row=rr, column=9, value=(
            f'=IF($B{rr}="","",IF(COUNTIF(C{rr}:H{rr},"Yes")=6,"Yes","No"))'
        )).font = F_CALC
    dv_list(ws, "=lstYesNo", [f"C{doc_first}:H{doc_last}"])
    cf_yes_no(ws, f"I{doc_first}:I{doc_last}")
    col_widths(ws, {"A": 3, "B": 30, "C": 11, "D": 11, "E": 12, "F": 11,
                    "G": 12, "H": 12, "I": 10})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:I{doc_last}",
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
        '="TYLER POLICE ACADEMY — "&cfgAcademyClass&" — DAILY ATTENDANCE ROSTER — "'
        '&TEXT(cfgSignInDate,"dddd, mmmm d, yyyy")')).font = F_KPI
    r += 2
    section_bar(ws, r, 2, 9, "Instruction scheduled this date")
    r += 1
    ws.cell(row=r, column=2, value=(
        '=IFERROR(FILTER(HSTACK(TEXT(nrSCH_Start,"h:mm AM/PM"),'
        'TEXT(nrSCH_End,"h:mm AM/PM"),nrSCH_Act,nrSCH_Instr,nrSCH_Loc),'
        'nrSCH_Date=cfgSignInDate),"— no schedule entered for this date —")'))
    r += 9
    section_bar(ws, r, 2, 9, "Cadet sign-in")
    r += 1
    header_row(ws, ["#", "Cadet", "PID", "Agency", "AM Signature",
                    "PM Signature", "Remarks"], row=r)
    r += 1
    first = r
    for i in range(CADETS):
        rr = first + i
        src = FIRST + i
        ws.cell(row=rr, column=2, value=(
            f'=IF(Cadets!$B{src}="","",ROW()-{first-1})')).font = F_CALC
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
    r = last + 2
    ws.cell(row=r, column=2, value="Instructor verification: ____________________"
            "____________     Time: ____________").font = F_BODY
    r += 1
    ws.cell(row=r, column=2, value=DL.ACADEMY_ADDRESS).font = F_SMALL
    col_widths(ws, {"A": 3, "B": 6, "C": 28, "D": 10, "E": 18, "F": 21,
                    "G": 21, "H": 22, "I": 12})
    page_setup_portrait(ws, print_area=f"B{HDR_ROW}:H{r}")
    protect(ws)
    unlock_range(ws, f"C{HDR_ROW}:C{HDR_ROW}")
    sheet_note(ws, "Enter the date, print, collect signatures — this is the "
                   "daily paper attendance record for the TCOLE file. "
                   "Exceptions still go on the Attendance log.")
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
    unlock_range(ws, f"C{HDR_ROW}:E{HDR_ROW}")
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
            f'=IF($B{rr}="","",IFERROR(SUMIFS(nrES_Rec,nrES_PID,Cadets!$B{src},'
            f'nrES_Seq,cfgCurrentExamNum,nrES_Final,"Yes"),""))')).font = F_CALC
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
            f'=IF($B{rr}="","",IF(nrFLcount=0,"",'
            f'sysFlags!$P{src}))')).font = F_CALC
    grid_last = grid_first + CADETS - 1
    # fix flag column formula (row-wise)
    for i in range(CADETS):
        rr = grid_first + i
        src = FIRST + i
        ws.cell(row=rr, column=10, value=(
            f'=IF($B{rr}="","",sysFlags!$P{src})'))
    define(wb, "nrEPVgrid", "EmailPreview",
           f"$B${grid_first}:$J${grid_last}")
    r = grid_last + 2
    section_bar(ws, r, 2, 12,
                "Discipline & counseling since last email to this agency")
    r += 1
    define(wb, "nrEPVsinceRow", "EmailPreview", f"$B${r+1}")
    header_row(ws, ["Date", "Cadet", "Type", "Severity/Kind", "Description"],
               row=r)
    r += 1
    since = ('LET(cutoff,IFERROR(MAXIFS(nrELdate,nrELagency,'
             'cfgPreviewAgency),0),')
    ws.cell(row=r, column=2, value=(
        '=IFERROR(' + since +
        'inc,FILTER(HSTACK(nrIN_Date,IFERROR(XLOOKUP(nrIN_PID,rngCadetPIDs,'
        'rngCadetNames),""),IF(SEQUENCE(ROWS(nrIN_Date)),"Incident"),'
        'nrIN_Sev,nrIN_Desc),(nrIN_Date>cutoff)*(nrIN_Dir="Negative")*'
        'IFERROR((XLOOKUP(nrIN_PID,rngCadetPIDs,nrCadetAgencyID)='
        'cfgPreviewAgency)+(cfgPreviewAgency=cfgHomeAgency),0)),'
        'cns,FILTER(HSTACK(nrCO_Date,IFERROR(XLOOKUP(nrCO_PID,rngCadetPIDs,'
        'rngCadetNames),""),IF(SEQUENCE(ROWS(nrCO_Date)),"Counseling"),'
        'nrCO_Type,nrCO_Desc),(nrCO_Date>cutoff)*'
        'IFERROR((XLOOKUP(nrCO_PID,rngCadetPIDs,nrCadetAgencyID)='
        'cfgPreviewAgency)+(cfgPreviewAgency=cfgHomeAgency),0)),'
        'SORT(VSTACK(inc,cns),1)),"— none since last email —")'))
    r += 16
    col_widths(ws, {"A": 3, "B": 12, "C": 24, "D": 11, "E": 11, "F": 16,
                    "G": 12, "H": 20, "I": 18, "J": 60})
    sheet_note(ws, "Preview of exactly what the Outlook draft will contain. "
                   "Buttons on Print Center / Dashboard build the drafts — "
                   "for review, never auto-sent.")
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
        ("Daily sign-in sheet", "Set the date on SignIn (or here)", "SignIn",
         "btnPrintSignIn"),
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
        "Buttons are added by the VBA install (see docs/Setup). Each button "
        "prints its sheet's defined print area to the default printer; "
        "Ctrl+P on the sheet does the same thing manually."))
    c.font = F_SMALL
    c.alignment = A_LEFT_WRAP
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 30
    col_widths(ws, {"A": 3, "B": 26, "C": 52, "D": 16, "E": 20})
    return ws


# --------------------------------------------------------------------------
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
    build_signin(wb)
    build_evalsheet(wb)
    build_spellingprint(wb)
    build_writinghandout(wb)
    build_emailpreview(wb)
    build_emaillog(wb)
    build_printcenter(wb)
    build_dashboard(wb)     # after ScoresGrid/Spelling exist (charts)
    build_namedranges(wb)   # last — captures all names

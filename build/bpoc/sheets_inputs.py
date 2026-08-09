"""Input sheets: Cadets, ExamScores, Spelling, Attendance, Makeup, Skills,
Writing, Incidents, Counseling, PT, Medical.

All keyed by PID; cadet grids run rows 6..55 (50 cadets); logs are long
tables. Blue cells = staff entry, gray = calculated.
"""
from openpyxl.utils import get_column_letter

from xlb import (
    HDR_ROW, DATA_ROW, CADETS, CADET_LAST, ROWS_EXAMSCORES, ROWS_ATTEND,
    ROWS_MAKEUP, ROWS_SKILLS, ROWS_INCIDENTS, ROWS_COUNSELING, ROWS_MEDICAL,
    F_HDR, F_CALC, FILL_CALC, F_LABEL, F_SMALL, F_INPUT, FILL_INPUT,
    FILL_YELLOW, A_LEFT_WRAP, A_CENTER, BOX, DATE, TIME,
    header_row, fill_rows, dv_list, sheet_note, cf_yes_no, cf_formula,
    FILL_WARNBG, FILL_OKBG, FILL_AMBER, col_widths, define,
)
import data_lists as DL
import data_writing as DW


# --------------------------------------------------------------------------
def build_cadets(wb):
    ws = wb.create_sheet("Cadets")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Last Name", "First Name", "MI", "Cadet Name",
                    "AgencyID", "Agency Name", "Status", "Sort",
                    "Enroll Date", "Separation Date", "Separation Reason"])
    first, last = DATA_ROW, CADET_LAST
    fill_rows(ws, first, last, {
        "B": (None, "in"), "C": (None, "in"), "D": (None, "in"),
        "E": (None, "in"),
        "F": ('IF($B{r}="","",$C{r}&", "&$D{r}&IF($E{r}="",""," "&$E{r}&"."))', "fx"),
        "G": (None, "in"),
        "H": ('IF($G{r}="","",IFERROR(INDEX(rngAgencyNames,MATCH($G{r},rngAgencyIDs,0)),"?"))', "fx"),
        "I": (None, "in"),
        "J": ('IF($B{r}="","",ROW()-{})'.replace("{}", str(HDR_ROW)), "fx"),
        "K": (None, "in"), "L": (None, "in"), "M": (None, "in"),
    })
    for r in range(first, last + 1):
        ws[f"K{r}"].number_format = DATE
        ws[f"L{r}"].number_format = DATE
    dv_list(ws, "=rngAgencyIDs", [f"G{first}:G{last}"])
    dv_list(ws, "=lstCadetStatus", [f"I{first}:I{last}"])
    define(wb, "rngCadetPIDs", "Cadets", f"$B${first}:$B${last}")
    define(wb, "rngCadetNames", "Cadets", f"$F${first}:$F${last}")
    define(wb, "nrCadetAgency", "Cadets", f"$H${first}:$H${last}")
    define(wb, "nrCadetAgencyID", "Cadets", f"$G${first}:$G${last}")
    define(wb, "nrCadetStatus", "Cadets", f"$I${first}:$I${last}")
    col_widths(ws, {"A": 3, "B": 10, "C": 14, "D": 14, "E": 5, "F": 24,
                    "G": 10, "H": 22, "I": 11, "J": 6, "K": 12, "L": 12,
                    "M": 26})
    sheet_note(ws, "PID never changes once set — every other sheet keys off "
                   "it. Capacity 50 cadets.")
    return ws


# --------------------------------------------------------------------------
def build_examscores(wb):
    ws = wb.create_sheet("ExamScores")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["RecordID", "Cadet Name", "PID", "Agency", "ExamCode",
                    "Exam Name", "Type", "Seq", "Passing", "Attempt #",
                    "Raw Score", "Recorded", "AttemptPass", "Pass?",
                    "FinalAttempt?", "RetakeReq?", "DismissReview?", "Date",
                    "Retest Due By", "Retest Status", "Entered By", "Notes"])
    first, last = DATA_ROW, DATA_ROW + ROWS_EXAMSCORES - 1
    fill_rows(ws, first, last, {
        "B": ('IF($C{r}="","",$D{r}&"-"&$F{r}&"-"&$K{r})', "fx"),
        "C": (None, "in"),
        "D": ('IF($C{r}="","",IFERROR(INDEX(rngCadetPIDs,MATCH($C{r},rngCadetNames,0)),"?"))', "fx"),
        "E": ('IF($D{r}="","",IFERROR(INDEX(nrCadetAgency,MATCH($D{r},rngCadetPIDs,0)),""))', "fx"),
        "F": (None, "in"),
        "G": ('IF($F{r}="","",IFERROR(INDEX(rngEPname,MATCH($F{r},rngEPcode,0)),"?"))', "fx"),
        "H": ('IF($F{r}="","",IFERROR(INDEX(rngEPtype,MATCH($F{r},rngEPcode,0)),"?"))', "fx"),
        "I": ('IF($F{r}="","",IFERROR(INDEX(rngEPseq,MATCH($F{r},rngEPcode,0)),""))', "fx"),
        "J": ('IF($F{r}="","",IFERROR(INDEX(rngEPpass,MATCH($F{r},rngEPcode,0)),cfgPassingScore))', "fx"),
        "K": (None, "in"), "L": (None, "in"),
        "M": ('IF($C{r}="","",IF($L{r}="","",IF($K{r}=1,'
              'IF(COUNTIFS(nrES_PID,$D{r},nrES_Code,$F{r},nrES_Att,2)>0,"",$L{r}),'
              'IF(COUNTIFS(nrES_PID,$D{r},nrES_Code,$F{r},nrES_Att,1)=0,$L{r},'
              'IF($L{r}>=$J{r},cfgRetakeRecordedCap,'
              'SUMIFS(nrES_Raw,nrES_PID,$D{r},nrES_Code,$F{r},nrES_Att,1))))))', "fx"),
        "N": ('IF($C{r}="","",IF($L{r}="","",IF($L{r}>=$J{r},"Yes","No")))', "fx"),
        "O": ('IF($C{r}="","",IF($M{r}="","",IF($M{r}>=$J{r},"Yes","No")))', "fx"),
        "P": ('IF($C{r}="","",IF($K{r}="","",IF(COUNTIFS(nrES_PID,$D{r},'
              'nrES_Code,$F{r},nrES_Att,">"&$K{r})=0,"Yes","No")))', "fx"),
        "Q": ('IF($C{r}="","",IF($L{r}="","",IF(AND($K{r}=1,$L{r}<$J{r},'
              'COUNTIFS(nrES_PID,$D{r},nrES_Code,$F{r},nrES_Att,2)=0),"Yes","No")))', "fx"),
        "R": ('IF($C{r}="","",IF($L{r}="","",IF(AND($K{r}=2,$L{r}<$J{r},'
              'COUNTIFS(nrES_PID,$D{r},nrES_Code,$F{r},nrES_Att,1)>0),"Yes","No")))', "fx"),
        "S": (None, "in"),
        "T": ('IF($Q{r}<>"Yes","",IF($S{r}="","(enter date)",'
              'LET(dn,IFERROR(XLOOKUP($S{r},nrCDdate,nrCDnum),""),'
              'IF(dn="","(not a class day)",'
              'IFERROR(XLOOKUP(dn+cfgRetestClassDays,nrCDnum,nrCDdate),"")))))', "fx"),
        "U": ('IF($Q{r}<>"Yes","",IF(COUNTIFS(nrES_PID,$D{r},nrES_Code,$F{r},'
              'nrES_Att,2)>0,"Retested",IF(OR($T{r}="",NOT(ISNUMBER($T{r}))),"Pending",'
              'IF(TODAY()>$T{r},"OVERDUE","Due "&TEXT($T{r},"mm/dd")))))', "fx"),
        "V": (None, "in"), "W": (None, "in"),
    })
    for r in range(first, last + 1):
        ws[f"S{r}"].number_format = DATE
        ws[f"T{r}"].number_format = DATE
    dv_list(ws, "=rngCadetNames", [f"C{first}:C{last}"])
    dv_list(ws, "=rngEPcode", [f"F{first}:F{last}"])
    dv_list(ws, "=lstAttemptNum", [f"K{first}:K{last}"])
    cf_formula(ws, f"U{first}:U{last}", f'$U{first}="OVERDUE"', FILL_WARNBG)
    cf_yes_no(ws, f"O{first}:O{last}")
    define(wb, "nrES_PID", "ExamScores", f"$D${first}:$D${last}")
    define(wb, "nrES_Code", "ExamScores", f"$F${first}:$F${last}")
    define(wb, "nrES_Type", "ExamScores", f"$H${first}:$H${last}")
    define(wb, "nrES_Seq", "ExamScores", f"$I${first}:$I${last}")
    define(wb, "nrES_Att", "ExamScores", f"$K${first}:$K${last}")
    define(wb, "nrES_Raw", "ExamScores", f"$L${first}:$L${last}")
    define(wb, "nrES_Rec", "ExamScores", f"$M${first}:$M${last}")
    define(wb, "nrES_Final", "ExamScores", f"$P${first}:$P${last}")
    define(wb, "nrES_RetReq", "ExamScores", f"$Q${first}:$Q${last}")
    define(wb, "nrES_Dis", "ExamScores", f"$R${first}:$R${last}")
    define(wb, "nrES_Date", "ExamScores", f"$S${first}:$S${last}")
    define(wb, "nrES_RetDue", "ExamScores", f"$T${first}:$T${last}")
    define(wb, "nrES_RetStat", "ExamScores", f"$U${first}:$U${last}")
    col_widths(ws, {"A": 3, "B": 16, "C": 22, "D": 9, "E": 18, "F": 10,
                    "G": 34, "H": 9, "I": 6, "J": 9, "K": 9, "L": 9, "M": 10,
                    "N": 11, "O": 8, "P": 12, "Q": 11, "R": 13, "S": 11,
                    "T": 12, "U": 12, "V": 14, "W": 26})
    sheet_note(ws, "One row per attempt. Attempt 2 of a failed exam records "
                   "at the 70 cap when passed (policy 300.5). Retest Due By "
                   "= 5 class days after the failed attempt's date.")
    return ws


# --------------------------------------------------------------------------
def build_spelling(wb):
    ws = wb.create_sheet("Spelling")
    ws.sheet_view.showGridLines = False
    tests = [f"S{n:02d}" for n in range(1, 13)]
    header_row(ws, ["PID", "Cadet Name"] + tests +
               ["Spelling Avg", "# Taken", "Intervention?"])
    first, last = DATA_ROW, CADET_LAST
    cols = {
        "B": (f'IF(Cadets!$B{{r}}="","",Cadets!$B{{r}})', "fx"),
        "C": (f'IF(Cadets!$B{{r}}="","",Cadets!$F{{r}})', "fx"),
        "P": ('IF($C{r}="","",IF($Q{r}=0,"",ROUND(SUM($D{r}:$O{r})/$Q{r},1)))', "fx"),
        "Q": ('IF($C{r}="","",COUNT($D{r}:$O{r}))', "fx"),
        "R": ('IF($C{r}="","",IF($P{r}="","",'
              'IF($P{r}<cfgSpellInterventionAvg,"INTERVENTION","OK")))', "fx"),
    }
    for i in range(12):
        cols[get_column_letter(4 + i)] = (None, "in")
    fill_rows(ws, first, last, cols)
    cf_formula(ws, f"R{first}:R{last}", f'$R{first}="INTERVENTION"', FILL_WARNBG)
    # class stats rows below the grid
    sr = last + 2
    ws.cell(row=sr, column=3, value="Scores entered:").font = F_LABEL
    ws.cell(row=sr + 1, column=3, value="Class average:").font = F_LABEL
    ws.cell(row=sr + 2, column=3, value="Test #:").font = F_SMALL
    for i in range(12):
        cl = get_column_letter(4 + i)
        ws[f"{cl}{sr}"] = f"=COUNT({cl}{first}:{cl}{last})"
        ws[f"{cl}{sr}"].font = F_CALC
        ws[f"{cl}{sr+1}"] = (f"=IF({cl}{sr}=0,\"\","
                             f"ROUND(AVERAGE({cl}{first}:{cl}{last}),1))")
        ws[f"{cl}{sr+1}"].font = F_CALC
        ws[f"{cl}{sr+2}"] = i + 1
        ws[f"{cl}{sr+2}"].font = F_SMALL
    define(wb, "nrSpellCounts", "Spelling", f"$D${sr}:$O${sr}")
    define(wb, "nrSpellClassAvg", "Spelling", f"$D${sr+1}:$O${sr+1}")
    define(wb, "nrSpellTestNums", "Spelling", f"$D${sr+2}:$O${sr+2}")
    define(wb, "nrSpellAvg", "Spelling", f"$P${first}:$P${last}")
    define(wb, "nrSpellTaken", "Spelling", f"$Q${first}:$Q${last}")
    define(wb, "nrSpellFlag", "Spelling", f"$R${first}:$R${last}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24, "P": 12, "Q": 9, "R": 14})
    for i in range(12):
        ws.column_dimensions[get_column_letter(4 + i)].width = 7
    sheet_note(ws, "Scores 0-100 (25 words x 4 pts). Average below the "
                   "Settings intervention threshold (default 75) flags "
                   "INTERVENTION per policy 300.4.B — see Counseling log.")
    return ws


# --------------------------------------------------------------------------
def build_attendance(wb):
    ws = wb.create_sheet("Attendance")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["EventID", "Date", "Cadet Name", "PID", "Agency",
                    "Event Type", "Reason", "Minutes", "Sessions", "Is PT?",
                    "Doc Status", "Excused?", "Counts?", "Notes"])
    first, last = DATA_ROW, DATA_ROW + ROWS_ATTEND - 1
    fill_rows(ws, first, last, {
        "B": ('IF($D{r}="","","A"&TEXT(ROW()-%d,"000"))' % HDR_ROW, "fx"),
        "C": (None, "in"), "D": (None, "in"),
        "E": ('IF($D{r}="","",IFERROR(INDEX(rngCadetPIDs,MATCH($D{r},rngCadetNames,0)),"?"))', "fx"),
        "F": ('IF($E{r}="","",IFERROR(INDEX(nrCadetAgency,MATCH($E{r},rngCadetPIDs,0)),""))', "fx"),
        "G": (None, "in"), "H": (None, "in"), "I": (None, "in"),
        "J": (None, "in"),
        "K": ('IF($G{r}="","",IF(OR($G{r}="PT Missed",$G{r}="PT Modified",'
              '$G{r}="PT Refused"),"Yes","No"))', "fx"),
        "L": (None, "in"), "M": (None, "in"),
        "N": ('IF($D{r}="","",IF($M{r}="Excused","No",'
              'IF(AND($G{r}="PT Modified",$L{r}="Received"),"No","Yes")))', "fx"),
        "O": (None, "in"),
    })
    for r in range(first, last + 1):
        ws[f"C{r}"].number_format = DATE
    dv_list(ws, "=rngCadetNames", [f"D{first}:D{last}"])
    dv_list(ws, "=lstAttendanceEvent", [f"G{first}:G{last}"])
    dv_list(ws, "=lstReason", [f"H{first}:H{last}"])
    dv_list(ws, "=lstDocumentation", [f"L{first}:L{last}"])
    dv_list(ws, "=lstExcused", [f"M{first}:M{last}"])
    define(wb, "nrAT_Date", "Attendance", f"$C${first}:$C${last}")
    define(wb, "nrAT_PID", "Attendance", f"$E${first}:$E${last}")
    define(wb, "nrAT_Type", "Attendance", f"$G${first}:$G${last}")
    define(wb, "nrAT_Min", "Attendance", f"$I${first}:$I${last}")
    define(wb, "nrAT_Sess", "Attendance", f"$J${first}:$J${last}")
    define(wb, "nrAT_IsPT", "Attendance", f"$K${first}:$K${last}")
    define(wb, "nrAT_Excused", "Attendance", f"$M${first}:$M${last}")
    define(wb, "nrAT_Counts", "Attendance", f"$N${first}:$N${last}")
    col_widths(ws, {"A": 3, "B": 9, "C": 11, "D": 22, "E": 9, "F": 18,
                    "G": 16, "H": 12, "I": 9, "J": 9, "K": 7, "L": 12,
                    "M": 10, "N": 9, "O": 30})
    sheet_note(ws, "Exception log: only missed/modified time gets a row. "
                   "Minutes count toward the classroom cap; Sessions toward "
                   "the PT cap. Excused or documented-modified-PT rows don't "
                   "count (policy 400).")
    return ws


def build_makeup(wb):
    ws = wb.create_sheet("Makeup")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["MakeupID", "Date", "Cadet Name", "PID", "Type",
                    "Minutes Credit", "Sessions Credit", "Linked Event",
                    "Doc Status", "Hold?", "Credit Applies?", "Notes"])
    first, last = DATA_ROW, DATA_ROW + ROWS_MAKEUP - 1
    fill_rows(ws, first, last, {
        "B": ('IF($D{r}="","","M"&TEXT(ROW()-%d,"000"))' % HDR_ROW, "fx"),
        "C": (None, "in"), "D": (None, "in"),
        "E": ('IF($D{r}="","",IFERROR(INDEX(rngCadetPIDs,MATCH($D{r},rngCadetNames,0)),"?"))', "fx"),
        "F": (None, "in"), "G": (None, "in"), "H": (None, "in"),
        "I": (None, "in"), "J": (None, "in"), "K": (None, "in"),
        "L": ('IF($D{r}="","",IF(AND($J{r}="Received",$K{r}<>"Yes"),"Yes","No"))', "fx"),
        "M": (None, "in"),
    })
    for r in range(first, last + 1):
        ws[f"C{r}"].number_format = DATE
    dv_list(ws, "=rngCadetNames", [f"D{first}:D{last}"])
    dv_list(ws, "=lstMakeupType", [f"F{first}:F{last}"])
    dv_list(ws, "=lstDocumentation", [f"J{first}:J{last}"])
    dv_list(ws, "=lstYesNo", [f"K{first}:K{last}"])
    define(wb, "nrMK_PID", "Makeup", f"$E${first}:$E${last}")
    define(wb, "nrMK_Type", "Makeup", f"$F${first}:$F${last}")
    define(wb, "nrMK_Min", "Makeup", f"$G${first}:$G${last}")
    define(wb, "nrMK_Sess", "Makeup", f"$H${first}:$H${last}")
    define(wb, "nrMK_Credit", "Makeup", f"$L${first}:$L${last}")
    col_widths(ws, {"A": 3, "B": 10, "C": 11, "D": 22, "E": 9, "F": 14,
                    "G": 13, "H": 13, "I": 13, "J": 12, "K": 8, "L": 13,
                    "M": 30})
    sheet_note(ws, "Makeup is minute-for-minute (policy 400.5). Classroom "
                   "skills time cannot be made up. Credit applies only when "
                   "documentation is Received and no Hold.")
    return ws


# --------------------------------------------------------------------------
def build_skills(wb):
    ws = wb.create_sheet("Skills")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["RecordID", "Cadet Name", "PID", "Category", "Max",
                    "Mode", "Passing", "Attempt #", "Result", "Score",
                    "Date", "Assessed By", "Attempts Used", "Status",
                    "DismissReview", "Notes"])
    first, last = DATA_ROW, DATA_ROW + ROWS_SKILLS - 1
    fill_rows(ws, first, last, {
        "B": ('IF($C{r}="","",$D{r}&"-"&$E{r}&"-"&$I{r})', "fx"),
        "C": (None, "in"),
        "D": ('IF($C{r}="","",IFERROR(INDEX(rngCadetPIDs,MATCH($C{r},rngCadetNames,0)),"?"))', "fx"),
        "E": (None, "in"),
        "F": ('IF($E{r}="","",IFERROR(INDEX(rngSM_max,MATCH($E{r},rngSM_cat,0)),""))', "fx"),
        "G": ('IF($E{r}="","",IFERROR(INDEX(rngSM_mode,MATCH($E{r},rngSM_cat,0)),""))', "fx"),
        "H": ('IF($E{r}="","",IFERROR(INDEX(rngSM_pass,MATCH($E{r},rngSM_cat,0)),""))', "fx"),
        "I": (None, "in"), "J": (None, "in"), "K": (None, "in"),
        "L": (None, "in"), "M": (None, "in"),
        "N": ('IF($C{r}="","",COUNTIFS(nrSK_PID,$D{r},nrSK_Cat,$E{r}))', "fx"),
        "O": ('IF($C{r}="","",LET(latest,MAXIFS(nrSK_Att,nrSK_PID,$D{r},'
              'nrSK_Cat,$E{r}),res,IF($I{r}=latest,$J{r},""),'
              'IF($I{r}<>latest,"(superseded)",'
              'IF(res="","Pending",IF(res="Pass","Qualified",'
              'IF(AND(res="Fail",$N{r}>=$F{r}),"FAILED OUT","Needs Remediation"))))))', "fx"),
        "P": ('IF($C{r}="","",IF($O{r}="FAILED OUT","Yes","No"))', "fx"),
        "Q": (None, "in"),
    })
    for r in range(first, last + 1):
        ws[f"L{r}"].number_format = DATE
    dv_list(ws, "=rngCadetNames", [f"C{first}:C{last}"])
    dv_list(ws, "=rngSM_cat", [f"E{first}:E{last}"])
    dv_list(ws, "=lstAttemptNum", [f"I{first}:I{last}"])
    dv_list(ws, "=lstSkillResult", [f"J{first}:J{last}"])
    cf_formula(ws, f"O{first}:O{last}", f'$O{first}="FAILED OUT"', FILL_WARNBG)
    define(wb, "nrSK_PID", "Skills", f"$D${first}:$D${last}")
    define(wb, "nrSK_Cat", "Skills", f"$E${first}:$E${last}")
    define(wb, "nrSK_Att", "Skills", f"$I${first}:$I${last}")
    define(wb, "nrSK_Res", "Skills", f"$J${first}:$J${last}")
    define(wb, "nrSK_Score", "Skills", f"$K${first}:$K${last}")
    define(wb, "nrSK_Status", "Skills", f"$O${first}:$O${last}")
    define(wb, "nrSK_Dis", "Skills", f"$P${first}:$P${last}")
    col_widths(ws, {"A": 3, "B": 14, "C": 22, "D": 9, "E": 12, "F": 6,
                    "G": 10, "H": 9, "I": 9, "J": 9, "K": 8, "L": 11,
                    "M": 16, "N": 12, "O": 16, "P": 12, "Q": 26})
    sheet_note(ws, "One row per attempt (firearms scores recorded; Top Gun "
                   "uses the firearms Score column). Exhausting max attempts "
                   "= separation review (policy 300.7).")
    return ws


# --------------------------------------------------------------------------
def build_writing(wb):
    ws = wb.create_sheet("Writing")
    ws.sheet_view.showGridLines = False
    n = len(DW.ASSIGNMENTS)  # 40
    hdrs = ["PID", "Cadet Name"] + [f"#{i}" for i in range(1, n + 1)] + \
           [f"Complete (of {n})", "Overdue Missing", "Writing Current?"]
    header_row(ws, hdrs)
    first, last = DATA_ROW, CADET_LAST
    first_ac = 4                       # column D
    last_ac = first_ac + n - 1         # column AQ (col 43)
    last_acL = get_column_letter(last_ac)
    cols = {
        "B": ('IF(Cadets!$B{r}="","",Cadets!$B{r})', "fx"),
        "C": ('IF(Cadets!$B{r}="","",Cadets!$F{r})', "fx"),
    }
    for i in range(n):
        cols[get_column_letter(first_ac + i)] = (None, "in")
    cA, cB, cC = (get_column_letter(last_ac + 1), get_column_letter(last_ac + 2),
                  get_column_letter(last_ac + 3))
    cols[cA] = ('IF($B{r}="","",COUNTIF(D{r}:%s{r},"Yes"))' % last_acL, "fx")
    cols[cB] = ('IF($B{r}="","",SUMPRODUCT((D{r}:%s{r}<>"Yes")*'
                '(TRANSPOSE(rngWMdue)<>"")*(TRANSPOSE(rngWMdue)<TODAY())))'
                % last_acL, "fx")
    cols[cC] = ('IF($B{r}="","",IF(%s{r}=0,"Yes","No"))' % cB, "fx")
    fill_rows(ws, first, last, cols)
    dv_list(ws, "=lstYesNo",
            [f"D{first}:{last_acL}{last}"])
    cf_yes_no(ws, f"{cC}{first}:{cC}{last}")
    # overdue columns highlighted red when due date passed (per column)
    for i in range(n):
        cl = get_column_letter(first_ac + i)
        cf_formula(ws, f"{cl}{first}:{cl}{last}",
                   f'AND($B{first}<>"",{cl}{first}<>"Yes",'
                   f'INDEX(rngWMdue,{i+1})<>"",INDEX(rngWMdue,{i+1})<TODAY())',
                   FILL_WARNBG)
    define(wb, "nrWRcurrent", "Writing", f"${cC}${first}:${cC}${last}")
    define(wb, "nrWRoverdue", "Writing", f"${cB}${first}:${cB}${last}")
    define(wb, "nrWRcomplete", "Writing", f"${cA}${first}:${cA}${last}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24})
    for i in range(n):
        ws.column_dimensions[get_column_letter(first_ac + i)].width = 5
    for cl in (cA, cB, cC):
        ws.column_dimensions[cl].width = 13
    sheet_note(ws, 'Mark "Yes" when an assignment is received. Red cells are '
                   "past their computed due date (see WritingMaster). Works "
                   "exactly like the previous workbook.")
    return ws


# --------------------------------------------------------------------------
def build_incidents(wb):
    ws = wb.create_sheet("Incidents")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["IncidentID", "Date", "Cadet Name", "PID", "Direction",
                    "Severity", "Category", "Description", "Reported By",
                    "Doc Ref", "ChainReview", "Resolution", "Notes"])
    first, last = DATA_ROW, DATA_ROW + ROWS_INCIDENTS - 1
    fill_rows(ws, first, last, {
        "B": ('IF($D{r}="","","I"&TEXT(ROW()-%d,"000"))' % HDR_ROW, "fx"),
        "C": (None, "in"), "D": (None, "in"),
        "E": ('IF($D{r}="","",IFERROR(INDEX(rngCadetPIDs,MATCH($D{r},rngCadetNames,0)),"?"))', "fx"),
        "F": (None, "in"), "G": (None, "in"), "H": (None, "in"),
        "I": (None, "in"), "J": (None, "in"), "K": (None, "in"),
        "L": ('IF($D{r}="","",IF(AND($F{r}="Negative",OR($G{r}="Critical",'
              '$G{r}="Major")),"Yes","No"))', "fx"),
        "M": (None, "in"), "N": (None, "in"),
    })
    for r in range(first, last + 1):
        ws[f"C{r}"].number_format = DATE
    dv_list(ws, "=rngCadetNames", [f"D{first}:D{last}"])
    dv_list(ws, "=lstIncidentDirection", [f"F{first}:F{last}"])
    dv_list(ws, "=lstSeverity", [f"G{first}:G{last}"])
    dv_list(ws, "=lstIssuetype", [f"H{first}:H{last}"])
    dv_list(ws, "=lstResolution", [f"M{first}:M{last}"])
    define(wb, "nrIN_Date", "Incidents", f"$C${first}:$C${last}")
    define(wb, "nrIN_PID", "Incidents", f"$E${first}:$E${last}")
    define(wb, "nrIN_Dir", "Incidents", f"$F${first}:$F${last}")
    define(wb, "nrIN_Sev", "Incidents", f"$G${first}:$G${last}")
    define(wb, "nrIN_Desc", "Incidents", f"$I${first}:$I${last}")
    define(wb, "nrIN_Chain", "Incidents", f"$L${first}:$L${last}")
    define(wb, "nrIN_Res", "Incidents", f"$M${first}:$M${last}")
    col_widths(ws, {"A": 3, "B": 10, "C": 11, "D": 22, "E": 9, "F": 13,
                    "G": 13, "H": 14, "I": 44, "J": 14, "K": 12, "L": 11,
                    "M": 11, "N": 26})
    sheet_note(ws, "Positive and negative incidents. Negative Major/Critical "
                   "trigger chain-of-command review (policy 600).")
    return ws


# --------------------------------------------------------------------------
def build_counseling(wb):
    ws = wb.create_sheet("Counseling")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["EntryID", "Date", "Cadet Name", "PID", "Type",
                    "Trigger / Related Flag", "Description / Plan",
                    "Conducted By", "Agency Notified?", "Notified Date",
                    "Follow-up Date", "Status", "Notes"])
    first, last = DATA_ROW, DATA_ROW + ROWS_COUNSELING - 1
    fill_rows(ws, first, last, {
        "B": ('IF($D{r}="","","C"&TEXT(ROW()-%d,"000"))' % HDR_ROW, "fx"),
        "C": (None, "in"), "D": (None, "in"),
        "E": ('IF($D{r}="","",IFERROR(INDEX(rngCadetPIDs,MATCH($D{r},rngCadetNames,0)),"?"))', "fx"),
        "F": (None, "in"), "G": (None, "in"), "H": (None, "in"),
        "I": (None, "in"), "J": (None, "in"), "K": (None, "in"),
        "L": (None, "in"), "M": (None, "in"), "N": (None, "in"),
    })
    for r in range(first, last + 1):
        for cl in ("C", "K", "L"):
            ws[f"{cl}{r}"].number_format = DATE
    dv_list(ws, "=rngCadetNames", [f"D{first}:D{last}"])
    dv_list(ws, "=lstCounselingType", [f"F{first}:F{last}"])
    dv_list(ws, "=lstYesNo", [f"J{first}:J{last}"])
    dv_list(ws, "=lstResolution", [f"M{first}:M{last}"])
    define(wb, "nrCO_Date", "Counseling", f"$C${first}:$C${last}")
    define(wb, "nrCO_PID", "Counseling", f"$E${first}:$E${last}")
    define(wb, "nrCO_Type", "Counseling", f"$F${first}:$F${last}")
    define(wb, "nrCO_Desc", "Counseling", f"$H${first}:$H${last}")
    define(wb, "nrCO_Status", "Counseling", f"$M${first}:$M${last}")
    col_widths(ws, {"A": 3, "B": 9, "C": 11, "D": 22, "E": 9, "F": 18,
                    "G": 24, "H": 40, "I": 16, "J": 13, "K": 12, "L": 12,
                    "M": 11, "N": 26})
    sheet_note(ws, "Documents every intervention (tutoring, counseling, "
                   "agency notification, performance plan). This is the "
                   "early-intervention paper trail policy 300.4.B expects — "
                   "and it feeds the since-last-email agency digest.")
    return ws


# --------------------------------------------------------------------------
def build_pt(wb):
    ws = wb.create_sheet("PT")
    ws.sheet_view.showGridLines = False
    events = [ev for ev, _, _, _ in DL.PT_EVENTS]
    short = ["Bench", "VertJump", "Pushups", "Situps", "Agility", "1.5Mi",
             "300m"]
    hdrs = (["PID", "Cadet Name"]
            + [f"Base {s}" for s in short] + ["Baseline Pass?"]
            + ["Midpoint Done?"]
            + [f"Final {s}" for s in short]
            + [f"Pts {s}" for s in short]
            + ["Final Points", "Final PT Pass?", "Improvement Index"])
    header_row(ws, hdrs)
    first, last = DATA_ROW, CADET_LAST
    cols = {
        "B": ('IF(Cadets!$B{r}="","",Cadets!$B{r})', "fx"),
        "C": ('IF(Cadets!$B{r}="","",Cadets!$F{r})', "fx"),
    }
    # base raw D..J, base pass K, midpoint L, final raw M..S, pts T..Z
    for i in range(7):
        cols[get_column_letter(4 + i)] = (None, "in")
    cols["K"] = (None, "in")
    cols["L"] = (None, "in")
    for i in range(7):
        cols[get_column_letter(13 + i)] = (None, "in")
    for i in range(7):
        cols[get_column_letter(20 + i)] = (None, "in")
    cols["AA"] = ('IF($B{r}="","",IF(COUNT(T{r}:Z{r})=0,"",SUM(T{r}:Z{r})))', "fx")
    cols["AB"] = ('IF($B{r}="","",IF($AA{r}="","",'
                  'IF(cfgPTFinalMinPoints=0,"(rubric pending)",'
                  'IF($AA{r}>=cfgPTFinalMinPoints,"Yes","No"))))', "fx")
    # improvement index: mean % gain on countable events (pushups, situps)
    # plus % time cut on runs (agility, 1.5mi, 300m); ignores blanks
    cols["AC"] = ('IF($B{r}="","",LET(pu,IFERROR(($O{r}-$F{r})/$F{r},""),'
                  'su,IFERROR(($P{r}-$G{r})/$G{r},""),'
                  'ag,IFERROR(($H{r}-$Q{r})/$H{r},""),'
                  'mi,IFERROR(($I{r}-$R{r})/$I{r},""),'
                  'tm,IFERROR(($J{r}-$S{r})/$J{r},""),'
                  'vals,IFERROR(FILTER(HSTACK(pu,su,ag,mi,tm),'
                  'ISNUMBER(HSTACK(pu,su,ag,mi,tm))),""),'
                  'IF(COUNT(vals)=0,"",ROUND(AVERAGE(vals)*100,1))))', "fx")
    fill_rows(ws, first, last, cols)
    dv_list(ws, "=lstYesNo", [f"K{first}:K{last}", f"L{first}:L{last}"])
    cf_yes_no(ws, f"AB{first}:AB{last}")
    cf_yes_no(ws, f"K{first}:K{last}")
    # pts cells yellow until rubric arrives
    for r in range(first, last + 1):
        for i in range(7):
            ws[f"{get_column_letter(20+i)}{r}"].fill = FILL_YELLOW
    define(wb, "nrPT_BasePass", "PT", f"$K${first}:$K${last}")
    define(wb, "nrPT_FinalPts", "PT", f"$AA${first}:$AA${last}")
    define(wb, "nrPT_FinalPass", "PT", f"$AB${first}:$AB${last}")
    define(wb, "nrPT_Improve", "PT", f"$AC${first}:$AC${last}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24, "AA": 11, "AB": 13, "AC": 14})
    for i in range(4, 27):
        ws.column_dimensions[get_column_letter(i)].width = 9
    sheet_note(ws, "Raw values per event: bench lbs, jump inches, pushup/situp "
                   "reps, agility & 300m seconds, 1.5mi minutes (decimal). "
                   "Pts columns take the approved rubric's points per event "
                   "(yellow until the chart arrives — see the rubric block on "
                   "Settings). Final PT failure blocks the Final Exam "
                   "(policy 500.1.H).")
    return ws


# --------------------------------------------------------------------------
def build_medical(wb):
    ws = wb.create_sheet("Medical")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["EntryID", "Date", "Cadet Name", "PID", "Type",
                    "Description", "Restriction?", "Restriction Start",
                    "Restriction End", "Cleared Date", "Doc Status",
                    "Current Status", "Notes"])
    first, last = DATA_ROW, DATA_ROW + ROWS_MEDICAL - 1
    fill_rows(ws, first, last, {
        "B": ('IF($D{r}="","","MD"&TEXT(ROW()-%d,"000"))' % HDR_ROW, "fx"),
        "C": (None, "in"), "D": (None, "in"),
        "E": ('IF($D{r}="","",IFERROR(INDEX(rngCadetPIDs,MATCH($D{r},rngCadetNames,0)),"?"))', "fx"),
        "F": (None, "in"), "G": (None, "in"), "H": (None, "in"),
        "I": (None, "in"), "J": (None, "in"), "K": (None, "in"),
        "L": (None, "in"),
        "M": ('IF($D{r}="","",IF($K{r}<>"","Cleared "&TEXT($K{r},"mm/dd"),'
              'IF($H{r}="Yes",IF(AND($J{r}<>"",$J{r}<TODAY()),"RESTRICTION EXPIRED",'
              '"ACTIVE RESTRICTION"),"Documented")))', "fx"),
        "N": (None, "in"),
    })
    for r in range(first, last + 1):
        for cl in ("C", "I", "J", "K"):
            ws[f"{cl}{r}"].number_format = DATE
    dv_list(ws, "=rngCadetNames", [f"D{first}:D{last}"])
    dv_list(ws, "=lstIssuetype", [f"F{first}:F{last}"])
    dv_list(ws, "=lstYesNo", [f"H{first}:H{last}"])
    dv_list(ws, "=lstDocumentation", [f"L{first}:L{last}"])
    cf_formula(ws, f"M{first}:M{last}",
               f'OR($M{first}="ACTIVE RESTRICTION",$M{first}="RESTRICTION EXPIRED")',
               FILL_AMBER)
    define(wb, "nrMD_PID", "Medical", f"$E${first}:$E${last}")
    define(wb, "nrMD_Restr", "Medical", f"$H${first}:$H${last}")
    define(wb, "nrMD_Status", "Medical", f"$M${first}:$M${last}")
    col_widths(ws, {"A": 3, "B": 9, "C": 11, "D": 22, "E": 9, "F": 13,
                    "G": 36, "H": 12, "I": 14, "J": 14, "K": 13, "L": 12,
                    "M": 18, "N": 26})
    sheet_note(ws, "Injuries, clearances and restrictions (policy 400.7/500). "
                   "Modified PT inside a documented restriction is not a "
                   "missed session when objectives are met — log the "
                   "restriction here and the PT Modified event on Attendance.")
    return ws


def build_all_inputs(wb):
    build_cadets(wb)
    build_examscores(wb)
    build_spelling(wb)
    build_attendance(wb)
    build_makeup(wb)
    build_skills(wb)
    build_writing(wb)
    build_incidents(wb)
    build_counseling(wb)
    build_pt(wb)
    build_medical(wb)

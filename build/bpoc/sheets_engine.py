"""Engine sheets (locked, password-protected): sysGrades, sysAttendance,
sysSkills, sysIncidents, sysPT, sysFlags, sysChecks, sysAwards, sysAudit.

Every engine sheet is one row per cadet (rows 6..55), mirroring Cadets order.
"""
from openpyxl.utils import get_column_letter

from xlb import (
    HDR_ROW, DATA_ROW, CADET_LAST, F_CALC, FILL_CALC, F_LABEL, F_SMALL,
    A_LEFT_WRAP, BOX, DATE, header_row, fill_rows, sheet_note, cf_yes_no,
    cf_formula, FILL_WARNBG, FILL_OKBG, FILL_AMBER, col_widths, define,
    protect,
)

FIRST, LAST = DATA_ROW, CADET_LAST


def _mirror():
    return {
        "B": ('IF(Cadets!$B{r}="","",Cadets!$B{r})', "fx"),
        "C": ('IF($B{r}="","",Cadets!$F{r})', "fx"),
        "D": ('IF($B{r}="","",Cadets!$I{r})', "fx"),
    }


# --------------------------------------------------------------------------
def build_sysgrades(wb):
    ws = wb.create_sheet("sysGrades")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Cadet Name", "Status", "MajCnt", "MinCnt",
                    "SpCnt", "FinCnt", "MajorAvg", "MinorAvg", "SpellAvg",
                    "FinalScore", "Current Grade", "Final Grade", "Retakes",
                    "DismissRev", "MajBelow", "MinBelow", "GradeNum",
                    "RankElig", "Rank", "AcademicElig",
                    "Need Maj Avg Left", "Need Min Avg Left", "Need Final",
                    "LastScore", "PrevScore", "GradeDrop", "ConsecFails"])
    cols = _mirror()
    cols.update({
        "E": ('IF($B{r}="","",COUNTIFS(nrES_PID,$B{r},nrES_Type,"Major",nrES_Final,"Yes"))', "fx"),
        "F": ('IF($B{r}="","",COUNTIFS(nrES_PID,$B{r},nrES_Type,"Minor",nrES_Final,"Yes"))', "fx"),
        "G": ('IF($B{r}="","",Spelling!Q{r})', "fx"),
        "H": ('IF($B{r}="","",COUNTIFS(nrES_PID,$B{r},nrES_Type,"Final",nrES_Final,"Yes"))', "fx"),
        "I": ('IF(OR($B{r}="",$E{r}=0),"",ROUND(AVERAGEIFS(nrES_Rec,nrES_PID,$B{r},nrES_Type,"Major",nrES_Final,"Yes"),2))', "fx"),
        "J": ('IF(OR($B{r}="",$F{r}=0),"",ROUND(AVERAGEIFS(nrES_Rec,nrES_PID,$B{r},nrES_Type,"Minor",nrES_Final,"Yes"),2))', "fx"),
        "K": ('IF($B{r}="","",Spelling!P{r})', "fx"),
        "L": ('IF(OR($B{r}="",$H{r}=0),"",ROUND(AVERAGEIFS(nrES_Rec,nrES_PID,$B{r},nrES_Type,"Final",nrES_Final,"Yes"),2))', "fx"),
        "M": ('IF($B{r}="","",LET(w,($E{r}>0)*cfgWeightMajor+($F{r}>0)*cfgWeightMinor'
              '+(N($G{r})>0)*cfgWeightSpelling+($H{r}>0)*cfgWeightFinal,'
              'IF(w=0,"",ROUND((($E{r}>0)*cfgWeightMajor*N($I{r})'
              '+($F{r}>0)*cfgWeightMinor*N($J{r})'
              '+(N($G{r})>0)*cfgWeightSpelling*N($K{r})'
              '+($H{r}>0)*cfgWeightFinal*N($L{r}))/w,2))))', "fx"),
        "N": ('IF(OR($B{r}="",$H{r}=0),"",$M{r})', "fx"),
        "O": ('IF($B{r}="","",COUNTIFS(nrES_PID,$B{r},nrES_Att,2))', "fx"),
        "P": ('IF($B{r}="","",COUNTIFS(nrES_PID,$B{r},nrES_Dis,"Yes"))', "fx"),
        "Q": ('IF($B{r}="","",IF(AND($E{r}>=cfgThresholdAfterExam,$I{r}<>"",'
              '$I{r}<cfgThresholdScore),"Yes","No"))', "fx"),
        "R": ('IF($B{r}="","",IF(AND($F{r}>=cfgThresholdAfterExam,$J{r}<>"",'
              '$J{r}<cfgThresholdScore),"Yes","No"))', "fx"),
        "S": ('IF($B{r}="","",IF($M{r}="",-1,$M{r}))', "fx"),
        "T": ('IF($B{r}="","",IF(AND($D{r}="Active",$P{r}=0,'
              'sysSkills!$J{r}<>"Yes",sysIncidents!$J{r}=0,$M{r}<>""),"Yes","No"))', "fx"),
        "U": ('IF($T{r}<>"Yes","",SUMPRODUCT(($T$%d:$T$%d="Yes")*'
              '(($S$%d:$S$%d>$S{r})+($S$%d:$S$%d=$S{r})*(ROW($S$%d:$S$%d)<ROW())))+1)'
              % (FIRST, LAST, FIRST, LAST, FIRST, LAST, FIRST, LAST), "fx"),
        "V": ('IF($B{r}="","",IF(AND(OR($E{r}=0,$I{r}>=cfgThresholdScore),'
              'OR($F{r}=0,$J{r}>=cfgThresholdScore),'
              'OR(N($G{r})=0,$K{r}>=cfgThresholdScore),'
              'OR($H{r}=0,$L{r}>=cfgThresholdScore)),"Yes","No"))', "fx"),
        # projections: average needed on remaining exams to reach 70 in category
        "W": ('IF($B{r}="","",LET(tot,COUNTIFS(rngEPuse,"Yes",rngEPtype,"Major"),'
              'left,tot-$E{r},IF(left<=0,"done",'
              'MAX(0,ROUND((cfgThresholdScore*tot-N($I{r})*$E{r})/left,1)))))', "fx"),
        "X": ('IF($B{r}="","",LET(tot,COUNTIFS(rngEPuse,"Yes",rngEPtype,"Minor"),'
              'left,tot-$F{r},IF(left<=0,"done",'
              'MAX(0,ROUND((cfgThresholdScore*tot-N($J{r})*$F{r})/left,1)))))', "fx"),
        "Y": ('IF($B{r}="","",IF($H{r}>0,"done",cfgThresholdScore))', "fx"),
        # trend helpers: last & previous recorded first-attempt scores by seq
        "Z": ('IF($B{r}="","",LET(s,MAXIFS(nrES_Seq,nrES_PID,$B{r},nrES_Att,1,'
              'nrES_Raw,">=0"),IF(s=0,"",SUMIFS(nrES_Raw,nrES_PID,$B{r},'
              'nrES_Att,1,nrES_Seq,s))))', "fx"),
        "AA": ('IF($B{r}="","",LET(s,MAXIFS(nrES_Seq,nrES_PID,$B{r},nrES_Att,1,'
               'nrES_Raw,">=0"),IF(s=0,"",LET(p,MAXIFS(nrES_Seq,nrES_PID,$B{r},'
               'nrES_Att,1,nrES_Seq,"<"&s),IF(p=0,"",SUMIFS(nrES_Raw,'
               'nrES_PID,$B{r},nrES_Att,1,nrES_Seq,p))))))', "fx"),
        "AB": ('IF(OR($B{r}="",$Z{r}="",$AA{r}=""),"",ROUND($AA{r}-$Z{r},1))', "fx"),
        "AC": ('IF($B{r}="","",LET(s,MAXIFS(nrES_Seq,nrES_PID,$B{r},nrES_Att,1,'
               'nrES_Raw,">=0"),IF(s=0,0,LET(a,IF(SUMIFS(nrES_Raw,nrES_PID,$B{r},'
               'nrES_Att,1,nrES_Seq,s)<cfgPassingScore,1,0),'
               'p,MAXIFS(nrES_Seq,nrES_PID,$B{r},nrES_Att,1,nrES_Seq,"<"&s),'
               'IF(OR(a=0,p=0),a,a+IF(SUMIFS(nrES_Raw,nrES_PID,$B{r},'
               'nrES_Att,1,nrES_Seq,p)<cfgPassingScore,1,0))))))', "fx"),
    })
    fill_rows(ws, FIRST, LAST, cols)
    define(wb, "nrGRrank", "sysGrades", f"$U${FIRST}:$U${LAST}")
    define(wb, "nrGRcurrent", "sysGrades", f"$M${FIRST}:$M${LAST}")
    define(wb, "nrGRfinal", "sysGrades", f"$N${FIRST}:$N${LAST}")
    define(wb, "nrGRacademic", "sysGrades", f"$V${FIRST}:$V${LAST}")
    define(wb, "nrGRmajavg", "sysGrades", f"$I${FIRST}:$I${LAST}")
    define(wb, "nrGRminavg", "sysGrades", f"$J${FIRST}:$J${LAST}")
    define(wb, "nrGRgradedrop", "sysGrades", f"$AB${FIRST}:$AB${LAST}")
    define(wb, "nrGRconsec", "sysGrades", f"$AC${FIRST}:$AC${LAST}")
    define(wb, "nrGRdisrev", "sysGrades", f"$P${FIRST}:$P${LAST}")
    define(wb, "nrGRretakes", "sysGrades", f"$O${FIRST}:$O${LAST}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24})
    sheet_note(ws, "Weighted per policy 300.2 (Major 40 / Minor 30 / Spelling "
                   "10 / Final 20); pass requires 70 in EACH category. "
                   "Projection columns show the average needed on remaining "
                   "exams. Locked engine — do not edit.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_sysattendance(wb):
    ws = wb.create_sheet("sysAttendance")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Cadet Name", "Status", "Cl Missed", "Cl MadeUp",
                    "Cl Net", "Cl Cap", "Cl %", "Cl Tier", "Cl Elig",
                    "Cl Owed", "PT Missed", "PT MadeUp", "PT Net", "PT Cap",
                    "PT %", "PT Tier", "PT Elig", "Makeup Complete?",
                    "Tardy", "FullAbs", "PartAbs", "PTMissCnt", "Excused"])
    cols = _mirror()
    cols.update({
        "E": ('IF($B{r}="","",SUMIFS(nrAT_Min,nrAT_PID,$B{r},nrAT_Counts,"Yes",nrAT_IsPT,"No"))', "fx"),
        "F": ('IF($B{r}="","",SUMIFS(nrMK_Min,nrMK_PID,$B{r},nrMK_Credit,"Yes",nrMK_Type,"Classroom"))', "fx"),
        "G": ('IF($B{r}="","",MAX(0,$E{r}-$F{r}))', "fx"),
        "H": ('IF($B{r}="","",cfgClassroomCapMinutes)', "fx"),
        "I": ('IF($B{r}="","",IF($H{r}=0,0,$G{r}/$H{r}))', "fx"),
        "J": ('IF($B{r}="","",IF($I{r}>=1,"At Cap",IF($I{r}>=cfgAttendanceCriticalPct,'
              '"Critical",IF($I{r}>=cfgAttendanceAdvisoryPct,"Advisory","OK"))))', "fx"),
        "K": ('IF($B{r}="","",IF($H{r}=0,"Yes",IF($G{r}<$H{r},"Yes","No")))', "fx"),
        "L": ('IF($B{r}="","",MAX(0,$E{r}-$F{r}))', "fx"),
        "M": ('IF($B{r}="","",SUMIFS(nrAT_Sess,nrAT_PID,$B{r},nrAT_Counts,"Yes",nrAT_IsPT,"Yes"))', "fx"),
        "N": ('IF($B{r}="","",SUMIFS(nrMK_Sess,nrMK_PID,$B{r},nrMK_Credit,"Yes",nrMK_Type,"PT"))', "fx"),
        "O": ('IF($B{r}="","",MAX(0,$M{r}-$N{r}))', "fx"),
        "P": ('IF($B{r}="","",cfgPTCapSessions)', "fx"),
        "Q": ('IF($B{r}="","",IF($P{r}=0,0,$O{r}/$P{r}))', "fx"),
        "R": ('IF($B{r}="","",IF($Q{r}>=1,"At Cap",IF($Q{r}>=cfgAttendanceCriticalPct,'
              '"Critical",IF($Q{r}>=cfgAttendanceAdvisoryPct,"Advisory","OK"))))', "fx"),
        "S": ('IF($B{r}="","",IF($P{r}=0,"Yes",IF($O{r}<$P{r},"Yes","No")))', "fx"),
        "T": ('IF($B{r}="","",IF($L{r}=0,"Yes","No"))', "fx"),
        "U": ('IF($B{r}="","",COUNTIFS(nrAT_PID,$B{r},nrAT_Type,"Tardy"))', "fx"),
        "V": ('IF($B{r}="","",COUNTIFS(nrAT_PID,$B{r},nrAT_Type,"Absent Full Day"))', "fx"),
        "W": ('IF($B{r}="","",COUNTIFS(nrAT_PID,$B{r},nrAT_Type,"Absent Partial Day"))', "fx"),
        "X": ('IF($B{r}="","",COUNTIFS(nrAT_PID,$B{r},nrAT_Type,"PT Missed"))', "fx"),
        "Y": ('IF($B{r}="","",COUNTIFS(nrAT_PID,$B{r},nrAT_Excused,"Excused"))', "fx"),
    })
    # fix column collision: header has 24 cols B..Y; L duplicated (Cl Owed)
    cols["L"] = ('IF($B{r}="","",MAX(0,$E{r}-$F{r}))', "fx")
    fill_rows(ws, FIRST, LAST, cols)
    define(wb, "nrATTclTier", "sysAttendance", f"$J${FIRST}:$J${LAST}")
    define(wb, "nrATTclElig", "sysAttendance", f"$K${FIRST}:$K${LAST}")
    define(wb, "nrATTclPct", "sysAttendance", f"$I${FIRST}:$I${LAST}")
    define(wb, "nrATTclOwed", "sysAttendance", f"$L${FIRST}:$L${LAST}")
    define(wb, "nrATTptTier", "sysAttendance", f"$R${FIRST}:$R${LAST}")
    define(wb, "nrATTptElig", "sysAttendance", f"$S${FIRST}:$S${LAST}")
    define(wb, "nrATTptPct", "sysAttendance", f"$Q${FIRST}:$Q${LAST}")
    define(wb, "nrATTmakeupOK", "sysAttendance", f"$T${FIRST}:$T${LAST}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24})
    sheet_note(ws, "Minutes vs the 5% classroom cap; PT sessions vs the "
                   "5-session cap (policy 400). 'Cl Owed' = missed minutes "
                   "not yet made up — must reach 0 before graduation. Locked.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_sysskills(wb):
    ws = wb.create_sheet("sysSkills")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Cadet Name", "Status", "Categories Attempted",
                    "Qualified", "Needs Remediation", "Pending",
                    "Failed Out Cats", "FailedOut?", "Skills Elig",
                    "Firearms Avg", "Firearms Best"])
    cols = _mirror()
    cols.update({
        "E": ('IF($B{r}="","",SUMPRODUCT((rngSM_cat<>"")*'
              '(COUNTIFS(nrSK_PID,$B{r},nrSK_Cat,rngSM_cat)>0)))', "fx"),
        "F": ('IF($B{r}="","",SUMPRODUCT((rngSM_cat<>"")*'
              '(COUNTIFS(nrSK_PID,$B{r},nrSK_Cat,rngSM_cat,nrSK_Status,"Qualified")>0)))', "fx"),
        "G": ('IF($B{r}="","",SUMPRODUCT((rngSM_cat<>"")*'
              '(COUNTIFS(nrSK_PID,$B{r},nrSK_Cat,rngSM_cat,nrSK_Status,"Needs Remediation")>0)))', "fx"),
        "H": ('IF($B{r}="","",SUMPRODUCT((rngSM_cat<>"")*'
              '(COUNTIFS(nrSK_PID,$B{r},nrSK_Cat,rngSM_cat,nrSK_Status,"Pending")>0)))', "fx"),
        "I": ('IF($B{r}="","",COUNTIFS(nrSK_PID,$B{r},nrSK_Dis,"Yes"))', "fx"),
        "J": ('IF($B{r}="","",IF($I{r}>0,"Yes","No"))', "fx"),
        "K": ('IF($B{r}="","",IF($J{r}="Yes","No",IF($G{r}>0,"No","Yes")))', "fx"),
        "L": ('IF($B{r}="","",IFERROR(ROUND(AVERAGEIFS(nrSK_Score,nrSK_PID,$B{r},'
              'nrSK_Cat,"Firearms"),2),""))', "fx"),
        "M": ('IF($B{r}="","",IFERROR(MAXIFS(nrSK_Score,nrSK_PID,$B{r},'
              'nrSK_Cat,"Firearms"),""))', "fx"),
    })
    fill_rows(ws, FIRST, LAST, cols)
    define(wb, "nrSKfailedout", "sysSkills", f"$J${FIRST}:$J${LAST}")
    define(wb, "nrSKelig", "sysSkills", f"$K${FIRST}:$K${LAST}")
    define(wb, "nrSKfirearmsAvg", "sysSkills", f"$L${FIRST}:$L${LAST}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24})
    sheet_note(ws, "Column letters I/J intentionally match V5.4 cross-refs "
                   "(sysGrades reads sysSkills!$I / $J-style flags). Locked.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_sysincidents(wb):
    ws = wb.create_sheet("sysIncidents")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Cadet Name", "Status", "Total", "Negative",
                    "Positive", "Open Negative", "NegMajorCritical",
                    "OpenChainReview", "ChainReviewCnt", "Incidents Elig",
                    "Counseling Cnt", "Open Counseling"])
    cols = _mirror()
    cols.update({
        "E": ('IF($B{r}="","",COUNTIF(nrIN_PID,$B{r}))', "fx"),
        "F": ('IF($B{r}="","",COUNTIFS(nrIN_PID,$B{r},nrIN_Dir,"Negative"))', "fx"),
        "G": ('IF($B{r}="","",COUNTIFS(nrIN_PID,$B{r},nrIN_Dir,"Positive"))', "fx"),
        "H": ('IF($B{r}="","",COUNTIFS(nrIN_PID,$B{r},nrIN_Dir,"Negative",'
              'nrIN_Res,"Open")+COUNTIFS(nrIN_PID,$B{r},nrIN_Dir,"Negative",'
              'nrIN_Res,"In Review"))', "fx"),
        "I": ('IF($B{r}="","",COUNTIFS(nrIN_PID,$B{r},nrIN_Dir,"Negative",'
              'nrIN_Sev,"Major")+COUNTIFS(nrIN_PID,$B{r},nrIN_Dir,"Negative",'
              'nrIN_Sev,"Critical"))', "fx"),
        "J": ('IF($B{r}="","",COUNTIFS(nrIN_PID,$B{r},nrIN_Chain,"Yes",'
              'nrIN_Res,"Open")+COUNTIFS(nrIN_PID,$B{r},nrIN_Chain,"Yes",'
              'nrIN_Res,"In Review"))', "fx"),
        "K": ('IF($B{r}="","",IF($J{r}=0,"Yes","No"))', "fx"),
        "L": ('IF($B{r}="","",COUNTIF(nrCO_PID,$B{r}))', "fx"),
        "M": ('IF($B{r}="","",COUNTIFS(nrCO_PID,$B{r},nrCO_Status,"Open")'
              '+COUNTIFS(nrCO_PID,$B{r},nrCO_Status,"In Review"))', "fx"),
    })
    fill_rows(ws, FIRST, LAST, cols)
    define(wb, "nrINCopenNeg", "sysIncidents", f"$H${FIRST}:$H${LAST}")
    define(wb, "nrINCelig", "sysIncidents", f"$K${FIRST}:$K${LAST}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24})
    sheet_note(ws, "Incident + counseling rollup per cadet. Locked.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_sysflags(wb):
    """Configurable flag engine: one column per flag, reason text, count."""
    ws = wb.create_sheet("sysFlags")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Cadet Name", "Status", "F:ConsecFails",
                    "F:GradeDrop", "F:CategoryRisk", "F:Spelling",
                    "F:Attendance", "F:Incidents", "F:Writing", "F:Retest",
                    "F:PT", "F:Medical", "Flag Count", "Reasons"])
    cols = _mirror()
    cols.update({
        "E": ('IF($B{r}="","",IF(N(nrGRconsec %s)>=cfgFlagConsecFails,1,0))'
              % "", "fx"),
        "F": ('IF($B{r}="","",IF(AND(sysGrades!$AB{r}<>"",'
              'sysGrades!$AB{r}>=cfgFlagGradeDrop),1,0))', "fx"),
        "G": ('IF($B{r}="","",IF(OR(AND(sysGrades!$I{r}<>"",'
              'sysGrades!$I{r}<cfgThresholdScore+cfgFlagCategoryMargin),'
              'AND(sysGrades!$J{r}<>"",sysGrades!$J{r}<cfgThresholdScore'
              '+cfgFlagCategoryMargin)),1,0))', "fx"),
        "H": ('IF($B{r}="","",IF(Spelling!$R{r}="INTERVENTION",1,0))', "fx"),
        "I": ('IF($B{r}="","",IF(OR(sysAttendance!$I{r}>=cfgFlagAttendancePct,'
              'sysAttendance!$Q{r}>=cfgFlagAttendancePct),1,0))', "fx"),
        "J": ('IF($B{r}="","",IF(sysIncidents!$H{r}>=cfgFlagOpenIncidents,1,0))', "fx"),
        "K": ('IF($B{r}="","",IF(N(Writing!$AS{r})>=cfgFlagOverdueWriting,1,0))', "fx"),
        "L": ('IF($B{r}="","",IF(COUNTIFS(nrES_PID,$B{r},nrES_RetStat,"OVERDUE")>0,1,0))', "fx"),
        "M": ('IF($B{r}="","",IF(OR(PT!$K{r}="No",PT!$AB{r}="No"),1,0))', "fx"),
        "N": ('IF($B{r}="","",IF(COUNTIFS(nrMD_PID,$B{r},nrMD_Status,'
              '"RESTRICTION EXPIRED")>0,1,0))', "fx"),
        "O": ('IF($B{r}="","",SUM($E{r}:$N{r}))', "fx"),
        "P": ('IF($B{r}="","",IF($O{r}=0,"",TEXTJOIN("; ",TRUE,'
              'IF($E{r}=1,"consecutive exam fails",""),'
              'IF($F{r}=1,"grade dropped "&sysGrades!$AB{r}&" pts",""),'
              'IF($G{r}=1,"category avg near 70",""),'
              'IF($H{r}=1,"spelling below "&cfgSpellInterventionAvg,""),'
              'IF($I{r}=1,"attendance at "&TEXT(MAX(sysAttendance!$I{r},'
              'sysAttendance!$Q{r}),"0%")&" of cap",""),'
              'IF($J{r}=1,"open negative incidents",""),'
              'IF($K{r}=1,N(Writing!$AS{r})&" overdue writing",""),'
              'IF($L{r}=1,"RETEST OVERDUE",""),'
              'IF($M{r}=1,"PT failure",""),'
              'IF($N{r}=1,"medical restriction expired",""))))', "fx"),
    })
    # E needs the row-scoped reference, not the whole named range
    cols["E"] = ('IF($B{r}="","",IF(N(sysGrades!$AC{r})>=cfgFlagConsecFails,1,0))', "fx")
    fill_rows(ws, FIRST, LAST, cols)
    define(wb, "nrFLcount", "sysFlags", f"$O${FIRST}:$O${LAST}")
    define(wb, "nrFLreasons", "sysFlags", f"$P${FIRST}:$P${LAST}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24, "P": 70})
    sheet_note(ws, "Each flag threshold lives on Settings. WatchList sorts "
                   "by Flag Count and shows Reasons verbatim. Locked.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_syschecks(wb):
    ws = wb.create_sheet("sysChecks")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Cadet Name", "Status", "Academic", "Classroom",
                    "PT Sessions", "Skills", "Incidents", "Writing",
                    "Makeup Complete", "Final PT Pass", "DismissReview",
                    "GraduationElig", "Blocking Issues", "Final Exam Elig"])
    cols = _mirror()
    cols.update({
        "E": ('IF($B{r}="","",sysGrades!$V{r})', "fx"),
        "F": ('IF($B{r}="","",sysAttendance!$K{r})', "fx"),
        "G": ('IF($B{r}="","",sysAttendance!$S{r})', "fx"),
        "H": ('IF($B{r}="","",sysSkills!$K{r})', "fx"),
        "I": ('IF($B{r}="","",sysIncidents!$K{r})', "fx"),
        "J": ('IF($B{r}="","",IF(Writing!$AT{r}="Yes","Yes","No"))', "fx"),
        "K": ('IF($B{r}="","",sysAttendance!$T{r})', "fx"),
        "L": ('IF($B{r}="","",IF(PT!$AB{r}="No","No","Yes"))', "fx"),
        "M": ('IF($B{r}="","",IF(OR(sysGrades!$P{r}>0,sysSkills!$J{r}="Yes",'
              'sysIncidents!$J{r}>0),"Yes","No"))', "fx"),
        "N": ('IF($B{r}="","",IF(AND($E{r}="Yes",$F{r}="Yes",$G{r}="Yes",'
              '$H{r}="Yes",$I{r}="Yes",$J{r}="Yes",$K{r}="Yes",$L{r}="Yes",'
              '$M{r}="No"),"Yes","No"))', "fx"),
        "O": ('IF($B{r}="","",IF($N{r}="Yes","Eligible",TRIM('
              'IF($E{r}<>"Yes","Academic; ","")&'
              'IF($F{r}<>"Yes","Classroom; ","")&'
              'IF($G{r}<>"Yes","PT sessions; ","")&'
              'IF($H{r}<>"Yes","Skills; ","")&'
              'IF($I{r}<>"Yes","Incidents; ","")&'
              'IF($J{r}<>"Yes","Writing; ","")&'
              'IF($K{r}<>"Yes","Makeup owed; ","")&'
              'IF($L{r}<>"Yes","Final PT; ","")&'
              'IF($M{r}="Yes","Dismissal review; ",""))))', "fx"),
        "P": ('IF($B{r}="","",IF(PT!$AB{r}="No","No","Yes"))', "fx"),
    })
    fill_rows(ws, FIRST, LAST, cols)
    cf_yes_no(ws, f"N{FIRST}:N{LAST}")
    define(wb, "nrCKgradElig", "sysChecks", f"$N${FIRST}:$N${LAST}")
    define(wb, "nrCKblocking", "sysChecks", f"$O${FIRST}:$O${LAST}")
    define(wb, "nrCKfinalExamElig", "sysChecks", f"$P${FIRST}:$P${LAST}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24, "O": 50})
    sheet_note(ws, "Graduation gate per policy: 70 in each category, under "
                   "attendance caps, makeup complete, skills qualified, "
                   "writing current, final PT passed, no open dismissal "
                   "review. 'Final Exam Elig' enforces 500.1.H. Locked.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_sysawards(wb):
    ws = wb.create_sheet("sysAwards")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Award", "Computed Winner", "Basis", "Coordinator Override",
                    "FINAL Winner", "Notes"])
    r = FIRST
    rows = [
        ("Valedictorian",
         'IFERROR(INDEX(rngCadetNames,MATCH(1,nrGRrank,0)),"")',
         '"Rank #1 — final grade "&IFERROR(INDEX(nrGRfinal,MATCH(1,nrGRrank,0)),"n/a")'),
        ("Physical Fitness",
         'IFERROR(INDEX(rngCadetNames,MATCH(MAX(nrPT_FinalPts),nrPT_FinalPts,0)),"")',
         '"Top final PT points: "&IFERROR(MAX(nrPT_FinalPts),"n/a")'),
        ("Top Gun",
         'IFERROR(INDEX(rngCadetNames,MATCH(MAX(nrSKfirearmsAvg),nrSKfirearmsAvg,0)),"")',
         '"Top firearms average: "&IFERROR(ROUND(MAX(nrSKfirearmsAvg),2),"n/a")'),
        ("Grit (most improved)",
         'IFERROR(INDEX(rngCadetNames,MATCH(MAX(nrPT_Improve),nrPT_Improve,0)),"")',
         '"Best improvement index (PT deltas): "&IFERROR(MAX(nrPT_Improve),"n/a")'
         '&"% — verify with academics trend before finalizing"'),
    ]
    for name, winner_fx, basis_fx in rows:
        ws.cell(row=r, column=2, value=name).font = F_LABEL
        c = ws.cell(row=r, column=3, value="=" + winner_fx)
        c.font = F_CALC
        c.fill = FILL_CALC
        c.border = BOX
        c2 = ws.cell(row=r, column=4, value="=" + basis_fx)
        c2.font = F_SMALL
        c2.alignment = A_LEFT_WRAP
        from xlb import FILL_INPUT, F_INPUT
        o = ws.cell(row=r, column=5)
        o.fill = FILL_INPUT
        o.font = F_INPUT
        o.border = BOX
        f = ws.cell(row=r, column=6,
                    value=f'=IF($E{r}<>"",$E{r},$C{r})')
        f.font = F_LABEL
        f.border = BOX
        n = ws.cell(row=r, column=7)
        n.fill = FILL_INPUT
        n.font = F_INPUT
        n.border = BOX
        ws.row_dimensions[r].height = 30
        r += 1
    from xlb import dv_list as _dv
    _dv(ws, "=rngCadetNames", [f"E{FIRST}:E{r-1}"])
    define(wb, "nrAWfinal", "sysAwards", f"$F${FIRST}:$F${r-1}")
    define(wb, "nrAWnames", "sysAwards", f"$B${FIRST}:$B${r-1}")
    col_widths(ws, {"A": 3, "B": 22, "C": 24, "D": 52, "E": 24, "F": 24,
                    "G": 30})
    sheet_note(ws, "Data-driven picks with coordinator override (override "
                   "always wins). Grit is decision support — the computed "
                   "index only sees PT deltas; weigh academic improvement "
                   "and perseverance before finalizing.")
    return ws


# --------------------------------------------------------------------------
def build_sysaudit(wb):
    """TCOLE audit-readiness rollups that need engine math."""
    ws = wb.create_sheet("sysAudit")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Check", "Value", "Target", "Status", "Detail"])
    r = FIRST
    checks = [
        ("Delivered hours vs 736 (exact)",
         "nrCHtotalDelivered", "cfgRequiredHours",
         'IF(N(nrCHtotalDelivered)<N(cfgRequiredHours),"SHORT",'
         'IF(N(nrCHtotalDelivered)>N(cfgRequiredHours),"OVER - report excess as #101","OK"))',
         '"Report the BPOC at exactly 736 hrs; excess must go to Addendum course #101"'),
        ("Chapters short of TCOLE minimum",
         'SUMPRODUCT((nrCHdeliv<>"")*(nrCHdeliv>0)*(nrCHdeliv<nrCHmin))', "0",
         'IF(SUMPRODUCT((nrCHdeliv<>"")*(nrCHdeliv>0)*(nrCHdeliv<nrCHmin))=0,"OK","CHECK")',
         '"Rule 218.1(C)(4): failure to meet minimum course length = denial of training"'),
        ("Chapter training files incomplete",
         'SUMPRODUCT((nrCHname<>"")*(nrCHfileOK="No"))', "0",
         'IF(SUMPRODUCT((nrCHname<>"")*(nrCHfileOK="No"))=0,"OK","CHECK")',
         '"SME lesson plan, instructor bio, sign-ins w/ PID, assessment, grade sheet, eval"'),
        ("Special TCOLE chapter requirements unmet",
         'SUMPRODUCT((nrCHspecial<>"")*(nrCHspecialMet<>"Yes")*(nrCHspecialMet<>"N/A"))', "0",
         'IF(SUMPRODUCT((nrCHspecial<>"")*(nrCHspecialMet<>"Yes")*(nrCHspecialMet<>"N/A"))=0,"OK","CHECK")',
         '"TIM, SFST, ALERRT, CPR/AED, TCIC #4800, Crime Stoppers, Seven Step, canine scenarios, CIT lead, SIT 3232"'),
        ("Chapters without evals collected",
         'SUMPRODUCT((nrCHname<>"")*(nrCHevals<>"On File")*(nrCHevals<>"N/A"))', "0",
         'IF(SUMPRODUCT((nrCHname<>"")*(nrCHevals<>"On File")*(nrCHevals<>"N/A"))=0,"OK","CHECK")',
         '"Course critiques per chapter (print from Print Center)"'),
        ("Program requirements unchecked (Audit sheet)",
         'SUMPRODUCT((nrPRGitems<>"")*(nrPRGmet<>"Yes"))', "0",
         'IF(SUMPRODUCT((nrPRGitems<>"")*(nrPRGmet<>"Yes"))=0,"OK","CHECK")',
         '"Rule 215.9 distribution, TCLEDDS roster, 736 exact reporting, facility, assessments"'),
        ("Instructors not audit-ready",
         'SUMPRODUCT((nrInstrReady="No")*1)', "0",
         'IF(SUMPRODUCT((nrInstrReady="No")*1)=0,"OK","CHECK")',
         '"Instructors sheet: needs TCOLE type or SME letter on file"'),
        ("Overdue retests",
         'COUNTIF(nrES_RetStat,"OVERDUE")', "0",
         'IF(COUNTIF(nrES_RetStat,"OVERDUE")=0,"OK","ACT NOW")',
         '"Policy 300.5: retest within 5 class days"'),
        ("Cadets with makeup owed",
         'SUMPRODUCT((nrATTmakeupOK="No")*1)', "0",
         'IF(SUMPRODUCT((nrATTmakeupOK="No")*1)=0,"OK","CHECK")',
         '"Missed minutes not yet made up"'),
        ("Active cadets failing a category",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrGRacademic="No"))', "0",
         'IF(SUMPRODUCT((nrCadetStatus="Active")*(nrGRacademic="No"))=0,"OK","CHECK")',
         '"70 in each category (policy 300.2)"'),
        ("Spelling interventions open",
         'COUNTIF(nrSpellFlag,"INTERVENTION")', "0",
         'IF(COUNTIF(nrSpellFlag,"INTERVENTION")=0,"OK","CHECK")',
         '"Document intervention on Counseling log (300.4.B)"'),
    ]
    for name, val_fx, target, stat_fx, detail_fx in checks:
        ws.cell(row=r, column=2, value=name).font = F_LABEL
        v = ws.cell(row=r, column=3, value="=" + val_fx)
        v.font = F_CALC
        v.border = BOX
        t = ws.cell(row=r, column=4, value="=" + target if target.startswith("cfg")
                    else ("=" + target if not target.isdigit() else int(target)))
        t.font = F_CALC
        s = ws.cell(row=r, column=5, value="=" + stat_fx)
        s.font = F_LABEL
        s.border = BOX
        d = ws.cell(row=r, column=6, value="=" + detail_fx)
        d.font = F_SMALL
        d.alignment = A_LEFT_WRAP
        r += 1
    cf_formula(ws, f"E{FIRST}:E{r-1}", f'$E{FIRST}<>"OK"', FILL_WARNBG)
    cf_formula(ws, f"E{FIRST}:E{r-1}", f'$E{FIRST}="OK"', FILL_OKBG)
    define(wb, "cfgRequiredHours", "sysAudit", f"$H${FIRST}")
    ws.cell(row=FIRST, column=8, value=736).font = F_SMALL
    define(wb, "nrAUDstatus", "sysAudit", f"$E${FIRST}:$E${r-1}")
    define(wb, "nrAUDcheck", "sysAudit", f"$B${FIRST}:$B${r-1}")
    col_widths(ws, {"A": 3, "B": 36, "C": 10, "D": 9, "E": 11, "F": 58,
                    "G": 3, "H": 8})
    sheet_note(ws, "Engine half of the TCOLE Audit sheet — the printable "
                   "Audit sheet reads these. Locked.")
    protect(ws)
    return ws


def build_all_engine(wb):
    build_sysgrades(wb)
    build_sysattendance(wb)
    build_sysskills(wb)
    build_sysincidents(wb)
    build_sysflags(wb)
    build_syschecks(wb)
    build_sysawards(wb)
    build_sysaudit(wb)

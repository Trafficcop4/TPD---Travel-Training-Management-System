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
        # counts and averages must use the SAME criteria or the count=0 guard
        # on I/J/L is defeated: a pending retest row is the final attempt but
        # carries no Recorded score, so it must not be counted either.
        "E": ('IF($B{r}="","",COUNTIFS(nrES_PID,$B{r},nrES_Type,"Major",nrES_Final,"Yes",nrES_Rec,">=0"))', "fx"),
        "F": ('IF($B{r}="","",COUNTIFS(nrES_PID,$B{r},nrES_Type,"Minor",nrES_Final,"Yes",nrES_Rec,">=0"))', "fx"),
        "G": ('IF($B{r}="","",Spelling!Q{r})', "fx"),
        "H": ('IF($B{r}="","",COUNTIFS(nrES_PID,$B{r},nrES_Type,"Final",nrES_Final,"Yes",nrES_Rec,">=0"))', "fx"),
        "I": ('IF(OR($B{r}="",$E{r}=0),"",IFERROR(ROUND(AVERAGEIFS(nrES_Rec,nrES_PID,$B{r},nrES_Type,"Major",nrES_Final,"Yes",nrES_Rec,">=0"),2),""))', "fx"),
        "J": ('IF(OR($B{r}="",$F{r}=0),"",IFERROR(ROUND(AVERAGEIFS(nrES_Rec,nrES_PID,$B{r},nrES_Type,"Minor",nrES_Final,"Yes",nrES_Rec,">=0"),2),""))', "fx"),
        "K": ('IF($B{r}="","",Spelling!P{r})', "fx"),
        "L": ('IF(OR($B{r}="",$H{r}=0),"",IFERROR(ROUND(AVERAGEIFS(nrES_Rec,nrES_PID,$B{r},nrES_Type,"Final",nrES_Final,"Yes",nrES_Rec,">=0"),2),""))', "fx"),
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
    # column L is the FINAL EXAM average; column N is the weighted composite.
    # Outputs that label a row "Final Exam" must read L, never N.
    define(wb, "nrGRfinalExam", "sysGrades", f"$L${FIRST}:$L${LAST}")
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
        # made-up minutes are read from the PER-EVENT reconciliation on
        # Attendance (col P), which is capped at each event's own duration —
        # NOT from the raw Makeup rows. Summing nrMK_Min directly let a
        # 360-minute makeup booked against a 120-minute event pay off 240
        # minutes of a completely different absence, so the gate below read
        # "no makeup owed" while that event's own ledger still said OPEN.
        "F": ('IF($B{r}="","",SUMIFS(nrAT_MadeUpMin,nrAT_PID,$B{r},'
              'nrAT_Counts,"Yes",nrAT_IsPT,"No"))', "fx"),
        "G": ('IF($B{r}="","",MAX(0,$E{r}-$F{r}))', "fx"),
        "H": ('IF($B{r}="","",cfgClassroomCapMinutes)', "fx"),
        "I": ('IF($B{r}="","",IF($H{r}=0,0,$G{r}/$H{r}))', "fx"),
        "J": ('IF($B{r}="","",IF($I{r}>=1,"At Cap",IF($I{r}>=cfgAttendanceCriticalPct,'
              '"Critical",IF($I{r}>=cfgAttendanceAdvisoryPct,"Advisory","OK"))))', "fx"),
        "K": ('IF($B{r}="","",IF($H{r}=0,"Yes",IF($G{r}<$H{r},"Yes","No")))', "fx"),
        "L": ('IF($B{r}="","",MAX(0,$E{r}-$F{r}))', "fx"),
        "M": ('IF($B{r}="","",SUMIFS(nrAT_Sess,nrAT_PID,$B{r},nrAT_Counts,"Yes",nrAT_IsPT,"Yes"))', "fx"),
        "N": ('IF($B{r}="","",SUMIFS(nrAT_MadeUpSess,nrAT_PID,$B{r},'
              'nrAT_Counts,"Yes",nrAT_IsPT,"Yes"))', "fx"),
        "O": ('IF($B{r}="","",MAX(0,$M{r}-$N{r}))', "fx"),
        "P": ('IF($B{r}="","",cfgPTCapSessions)', "fx"),
        "Q": ('IF($B{r}="","",IF($P{r}=0,0,$O{r}/$P{r}))', "fx"),
        "R": ('IF($B{r}="","",IF($Q{r}>=1,"At Cap",IF($Q{r}>=cfgAttendanceCriticalPct,'
              '"Critical",IF($Q{r}>=cfgAttendanceAdvisoryPct,"Advisory","OK"))))', "fx"),
        "S": ('IF($B{r}="","",IF($P{r}=0,"Yes",IF($O{r}<$P{r},"Yes","No")))', "fx"),
        "T": ('IF($B{r}="","",IF(AND($L{r}=0,$O{r}=0),"Yes","No"))', "fx"),
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
                   "5-session cap (policy 400). 'Cl MadeUp'/'PT MadeUp' are "
                   "the sum of the PER-EVENT credits on Attendance (P/Q), "
                   "each capped at that event's own duration — so these can "
                   "never disagree with the OPEN/CLEARED banner. 'Cl Owed' = "
                   "missed minutes not yet made up; 'Makeup Complete?' "
                   "requires both owed classroom minutes AND owed PT sessions "
                   "at 0 before graduation. Locked.")
    protect(ws)
    return ws


# --------------------------------------------------------------------------
def build_sysskills(wb):
    ws = wb.create_sheet("sysSkills")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["PID", "Cadet Name", "Status", "Categories Attempted",
                    "Qualified", "Needs Remediation", "Pending",
                    "Failed Out Cats", "FailedOut?", "Skills Elig",
                    "Firearms Avg", "Firearms Best", "CoF1 Best",
                    "CoF2 Best", "Both CoF ≥70?"])
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
        "N": ('IF($B{r}="","",LET(v,MAXIFS(nrSK_Score,nrSK_PID,$B{r},'
              'nrSK_Cat,"Firearms",nrSK_CoF,1),IF(v=0,"",v)))', "fx"),
        "O": ('IF($B{r}="","",LET(v,MAXIFS(nrSK_Score,nrSK_PID,$B{r},'
              'nrSK_Cat,"Firearms",nrSK_CoF,2),IF(v=0,"",v)))', "fx"),
        "P": ('IF($B{r}="","",IF(AND($N{r}="",$O{r}=""),"",'
              'IF(AND($N{r}<>"",$O{r}<>"",N($N{r})>=70,N($O{r})>=70),"Yes",'
              'IF(OR(AND($N{r}<>"",N($N{r})<70),AND($O{r}<>"",N($O{r})<70)),'
              '"No","(one pending)"))))', "fx"),
    })
    fill_rows(ws, FIRST, LAST, cols)
    define(wb, "nrSKfailedout", "sysSkills", f"$J${FIRST}:$J${LAST}")
    define(wb, "nrSKelig", "sysSkills", f"$K${FIRST}:$K${LAST}")
    define(wb, "nrSKfirearmsAvg", "sysSkills", f"$L${FIRST}:$L${LAST}")
    define(wb, "nrSKbothCoF", "sysSkills", f"$P${FIRST}:$P${LAST}")
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
                    "OpenChainReview", "Incidents Elig",
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
                    "F:PT", "F:Medical", "F:Certs", "F:Memos", "F:OpenTime",
                    "Flag Count", "Reasons"])
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
        # CHECK DATE counts too: a retest whose deadline could not be computed
        # is not "fine", it is a policy 300.5 clock that never started
        "L": ('IF($B{r}="","",IF(COUNTIFS(nrES_PID,$B{r},nrES_RetStat,"OVERDUE")'
              '+COUNTIFS(nrES_PID,$B{r},nrES_RetStat,"CHECK DATE")>0,1,0))', "fx"),
        "M": ('IF($B{r}="","",IF(OR(PT!$K{r}="No",PT!$AB{r}="No"),1,0))', "fx"),
        "N": ('IF($B{r}="","",IF(COUNTIFS(nrMD_PID,$B{r},nrMD_Status,'
              '"RESTRICTION EXPIRED")>0,1,0))', "fx"),
        "O": ('IF($B{r}="","",IF(Certifications!$U{r}<>"",1,0))', "fx"),
        "P": ('IF($B{r}="","",IF(COUNTIFS(nrME_PID,$B{r},nrME_Status,'
              '"OVERDUE")+COUNTIFS(nrME_PID,$B{r},nrME_Status,'
              '"CHECK DATE")>0,1,0))', "fx"),
        "Q": ('IF($B{r}="","",IF(SUMPRODUCT((nrAT_PID=$B{r})*'
              '(nrAT_Cleared="OPEN"))>0,1,0))', "fx"),
        "R": ('IF($B{r}="","",SUM($E{r}:$Q{r}))', "fx"),
        "S": ('IF($B{r}="","",IF($R{r}=0,"",TEXTJOIN("; ",TRUE,'
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
              'IF($N{r}=1,"medical restriction expired",""),'
              'IF($O{r}=1,"cert copies outstanding: "&Certifications!$U{r},""),'
              'IF($P{r}=1,"OVERDUE MEMO",""),'
              'IF($Q{r}=1,SUMPRODUCT((nrAT_PID=$B{r})*(nrAT_Cleared="OPEN"))'
              '&" uncleared missed-time event(s)",""))))', "fx"),
    })
    # E needs the row-scoped reference, not the whole named range
    cols["E"] = ('IF($B{r}="","",IF(N(sysGrades!$AC{r})>=cfgFlagConsecFails,1,0))', "fx")
    fill_rows(ws, FIRST, LAST, cols)
    define(wb, "nrFLcount", "sysFlags", f"$R${FIRST}:$R${LAST}")
    define(wb, "nrFLreasons", "sysFlags", f"$S${FIRST}:$S${LAST}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24, "S": 70})
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
                    "GraduationElig", "Blocking Issues", "Final Exam Elig",
                    "Certs", "Exams Recorded", "Skills Assessed"])
    cols = _mirror()
    cols.update({
        "E": ('IF($B{r}="","",sysGrades!$V{r})', "fx"),
        "F": ('IF($B{r}="","",sysAttendance!$K{r})', "fx"),
        "G": ('IF($B{r}="","",sysAttendance!$S{r})', "fx"),
        "H": ('IF($B{r}="","",sysSkills!$K{r})', "fx"),
        "I": ('IF($B{r}="","",sysIncidents!$K{r})', "fx"),
        "J": ('IF($B{r}="","",IF(Writing!$AT{r}="Yes","Yes","No"))', "fx"),
        "K": ('IF($B{r}="","",sysAttendance!$T{r})', "fx"),
        # AFFIRMATIVE test: only an actual pass reads "Yes". A blank PT!AB
        # (final PT never assessed) and "(rubric pending)" (cfgPTFinalMinPoints
        # still 0) are distinct blocking states, not passes.
        "L": ('IF($B{r}="","",IF(PT!$AB{r}="Yes","Yes",IF(PT!$AB{r}="No","No",'
              'IF(PT!$AB{r}="(rubric pending)","Pending","Not taken"))))', "fx"),
        "M": ('IF($B{r}="","",IF(OR(sysGrades!$P{r}>0,sysSkills!$J{r}="Yes",'
              'sysIncidents!$J{r}>0),"Yes","No"))', "fx"),
        "N": ('IF($B{r}="","",IF(AND($E{r}="Yes",$F{r}="Yes",$G{r}="Yes",'
              '$H{r}="Yes",$I{r}="Yes",$J{r}="Yes",$K{r}="Yes",$L{r}="Yes",'
              '$Q{r}="Yes",$R{r}="Yes",$S{r}="Yes",$M{r}="No"),"Yes","No"))', "fx"),
        "O": ('IF($B{r}="","",IF($N{r}="Yes","Eligible",TRIM('
              'IF($E{r}<>"Yes","Academic; ","")&'
              'IF($R{r}<>"Yes","Exams not all recorded; ","")&'
              'IF($F{r}<>"Yes","Classroom; ","")&'
              'IF($G{r}<>"Yes","PT sessions; ","")&'
              'IF($H{r}<>"Yes","Skills; ","")&'
              'IF($S{r}<>"Yes","Skills not all assessed; ","")&'
              'IF($I{r}<>"Yes","Incidents; ","")&'
              'IF($J{r}<>"Yes","Writing; ","")&'
              'IF($K{r}<>"Yes","Makeup owed; ","")&'
              'IF($L{r}="No","Final PT failed; ",'
              'IF($L{r}="Pending","Final PT rubric not set; ",'
              'IF($L{r}<>"Yes","Final PT not assessed; ","")))&'
              'IF($Q{r}<>"Yes","Certs; ","")&'
              'IF($M{r}="Yes","Dismissal review; ",""))))', "fx"),
        "P": ('IF($B{r}="","",IF(PT!$AB{r}="Yes","Yes",IF(PT!$AB{r}="No","No",'
              'IF(PT!$AB{r}="(rubric pending)","Pending","Not taken"))))', "fx"),
        "Q": ('IF($B{r}="","",IF(Certifications!$T{r}="Yes","Yes","No"))', "fx"),
        # completeness: sysGrades V waives a category whose count is 0 so
        # mid-academy cadets don't read Academic="No". Graduation must not
        # inherit that waiver — every category needs at least one record.
        "R": ('IF($B{r}="","",IF(AND(N(sysGrades!$E{r})>0,N(sysGrades!$F{r})>0,'
              'N(sysGrades!$G{r})>0,N(sysGrades!$H{r})>0),"Yes","No"))', "fx"),
        # same shape as R, for skills: sysSkills K only asks "not failed out
        # and nothing currently in remediation", so a cadet with ZERO skills
        # records scored as a pass. Graduation additionally requires every
        # SkillsMaster category to be Qualified (sysSkills F = categories
        # qualified). Deliberately NOT folded into sysSkills K / nrSKelig,
        # which stays the mid-academy "on track" label on CadetProfile.
        "S": ('IF($B{r}="","",IF(N(sysSkills!$F{r})>='
              'SUMPRODUCT((rngSM_cat<>"")*1),"Yes","No"))', "fx"),
    })
    fill_rows(ws, FIRST, LAST, cols)
    cf_yes_no(ws, f"N{FIRST}:N{LAST}")
    define(wb, "nrCKgradElig", "sysChecks", f"$N${FIRST}:$N${LAST}")
    define(wb, "nrCKblocking", "sysChecks", f"$O${FIRST}:$O${LAST}")
    define(wb, "nrCKfinalExamElig", "sysChecks", f"$P${FIRST}:$P${LAST}")
    define(wb, "nrCKexamsRecorded", "sysChecks", f"$R${FIRST}:$R${LAST}")
    define(wb, "nrCKskillsAssessed", "sysChecks", f"$S${FIRST}:$S${LAST}")
    col_widths(ws, {"A": 3, "B": 10, "C": 24, "O": 50})
    sheet_note(ws, "Graduation gate per policy: 70 in each category, every "
                   "category actually recorded (Exams Recorded), under "
                   "attendance caps, makeup complete, EVERY skills category "
                   "actually qualified (Skills Assessed — no records is not "
                   "a pass), "
                   "writing current, no open chain-of-command incident "
                   "review, final PT PASSED (a blank or '(rubric pending)' "
                   "final PT blocks — it is not a pass), all cert copies on "
                   "file, no open dismissal review. 'Final Exam Elig' "
                   "enforces 500.1.H. Locked.")
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
                   "and perseverance before finalizing. Only the override (E) "
                   "and Notes (G) columns are editable — the computed winner "
                   "and the FINAL winner feed the printed transcript.")
    # locked like every other engine sheet: C/F are formulas that reach the
    # OFFICIAL TRANSCRIPT's awards line. E and G are the reason this sheet
    # exists, so they are explicitly unlocked before protection goes on.
    protect(ws)
    from xlb import unlock_range as _unlock
    _unlock(ws, f"E{FIRST}:E{r-1}")
    _unlock(ws, f"G{FIRST}:G{r-1}")
    return ws


# --------------------------------------------------------------------------
# Module-level so the printable Audit sheet can size its mirror block from
# len(AUDIT_CHECKS) instead of a hard-coded literal — a check added here but
# not printed there is a red Dashboard tile with no visible cause.
AUDIT_CHECKS = [
        ("Delivered hours vs 736 (exact)",
         "nrCHtotalDelivered", "cfgRequiredHours",
         'IF(N(nrCHtotalDelivered)<N(cfgRequiredHours),"SHORT",'
         'IF(N(nrCHtotalDelivered)>N(cfgRequiredHours),"OVER - report excess as #101","OK"))',
         '"Report the BPOC at exactly 736 hrs; excess goes to the reporting course '
         'the Addendum sheet names - #2040 Arrest and Control, #2046 Professional '
         'Police Driving, #2055 Firearms, #101 Addendum to BPOC for the rest"'),
        # <>0 (not >0): a chapter whose delivered hours went NEGATIVE (a
        # swapped start/end time on a Schedule block) is the most broken
        # case there is and must not be skipped as "not taught yet".
        ("Chapters short of TCOLE minimum",
         'SUMPRODUCT((nrCHdeliv<>"")*(nrCHdeliv<>0)*(nrCHdeliv<nrCHmin))', "0",
         'IF(SUMPRODUCT((nrCHdeliv<>"")*(nrCHdeliv<>0)*(nrCHdeliv<nrCHmin))=0,"OK","CHECK")',
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
        # N/A is an offered answer on the Audit sheet's own dropdown and its
        # conditional format paints it green — the engine must agree, or one
        # legitimate N/A pins this row red forever with nothing to fix.
        ("Program requirements unchecked (Audit sheet)",
         'SUMPRODUCT((nrPRGitems<>"")*(nrPRGmet<>"Yes")*(nrPRGmet<>"N/A"))', "0",
         'IF(SUMPRODUCT((nrPRGitems<>"")*(nrPRGmet<>"Yes")*(nrPRGmet<>"N/A"))=0,"OK","CHECK")',
         '"Rule 215.9 distribution, TCLEDDS roster, 736 exact reporting, facility, assessments"'),
        ("Teaching instructors lacking documentation",
         'SUMPRODUCT((nrInstrOnSched="Yes")*(nrInstrReady="No"))', "0",
         'IF(SUMPRODUCT((nrInstrOnSched="Yes")*(nrInstrReady="No"))=0,"OK","CHECK")',
         '"IRG: every instructor/co-teacher teaching an LO needs bio + TCOLE cert or SME letter"'),
        ("Schedule blocks with unrecognized instructor",
         'COUNTIF(nrSCH_InstrOK,"UNRECOGNIZED")', "0",
         'IF(COUNTIF(nrSCH_InstrOK,"UNRECOGNIZED")=0,"OK","CHECK")',
         '"Instructor text matches nobody on the Instructors roster - fix spelling or add them"'),
        # a swapped start/end time is arithmetic, not a typo Excel can see:
        # the block's hours land in ChapterMaster "Delivered Hrs" and in the
        # Settings schedule-minutes detector that sets the 5% attendance cap
        ("Schedule blocks with impossible times",
         'SUMPRODUCT((nrSCH_TimeCheck<>"")*(nrSCH_TimeCheck<>"OK"))', "0",
         'IF(SUMPRODUCT((nrSCH_TimeCheck<>"")*(nrSCH_TimeCheck<>"OK"))=0,'
         '"OK","CHECK")',
         '"End time is at or before Start: the block\'s hours are wrong, '
         'which moves delivered chapter hours and the 5% attendance cap"'),
        # CHECK DATE rows are counted too: a retest dated on a weekend or a
        # holiday used to sit at "Pending" forever, so the 5-class-day clock
        # was unenforceable for exactly the rows that needed it.
        ("Overdue retests (incl. unusable dates)",
         'COUNTIF(nrES_RetStat,"OVERDUE")+COUNTIF(nrES_RetStat,"CHECK DATE")', "0",
         'IF(COUNTIF(nrES_RetStat,"OVERDUE")+COUNTIF(nrES_RetStat,"CHECK DATE")'
         '=0,"OK","ACT NOW")',
         '"Policy 300.5: retest within 5 class days. CHECK DATE = the exam date is '
         'missing or resolves past the last class day, so no deadline exists"'),
        ("Cadets with makeup owed",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrATTmakeupOK="No"))', "0",
         'IF(SUMPRODUCT((nrCadetStatus="Active")*(nrATTmakeupOK="No"))=0,"OK","CHECK")',
         '"Missed classroom minutes / PT sessions not yet made up (active cadets)"'),
        ("Active cadets failing a category",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrGRacademic="No"))', "0",
         'IF(SUMPRODUCT((nrCadetStatus="Active")*(nrGRacademic="No"))=0,"OK","CHECK")',
         '"70 in each category (policy 300.2)"'),
        ("Spelling interventions open",
         'COUNTIF(nrSpellFlag,"INTERVENTION")', "0",
         'IF(COUNTIF(nrSpellFlag,"INTERVENTION")=0,"OK","CHECK")',
         '"Document intervention on Counseling log (300.4.B)"'),
        ("Cadets missing certification copies",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrCERTmissing<>""))', "0",
         'IF(SUMPRODUCT((nrCadetStatus="Active")*(nrCERTmissing<>""))=0,"OK","CHECK")',
         '"TIM, SFST, TCIC, CPR/AED, ALERRT, ICS copies - see Certifications sheet"'),
        ("Firearms: both courses of fire not passed",
         'SUMPRODUCT((nrCadetStatus="Active")*(nrSKbothCoF="No"))', "0",
         'IF(SUMPRODUCT((nrCadetStatus="Active")*(nrSKbothCoF="No"))=0,"OK","CHECK")',
         '"IRG requires 70%+ on BOTH firearms courses of fire (ch 41)"'),
        ("Advisory board met within last 12 months",
         'IF(COUNT(nrAB_Date)=0,"none",TEXT(MAX(nrAB_Date),"mm/dd/yyyy"))', "recent",
         'IF(COUNT(nrAB_Date)=0,"CHECK",IF(MAX(nrAB_Date)>=TODAY()-366,'
         '"OK","CHECK"))',
         '"Board meets 1-2x/year; running list + minutes folder on AdvisoryBoard sheet"'),
        ("Governance alignment done for this academy",
         'cfgBoardReviewed&" / "&cfgRulesAligned', "Yes / Yes",
         'IF(AND(cfgBoardReviewed="Yes",cfgRulesAligned="Yes"),"OK","CHECK")',
         '"Board minutes reviewed + rule changes applied to this workbook (Startup Review)"'),
        ("Policy manual version recorded",
         'IF(cfgPolicyVersion="","(blank)",cfgPolicyVersion)', "set",
         'IF(cfgPolicyVersion="","CHECK","OK")',
         '"Which policy manual this academy runs under (AdvisoryBoard sheet)"'),
        ("Log rows with unrecognized cadet name",
         'COUNTIF(nrES_PID,"?")+COUNTIF(nrAT_PID,"?")+COUNTIF(nrMK_PID,"?")'
         '+COUNTIF(nrSK_PID,"?")+COUNTIF(nrIN_PID,"?")+COUNTIF(nrCO_PID,"?")'
         '+COUNTIF(nrME_PID,"?")+COUNTIF(nrMD_PID,"?")', "0",
         'IF(COUNTIF(nrES_PID,"?")+COUNTIF(nrAT_PID,"?")+COUNTIF(nrMK_PID,"?")'
         '+COUNTIF(nrSK_PID,"?")+COUNTIF(nrIN_PID,"?")+COUNTIF(nrCO_PID,"?")'
         '+COUNTIF(nrME_PID,"?")+COUNTIF(nrMD_PID,"?")=0,"OK","CHECK")',
         '"A log row\'s cadet name matches nobody on the Cadets roster - '
         'that cadet\'s scores/events are orphaned; fix the spelling"'),
        # every SUMIFS/COUNTIFS in the engine is keyed on PID, and every log
        # sheet resolves PID by MATCH on the cadet NAME, so a repeat of
        # either silently MERGES two cadets' records in both directions.
        ("Duplicate cadet PID or name on the roster",
         'SUMPRODUCT((rngCadetPIDs<>"")*(COUNTIF(rngCadetPIDs,rngCadetPIDs)>1))'
         '+SUMPRODUCT((rngCadetNames<>"")*(COUNTIF(rngCadetNames,rngCadetNames)>1))',
         "0",
         'IF(SUMPRODUCT((rngCadetPIDs<>"")*(COUNTIF(rngCadetPIDs,rngCadetPIDs)>1))'
         '+SUMPRODUCT((rngCadetNames<>"")*(COUNTIF(rngCadetNames,rngCadetNames)>1))'
         '=0,"OK","CHECK")',
         '"Two Cadets rows sharing a PID (or a name) merge both cadets\' grades, '
         'attendance and gates - and swap them onto each other\'s transcript. '
         'The duplicate cells are highlighted red on Cadets"'),
        ("Exam rows failing Row Check",
         'SUMPRODUCT((nrES_RowCheck<>"")*(nrES_RowCheck<>"OK"))', "0",
         'IF(SUMPRODUCT((nrES_RowCheck<>"")*(nrES_RowCheck<>"OK"))=0,"OK","CHECK")',
         '"ExamScores Row Check: a non-numeric or out-of-range Raw Score, a '
         'duplicate RecordID, an attempt 2 with no attempt 1 on file, or a '
         'retest row logged with no score - each silently distorts the '
         'category average or stops the 5-class-day retest clock"'),
        ("Spelling rows failing Row Check",
         'SUMPRODUCT((nrSpellRowCheck<>"")*(nrSpellRowCheck<>"OK"))', "0",
         'IF(SUMPRODUCT((nrSpellRowCheck<>"")*(nrSpellRowCheck<>"OK"))=0,'
         '"OK","CHECK")',
         '"A spelling score that is text or outside 0-100 inflates the '
         'cadet average and can hide the policy 300.4.B intervention flag"'),
        ("Memos with an unusable due date",
         'COUNTIF(nrME_Status,"CHECK DATE")', "0",
         'IF(COUNTIF(nrME_Status,"CHECK DATE")=0,"OK","CHECK")',
         '"The Assigned date resolves past the last class day, so no due date '
         'could be computed and the memo can never read OVERDUE"'),
        # NOTE: no audit row for "skills not all assessed". That condition is
        # true for every cadet for most of the academy, so an audit row would
        # sit red from day one and train the coordinator to ignore the block.
        # It is enforced per cadet instead — sysChecks S -> GradChecklist
        # ELIGIBLE + the named Blocking Issue.
    ]


def build_sysaudit(wb):
    """TCOLE audit-readiness rollups that need engine math."""
    ws = wb.create_sheet("sysAudit")
    ws.sheet_view.showGridLines = False
    header_row(ws, ["Check", "Value", "Target", "Status", "Detail"])
    r = FIRST
    checks = AUDIT_CHECKS
    for name, val_fx, target, stat_fx, detail_fx in checks:
        ws.cell(row=r, column=2, value=name).font = F_LABEL
        v = ws.cell(row=r, column=3, value="=" + val_fx)
        v.font = F_CALC
        v.border = BOX
        # cfg* targets are live formulas; digits are numbers; anything else
        # is a literal display string (never prefix those with "=")
        t = ws.cell(row=r, column=4, value="=" + target if target.startswith("cfg")
                    else (int(target) if target.isdigit() else target))
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

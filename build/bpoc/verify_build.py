"""Structural verification of the built BPOC workbook.

Run after build_bpoc.py:  python3 build/bpoc/verify_build.py
Exits non-zero on any failure. (Formula *values* need Excel; this verifies
structure: names, prefixes, seeds, protection, print areas, postprocessor.)
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from openpyxl import load_workbook

import postprocess
import data_spelling as DS
import data_writing as DW
import data_chapters as DC

WB_PATH = os.path.join(os.path.dirname(os.path.dirname(HERE)),
                       "workbooks", "BPOC_Academy_Management_V6.xlsx")

failures = []


def check(cond, msg):
    if cond:
        print(f"  ok  {msg}")
    else:
        failures.append(msg)
        print(f"FAIL  {msg}")


def test_postprocess_units():
    f = postprocess.fix_formula('=XLOOKUP(A1,B:B,C:C)')
    check(f == '=_xlfn.XLOOKUP(A1,B:B,C:C)', "postprocess: XLOOKUP prefixed")
    f = postprocess.fix_formula('=FILTER(A:A,B:B=1)')
    check(f == '=_xlfn._xlws.FILTER(A:A,B:B=1)', "postprocess: FILTER namespaced")
    f = postprocess.fix_formula('=LET(x,1,x+1)')
    check(f == '=_xlfn.LET(_xlpm.x,1,_xlpm.x+1)', "postprocess: LET params")
    f = postprocess.fix_formula('=LET(s,MAX(A:A),IF(s=0,"",LET(p,s+1,p*2)))')
    check("_xlpm.s" in f and "_xlpm.p" in f and f.count("_xlfn.LET(") == 2,
          "postprocess: nested LET")
    f = postprocess.fix_formula('=IF(A1="FILTER(","x",LET(d,1,d))')
    check('"FILTER("' in f and "_xlpm.d,1,_xlpm.d" in f,
          "postprocess: strings untouched")
    f = postprocess.fix_formula('=WORKDAY.INTL(A1,5,"0000011",B:B)')
    check(f.startswith('=_xlfn.WORKDAY.INTL('), "postprocess: WORKDAY.INTL")


def test_workbook():
    wb = load_workbook(WB_PATH)
    check(len(wb.sheetnames) == 60, f"60 sheets ({len(wb.sheetnames)} found)")

    # every referenced name is defined
    defined = set(wb.defined_names.keys())
    pat = re.compile(r"\b(nr|cfg|lst|rng)[A-Za-z_0-9]+")
    missing = set()
    unprefixed = []
    fnpat = re.compile(
        r'(?<![A-Za-z0-9_.])(XLOOKUP|FILTER|SORTBY|SORT|TEXTJOIN|MAXIFS|'
        r'MINIFS|HSTACK|VSTACK|TAKE|SEQUENCE|LET|WORKDAY\.INTL)\(')
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    for m in pat.finditer(v):
                        if m.group(0) not in defined:
                            missing.add(m.group(0))
                    if v.count("(") != v.count(")"):
                        unprefixed.append(f"paren {ws.title}!{c.coordinate}")
                    for m in fnpat.finditer(v):
                        unprefixed.append(
                            f"unprefixed {m.group(1)} {ws.title}!{c.coordinate}")
    check(not missing, f"all formula name refs defined {sorted(missing)[:5]}")
    check(not unprefixed, f"all modern fns prefixed {unprefixed[:5]}")

    # seed data
    sm = wb["SpellingMaster"]
    check(sm["C6"].value == DS.TESTS[1][0] and sm["N30"].value == DS.TESTS[12][24],
          "SpellingMaster seeded (12 tests x 25 words)")
    wm = wb["WritingMaster"]
    check(wm["B6"].value == 1 and wm["B45"].value == 40 and
          isinstance(wm["M6"].value, str) and "Code of Ethics" in wm["C6"].value,
          "WritingMaster seeded (40 assignments with prompts)")
    cm = wb["ChapterMaster"]
    check(cm["D7"].value == "Professionalism and Ethics" and
          cm["E7"].value == 12, "ChapterMaster seeded with TCOLE minimums")
    check(cm["S28"].value is not None and "TIM" in cm["S28"].value,
          "ChapterMaster special TCOLE requirements seeded (ch 22)")

    # engine protection
    for n in ("sysGrades", "sysAttendance", "sysChecks", "sysAudit",
              "ScoresGrid"):
        check(wb[n].protection.sheet, f"{n} locked")
    for n in ("Cadets", "ExamScores", "Attendance", "Counseling", "PT"):
        check(not wb[n].protection.sheet, f"{n} open for entry")

    # printables have print areas
    for n in ("SignIn", "EvalSheet", "SpellingPrint", "WritingHandout",
              "Transcript", "CadetProfile", "Ranking", "GradChecklist",
              "Audit", "Addendum", "Schedule"):
        check(bool(wb[n].print_area), f"{n} has a print area")
    ad = wb["Addendum"]
    check("#2055" in str(ad["G50"].value or "") or
          any("#2055" in str(ad.cell(row=rr, column=7).value or "")
              for rr in range(10, 60)),
          "Addendum maps Firearms excess to course #2055")

    # key policy formulas present
    es = wb["ExamScores"]
    check("cfgRetakeRecordedCap" in es["M6"].value,
          "retake cap rule in ExamScores")
    check("cfgRetestClassDays" in es["T6"].value,
          "5-class-day retest deadline in ExamScores")
    sg = wb["sysGrades"]
    check("cfgWeightMajor" in sg["M6"].value, "weighted grade in sysGrades")
    ck = wb["sysChecks"]
    check("PT!$AB" in ck["L6"].value, "final-PT gate wired into sysChecks")
    check(wb["Control"]["I6"].value.startswith("=IF")
          and "WORKDAY" in wb["Control"]["I6"].value,
          "class-day calendar generated on Control")
    cm2 = wb["ChapterMaster"]
    check("nrSCH_ChNum" in cm2["G6"].value,
          "chapter delivered-hours roll up by chapter number")
    check("nrSUBname" in wb["Schedule"]["H6"].value,
          "schedule resolves sub-classes to parent chapters")
    check("nrSUBname" in wb.defined_names and
          any(cm2.cell(row=r, column=2).value == "Crash Investigation"
              for r in range(50, 75)),
          "TPD sub-class block on ChapterMaster")
    ins = wb["Instructors"]
    check("SEARCH($B6,nrSCH_Instr)" in ins["L6"].value and
          "nrSCH_ChNum" in ins["M6"].value,
          "per-instructor on-schedule scan + chapters taught")
    check('$I6="On File"' in ins["K6"].value and
          '"Guest/Outside","N/A"' in ins["K6"].value.replace("'", '"'),
          "audit-ready requires bio; guests exempt")
    check("UNRECOGNIZED" in wb["Schedule"]["N6"].value,
          "schedule flags unrecognized instructor entries")
    check("nrBankTopics" in wb.defined_names and
          "nrBankSel" in wb.defined_names and
          any(wb["InstructorBanks"].cell(row=r, column=2).value ==
              "Arrest and Control" for r in range(6, 110)),
          "InstructorBanks sheet with per-topic rows")
    sch_dvs = [dv.formula1 for dv in wb["Schedule"].data_validations.dataValidation]
    check(any("nrBankSel" in (f or "") for f in sch_dvs),
          "schedule instructor dropdown is bank-dependent")
    cp = wb["ChapterPacket"]
    check("cfgPacketChapter" in wb.defined_names and
          any("nrInstrChTaught" in str(c.value) for row in
              cp.iter_rows(min_row=5, max_row=40) for c in row
              if isinstance(c.value, str)) and bool(cp.print_area),
          "ChapterPacket one-page training-file view")
    ex = wb["ExamSheet"]
    check("cfgGradeSheetExam" in wb.defined_names and bool(ex.print_area) and
          any("nrES_Raw" in str(c.value) for row in
              ex.iter_rows(min_row=10, max_row=12) for c in row
              if isinstance(c.value, str)),
          "ExamSheet per-assessment grade sheet")
    check(wb["ExamScores"].freeze_panes == "D6" and
          wb["Schedule"].freeze_panes == "B6" and
          wb["ExamScores"].auto_filter.ref == "B5:W5",
          "freeze panes + log filters")
    check(any("nrSCH_Date=TODAY()" in str(c.value) for row in
              wb["Dashboard"].iter_rows(min_row=5, max_row=25) for c in row
              if isinstance(c.value, str)),
          "Dashboard Today panel")
    wmG = wm["G6"].value
    check("nrCHfirst" in wmG and "nrCDdate" in wmG,
          "writing due dates computed from schedule")

    cert = wb["Certifications"]
    check("TEXTJOIN" in str(cert["U6"].value) and
          cert["D5"].value == "TIM Date",
          "Certifications grid with to-collect rollup")
    check("Certifications!$T" in wb["sysChecks"]["Q6"].value and
          '$Q6="Yes"' not in wb["sysChecks"]["N6"].value,
          "certs are informational on sysChecks, not a grad gate")
    check("Certifications!$U" in wb["sysFlags"]["O6"].value and
          "cert copies outstanding" in wb["sysFlags"]["S6"].value,
          "cert warning flag in sysFlags with reason text")
    wr = wb["Writing"]
    check('COUNTIF(D6:AQ6,"X")' in wr["AR6"].value and
          '(D6:AQ6="")' in wr["AS6"].value,
          "Writing grid counts X marks, blank = not done")
    check(wr["D6"].alignment.horizontal == "center",
          "Writing X cells centered")
    check("InputGuide" in wb.sheetnames and
          "HYPERLINK" in str(wb["InputGuide"]["B7"].value),
          "InputGuide page with hyperlinks")
    check("HYPERLINK" in str(wb["Cadets"]["B1"].value) and
          "HYPERLINK" in str(wb["sysGrades"]["B1"].value),
          "home links back to Dashboard on sheets")
    dash_text = [str(c.value) for row in
                 wb["Dashboard"].iter_rows(min_row=5, max_row=12)
                 for c in row if isinstance(c.value, str)]
    check(any("Final Test" in t for t in dash_text) and
          any("State Test" in t for t in dash_text) and
          any("Graduation" in t for t in dash_text),
          "Dashboard date tiles (start/final/state/graduation)")
    check("nrSK_CoF" in wb["sysSkills"]["N6"].value,
          "firearms course-of-fire bests in sysSkills")
    se = wb["StateExam"]
    check("new BPOC required" in se["K6"].value,
          "state exam 3-attempt rule")
    check("nrCERTmissing" in str(wb["Dashboard"]["B18"].value or "") or
          any("nrCERTmissing" in str(c.value)
              for row in wb["Dashboard"].iter_rows(min_row=5, max_row=40)
              for c in row if isinstance(c.value, str)),
          "cert reminders on Dashboard")

    at = wb["Attendance"]
    check("nrMK_Link" in at["P6"].value and "CLEARED" in at["S6"].value,
          "per-event makeup reconciliation with CLEARED status")
    mk_dvs = [dv.formula1 for dv in wb["Makeup"].data_validations.dataValidation]
    check(any("nrAT_ID" in (f or "") for f in mk_dvs),
          "Makeup Linked Event dropdown = attendance EventIDs")
    me = wb["Memos"]
    check("cfgMemoDueClassDays" in me["H6"].value and
          "OVERDUE" in me["L6"].value,
          "Memos: computed due dates + overdue status")
    me_dvs = [dv.formula1 for dv in wb["Memos"].data_validations.dataValidation]
    check(any("nrAllRefIDs" in (f or "") for f in me_dvs),
          "Memo Linked Ref dropdown of I/A/C IDs")
    check("nrME_Status" in wb["sysFlags"]["P6"].value and
          'nrAT_Cleared="OPEN"' in wb["sysFlags"]["Q6"].value,
          "overdue-memo and open-time warning flags")
    check('nrIN_Report="Yes"' in str(wb["EmailPreview"]["B65"].value or "") or
          any('nrIN_Report="Yes"' in str(c.value) for row in
              wb["EmailPreview"].iter_rows(min_row=55, max_row=75)
              for c in row if isinstance(c.value, str)),
          "email digest filters on Report-to-Agency marks")
    dl = wb["DailyLog"]
    check("nrCDdate" in dl["C6"].value and "nrME_Received" in dl["K6"].value,
          "DailyLog computes day #, class type and counters")

    check("nrAB_Date" in wb.defined_names and
          "cfgPolicyVersion" in wb.defined_names and
          "cfgBoardReviewed" in wb.defined_names and
          wb["AdvisoryBoard"]["D6"].value == "May 2026",
          "AdvisoryBoard governance alignment + meeting reference list")
    aud = wb["Audit"]
    found_ack = any("Rules Ack" == str(c.value) for row in
                    aud.iter_rows(min_row=5, max_row=60) for c in row)
    check(found_ack, "Rules Ack column in enrollment docs grid")

    check(wb.calculation.fullCalcOnLoad, "fullCalcOnLoad set")


def main():
    print("== postprocess unit checks ==")
    test_postprocess_units()
    print("== workbook structure ==")
    test_workbook()
    print()
    if failures:
        print(f"{len(failures)} FAILURE(S)")
        sys.exit(1)
    print("All checks passed.")


if __name__ == "__main__":
    main()

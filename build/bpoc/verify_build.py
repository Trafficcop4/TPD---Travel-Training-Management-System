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
    check(len(wb.sheetnames) == 51, f"51 sheets ({len(wb.sheetnames)} found)")

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
    wmG = wm["G6"].value
    check("nrCHfirst" in wmG and "nrCDdate" in wmG,
          "writing due dates computed from schedule")

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

"""Build the BPOC Academy Management Workbook (V6).

Usage:  python3 build/bpoc/build_bpoc.py
Output: workbooks/BPOC_Academy_Management_V6.xlsx
        (VBA installed separately via tools/Install-VBA.ps1 -> .xlsm)
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from openpyxl import Workbook

import sheets_config
import sheets_inputs
import sheets_engine
import sheets_outputs
import postprocess

TAB_INPUT = "2E75B6"     # blue tabs = data entry
TAB_CONFIG = "C8A24B"    # gold = setup/masters
TAB_OUTPUT = "1E7145"    # green = outputs/printables
TAB_SYS = "808080"       # gray = locked engine

TAB_COLORS = {
    TAB_CONFIG: ["StartHere", "InputGuide", "Settings", "Lists", "Agencies",
                 "Instructors", "InstructorBanks",
                 "ChapterMaster", "ExamMaster", "ExamPlan", "SkillsMaster",
                 "SpellingMaster", "WritingMaster", "Control", "Schedule"],
    TAB_INPUT: ["Cadets", "ExamScores", "Spelling", "Attendance", "Makeup",
                "Skills", "SkillsCheck", "Writing", "Incidents",
                "Counseling", "PT",
                "Medical", "Certifications", "StateExam", "Memos", "DailyLog",
                "AdvisoryBoard", "DismissalLog", "EmailLog"],
    TAB_OUTPUT: ["Dashboard", "ScoresGrid", "Ranking", "WatchList",
                 "CadetProfile", "Transcript", "GradChecklist", "Audit",
                 "Addendum", "ChapterPacket", "ExamSheet", "SignIn",
                 "EvalSheet", "SpellingPrint", "WritingHandout",
                 "EmailPreview", "PrintCenter", "NamedRanges"],
    TAB_SYS: ["sysGrades", "sysAttendance", "sysSkills", "sysIncidents",
              "sysFlags", "sysChecks", "sysAwards", "sysAudit",
              "sysListsHelper"],
}

SHEET_ORDER = [
    "Dashboard", "StartHere", "InputGuide", "PrintCenter",
    "Cadets", "ExamScores", "Spelling", "Attendance", "Makeup", "Skills",
    "SkillsCheck",
    "Writing", "Incidents", "Counseling", "Memos", "DailyLog", "PT",
    "Medical", "Certifications", "StateExam", "AdvisoryBoard", "DismissalLog",
    "ScoresGrid", "Ranking", "WatchList", "CadetProfile", "Transcript",
    "GradChecklist", "Audit", "Addendum", "ChapterPacket", "ExamSheet",
    "SignIn", "EvalSheet", "SpellingPrint", "WritingHandout",
    "EmailPreview", "EmailLog",
    "Settings", "Lists", "Agencies", "Instructors", "InstructorBanks",
    "ChapterMaster",
    "ExamMaster", "ExamPlan", "SkillsMaster", "SpellingMaster",
    "WritingMaster", "Control", "Schedule",
    "sysGrades", "sysAttendance", "sysSkills", "sysIncidents", "sysFlags",
    "sysChecks", "sysAwards", "sysAudit", "sysListsHelper", "NamedRanges",
]


def build():
    wb = Workbook()
    wb.remove(wb.active)

    sheets_config.build_all_config(wb)
    sheets_inputs.build_all_inputs(wb)
    sheets_engine.build_all_engine(wb)
    sheets_outputs.build_all_outputs(wb)

    for color, names in TAB_COLORS.items():
        for n in names:
            if n in wb.sheetnames:
                wb[n].sheet_properties.tabColor = color

    order = [n for n in SHEET_ORDER if n in wb.sheetnames]
    order += [n for n in wb.sheetnames if n not in order]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0

    # daily-use ergonomics: frozen headers + filterable logs
    FREEZE = {
        "Cadets": "D6", "ExamScores": "D6", "Spelling": "D6", "Writing": "D6",
        "PT": "D6", "Certifications": "D6", "ScoresGrid": "D6",
        "StateExam": "D6", "GradChecklist": "D6", "sysGrades": "D6",
        "sysAttendance": "D6", "sysSkills": "D6", "sysIncidents": "D6",
        "sysFlags": "D6", "sysChecks": "D6",
        "Attendance": "B6", "Makeup": "B6", "Skills": "B6", "Incidents": "B6",
        "SkillsCheck": "D6",
        "Counseling": "B6", "Medical": "B6", "DismissalLog": "B6",
        "EmailLog": "B6", "Schedule": "B6", "Agencies": "B6",
        "Memos": "B6", "DailyLog": "B6", "AdvisoryBoard": "B12",
        "Instructors": "C6", "InstructorBanks": "C6", "ChapterMaster": "E6",
        "WritingMaster": "D6", "SpellingMaster": "C6", "ExamMaster": "B6",
        "ExamPlan": "B6",
        # Long scrollable lists that lost their header row on scroll. The
        # one-page printables (Transcript, CadetProfile, EvalSheet, SignIn,
        # SpellingPrint, WritingHandout, ChapterPacket, ExamSheet) are
        # deliberately NOT frozen - they are forms, not lists, and freezing
        # a form's top is just a stripe across the page.
        "InputGuide": "B6", "Ranking": "B6", "WatchList": "B6",
        "Audit": "B6", "Addendum": "B6", "EmailPreview": "B6",
        "Settings": "B6", "Lists": "B6", "SkillsMaster": "B6",
        "Control": "B6", "sysAudit": "B6", "NamedRanges": "B6",
    }
    for name, cell in FREEZE.items():
        if name in wb.sheetnames:
            wb[name].freeze_panes = cell
    FILTERS = {
        "ExamScores": "B5:Y5", "Attendance": "B5:T5", "Makeup": "B5:N5",
        "Skills": "B5:R5", "Incidents": "B5:O5", "Counseling": "B5:N5",
        "Medical": "B5:N5", "Schedule": "B5:P5", "DismissalLog": "B5:P5",
        "EmailLog": "B5:I5", "Memos": "B5:N5", "DailyLog": "B5:N5",
    }
    for name, ref in FILTERS.items():
        if name in wb.sheetnames:
            wb[name].auto_filter.ref = ref

    # a dynamic-array panel must never answer "nothing to report" because
    # something inside it broke: the all-clear message moves into FILTER's
    # own if_empty argument and the outer IFERROR keeps a diagnostic.
    hardened = postprocess.harden_workbook(wb)
    print(f"Hardened {hardened} dynamic-array safety panels")

    # Daily-use ergonomics, the biggest one: on every sheet the coordinator
    # TYPES into, lock the calculated cells and leave only the inputs
    # editable. Tab and Enter then walk the input cells alone - on Writing
    # that is the 40 assignment columns rather than all 45 - and a formula
    # column cannot be typed over by accident. The lock is derived from the
    # blue input font, the same signal the colour key already promises, so
    # "blue = you type here" and "the cursor stops here" cannot drift apart.
    # Locked cells stay selectable and copyable.
    TYPED_SHEETS = [
        "Cadets", "ExamScores", "Spelling", "Attendance", "Makeup", "Skills",
        "SkillsCheck", "Writing", "Incidents", "Counseling", "Memos",
        "DailyLog", "PT", "Medical", "Certifications", "StateExam",
        "AdvisoryBoard", "DismissalLog", "EmailLog",
        "Settings", "Lists", "Agencies", "Instructors", "InstructorBanks",
        "ChapterMaster", "ExamMaster", "ExamPlan", "SkillsMaster",
        "SpellingMaster", "WritingMaster", "Control", "Schedule",
    ]
    from xlb import protect_inputs
    locked_sheets, open_cells = 0, 0
    for _n in TYPED_SHEETS:
        if _n in wb.sheetnames:
            open_cells += protect_inputs(wb[_n])
            locked_sheets += 1
    print(f"Protected {locked_sheets} typed sheets "
          f"({open_cells:,} cells left editable)")

    # Collapse the columns that are pure REFERENCE - values echoed from a
    # master sheet so the row reads on its own - into an outline the
    # coordinator can expand with the + button above column A.
    #
    # Deliberately NOT collapsed: the calculated columns that carry a
    # VERDICT. Writing's "Overdue Missing" / "Writing Current?",
    # SkillsCheck's "Failed" / "Skills P/F OK?", Certifications' "To
    # Collect", Spelling's "Intervention?", Attendance's makeup ledger and
    # the Schedule's three check columns are the whole reason those sheets
    # compute anything. Hiding them by default would bury the answer the
    # sheet exists to give, which is the opposite of what the rest of this
    # workbook tries to do.
    REFERENCE_BLOCKS = {
        # echoed from ExamPlan: name, type, sequence, passing score
        "ExamScores": [("G", "J")],
        # echoed from SkillsMaster: max attempts, scoring mode, passing
        "Skills": [("F", "H")],
        # per-event PT points. The rollups that decide anything - Final
        # Points, Final PT Pass?, Improvement Index - stay visible at AA:AC.
        "PT": [("T", "Z")],
    }
    grouped = 0
    for _n, _blocks in REFERENCE_BLOCKS.items():
        if _n not in wb.sheetnames:
            continue
        for _a, _b in _blocks:
            wb[_n].column_dimensions.group(_a, _b, outline_level=1,
                                           hidden=True)
            grouped += 1
    print(f"Collapsed {grouped} reference-column blocks "
          f"(verdict columns left visible)")

    # store modern functions with _xlfn/_xlws/_xlpm prefixes so Excel
    # resolves them instead of showing #NAME?
    fixed = postprocess.fix_workbook(wb)
    print(f"Prefixed {fixed} formulas for Excel compatibility")

    # force full recalculation on open so every formula evaluates
    wb.calculation.fullCalcOnLoad = True

    out_dir = os.path.join(os.path.dirname(os.path.dirname(HERE)), "workbooks")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "BPOC_Academy_Management_V6.xlsx")
    wb.save(out)
    print(f"Saved {out}")
    print(f"Sheets: {len(wb.sheetnames)}  Named ranges: {len(wb.defined_names)}")
    return out


if __name__ == "__main__":
    build()

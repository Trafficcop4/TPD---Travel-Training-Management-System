"""Produce a workbook ready for a NEW academy, from the clean V6 template.

The template already ships the things that persist between classes - the
TCOLE chapters, exam/spelling/writing masters and the instructor roster - so
this only has to set the academy's own configuration and carry across the
per-agency contact list, which lives in the coordinator's workbook rather
than in this repo's data modules.

It deliberately does NOT copy cadets, scores, attendance, incidents or any
other per-class record: that is the point of starting a new academy.

Usage:
    python3 build/bpoc/start_academy.py --start 2027-02-01 --end 2027-07-30 \
        --label BPOC-2027-01 [--agencies-from <workbook.xlsm|.xlsx>] \
        [--out workbooks/BPOC_Academy_Management_V6_2027.xlsx]
"""
import argparse
import datetime as dt
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(HERE))
TEMPLATE = os.path.join(ROOT, "workbooks", "BPOC_Academy_Management_V6.xlsx")


def find_setting_row(ws, internal_name):
    for r in range(5, 60):
        if str(ws.cell(row=r, column=5).value or "").strip() == internal_name:
            return r
    return None


def set_setting(ws, name, value):
    r = find_setting_row(ws, name)
    if r is None:
        print(f"  WARNING: setting {name} not found")
        return
    ws.cell(row=r, column=3).value = value


def copy_agencies(dst, src_path):
    """Agency codes, names, contacts, emails, phones, addresses."""
    src = load_workbook(src_path, data_only=True, keep_vba=False)
    if "Agencies" not in src.sheetnames:
        print("  WARNING: source has no Agencies sheet")
        return 0
    s, d = src["Agencies"], dst["Agencies"]
    # find the source header row by its AgencyID column
    hdr = None
    for r in range(1, 12):
        for c in range(1, 8):
            if str(s.cell(row=r, column=c).value or "").strip() in (
                    "AgencyID", "Agency Code", "Code"):
                hdr, c0 = r, c
                break
        if hdr:
            break
    if hdr is None:
        print("  WARNING: could not find the source Agencies header")
        return 0
    out = 6
    n = 0
    for r in range(hdr + 1, s.max_row + 1):
        code = s.cell(row=r, column=c0).value
        if code in (None, ""):
            continue
        for off in range(0, 6):
            d.cell(row=out, column=2 + off).value = s.cell(
                row=r, column=c0 + off).value
        out += 1
        n += 1
    return n


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--start", required=True, help="YYYY-MM-DD")
    p.add_argument("--end", required=True, help="YYYY-MM-DD")
    p.add_argument("--label", required=True)
    p.add_argument("--agencies-from")
    p.add_argument("--out")
    a = p.parse_args()

    start = dt.datetime.strptime(a.start, "%Y-%m-%d").date()
    end = dt.datetime.strptime(a.end, "%Y-%m-%d").date()
    if end <= start:
        sys.exit("end date must be after the start date")

    wb = load_workbook(TEMPLATE)
    st = wb["Settings"]
    set_setting(st, "cfgAcademyClass", a.label)
    set_setting(st, "cfgStartDate", start)
    set_setting(st, "cfgEndDate", end)
    # a new class has sent no agency emails yet
    set_setting(st, "cfgCurrentExamNum", 1)
    set_setting(st, "cfgCurrentSpellingNum", 1)
    print(f"Academy: {a.label}  {start} -> {end} "
          f"({(end - start).days} days, {(end - start).days / 7:.1f} weeks)")

    n = 0
    if a.agencies_from:
        n = copy_agencies(wb, a.agencies_from)
    print(f"Agencies carried over: {n}")
    print("Carried by the template itself: TCOLE chapters, exam/spelling/"
          "writing masters, instructor roster")
    print("EMPTY and yours to fill: Schedule (build it on the Schedule "
          "sheet), InstructorBanks, Cadets and every per-class record")

    out = a.out or os.path.join(
        ROOT, "workbooks",
        f"BPOC_Academy_Management_V6_{a.label.replace('/', '-')}.xlsx")
    wb.save(out)
    print(f"Saved {out}")


if __name__ == "__main__":
    main()

"""Rewrite formulas for on-disk Excel compatibility.

Functions newer than Excel 2007 must be stored with an `_xlfn.` prefix
(dynamic-array functions SORT/FILTER additionally use `_xlfn._xlws.`), and
LET/LAMBDA parameter names must be stored as `_xlpm.<name>`. Excel restores
the friendly names on load; without the prefixes every such formula opens as
#NAME?.
"""
import re

# function -> stored prefix (dot included)
XLFN = {
    "XLOOKUP": "_xlfn.", "XMATCH": "_xlfn.", "LET": "_xlfn.",
    "LAMBDA": "_xlfn.", "TEXTJOIN": "_xlfn.", "CONCAT": "_xlfn.",
    "IFS": "_xlfn.", "SWITCH": "_xlfn.", "MAXIFS": "_xlfn.",
    "MINIFS": "_xlfn.", "SEQUENCE": "_xlfn.", "UNIQUE": "_xlfn.",
    "SORTBY": "_xlfn.", "RANDARRAY": "_xlfn.", "HSTACK": "_xlfn.",
    "VSTACK": "_xlfn.", "TAKE": "_xlfn.", "DROP": "_xlfn.",
    "TOCOL": "_xlfn.", "TOROW": "_xlfn.", "ISOWEEKNUM": "_xlfn.",
    "SORT": "_xlfn._xlws.", "FILTER": "_xlfn._xlws.",
}
# NOTE: WORKDAY.INTL / NETWORKDAYS.INTL must NOT be listed above. They are
# native ECMA-376 built-ins (openpyxl.utils.FORMULAE lists both), so an
# _xlfn. prefix makes Excel open them as #NAME?. Only functions that are
# ABSENT from that catalog are "future functions" needing a prefix; the
# assertion in verify_build.py enforces exactly that.

_FUNC_PAT = re.compile(
    r"(?<![A-Za-z0-9_.])(" +
    "|".join(sorted((re.escape(k) for k in XLFN), key=len, reverse=True)) +
    r")\(")


def _prefix_functions(f):
    out = []
    i = 0
    in_str = False
    while i < len(f):
        ch = f[i]
        if ch == '"':
            in_str = not in_str
            out.append(ch)
            i += 1
            continue
        if not in_str:
            m = _FUNC_PAT.match(f, i)
            if m and (i == 0 or f[i - 1] not in "._" ):
                name = m.group(1)
                out.append(XLFN[name] + name + "(")
                i = m.end()
                continue
        out.append(ch)
        i += 1
    return "".join(out)


def _find_matching(f, start):
    """start = index of '('; return index of matching ')'."""
    depth = 0
    in_str = False
    for i in range(start, len(f)):
        ch = f[i]
        if ch == '"':
            in_str = not in_str
        elif not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return i
    return -1


def _split_top(f):
    """Split argument string on top-level commas (string- and paren-aware)."""
    args, depth, in_str, cur = [], 0, False, []
    for ch in f:
        if ch == '"':
            in_str = not in_str
        if not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                args.append("".join(cur))
                cur = []
                continue
        cur.append(ch)
    args.append("".join(cur))
    return args


def _prefix_var(text, var):
    """Prefix standalone `var` tokens with _xlpm., skipping string literals
    and tokens already prefixed / part of refs or function names."""
    pat = re.compile(r'(?<![A-Za-z0-9_.$])' + re.escape(var) +
                     r'(?![A-Za-z0-9_.!(])')
    out = []
    parts = text.split('"')
    for i, part in enumerate(parts):
        if i % 2 == 0:  # outside string
            out.append(pat.sub("_xlpm." + var, part))
        else:
            out.append(part)
    return '"'.join(out)


def _transform_lets(f):
    """Prefix LET parameter names with _xlpm. (handles nesting)."""
    key = "_xlfn.LET("
    pos = f.find(key)
    while pos != -1:
        open_i = pos + len(key) - 1
        close_i = _find_matching(f, open_i)
        if close_i == -1:
            break
        inner = f[open_i + 1:close_i]
        inner = _transform_lets(inner)          # nested LETs first
        args = _split_top(inner)
        names = [args[i].strip() for i in range(0, len(args) - 1, 2)]
        names = [n for n in names if re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", n)]
        rebuilt = list(args)
        for ai in range(len(rebuilt)):
            is_name_slot = (ai % 2 == 0 and ai < len(rebuilt) - 1)
            if is_name_slot and rebuilt[ai].strip() in names:
                rebuilt[ai] = "_xlpm." + rebuilt[ai].strip()
            else:
                for n in names:
                    rebuilt[ai] = _prefix_var(rebuilt[ai], n)
        f = f[:open_i + 1] + ",".join(rebuilt) + f[close_i:]
        # recompute close and continue after this LET
        pos = f.find(key, open_i + 1)
    return f


def fix_formula(f):
    if not isinstance(f, str) or not f.startswith("="):
        return f
    out = _prefix_functions(f)
    if "_xlfn.LET(" in out:
        out = _transform_lets(out)
    return out


def fix_workbook(wb):
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    nv = fix_formula(v)
                    if nv != v:
                        c.value = nv
                        n += 1
    return n


# ---------------------------------------------------------------------------
# dynamic-array panel hardening
#
# Every safety-net panel in the workbook was written as
#     =IFERROR(TAKE(FILTER(...)),"No flags — clear")
# so ANY error inside the array — the #REF! a stale cross-sheet column letter
# leaves behind, #VALUE!, #NAME?, a broken named range — rendered as an
# affirmative ALL-CLEAR on the very panels a coordinator reads to find out
# whether anything is wrong. FILTER has a native `if_empty` argument that
# fires ONLY on a genuinely empty result; the message belongs there. The
# outer IFERROR stays as a last-resort net, but its fallback must never
# assert that everything is fine.
PANEL_DIAG = ('"⚠ LIST UNAVAILABLE - a reference in this panel is '
              'broken; do NOT read it as \'nothing to report\'"')


def _outside_string(f, idx):
    return f.count('"', 0, idx) % 2 == 0


def _add_if_empty(expr, msg):
    """Give every 2-argument FILTER in `expr` the panel's message as its
    native if_empty. SORT/SORTBY wrappers take a second FILTER as the
    by-array; that one needs the argument too, or a legitimately empty list
    returns #CALC! from the by-array and the wrapper propagates it."""
    out, i = expr, 0
    while True:
        j = out.find("FILTER(", i)
        if j == -1:
            return out
        if (j > 0 and (out[j - 1].isalnum() or out[j - 1] in "._")) or \
                not _outside_string(out, j):
            i = j + 7
            continue
        open_p = j + len("FILTER")
        close_p = _find_matching(out, open_p)
        if close_p == -1:
            return out
        if len(_split_top(out[open_p + 1:close_p])) == 2:
            out = out[:close_p] + "," + msg + out[close_p:]
            i = close_p + len(msg) + 2
        else:
            i = close_p


def harden_formula(f):
    if not isinstance(f, str) or "FILTER(" not in f or "IFERROR(" not in f:
        return f
    i = 0
    while True:
        j = f.find("IFERROR(", i)
        if j == -1:
            return f
        if not _outside_string(f, j):
            i = j + 8
            continue
        open_p = j + len("IFERROR")
        close_p = _find_matching(f, open_p)
        if close_p == -1:
            return f
        args = _split_top(f[open_p + 1:close_p])
        msg = args[-1].strip() if len(args) == 2 else ""
        # only a NON-EMPTY message is a claim about the data. An IFERROR
        # falling back to "" (a helper column, a picker source list) says
        # nothing to the reader and must be left alone.
        if (len(args) == 2 and "FILTER(" in args[0] and
                len(msg) >= 3 and msg[0] == '"' and msg[-1] == '"'):
            # a panel whose FILTERs ALREADY carry if_empty still needs its
            # outer fallback demoted - that is where the false all-clear
            # actually gets printed.
            new = _add_if_empty(args[0], msg)
            f = f[:open_p + 1] + new + "," + PANEL_DIAG + f[close_p:]
            i = open_p + 1
            continue
        i = j + 8


def harden_workbook(wb):
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    nv = harden_formula(v)
                    if nv != v:
                        c.value = nv
                        n += 1
    return n

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
    "WORKDAY.INTL": "_xlfn.", "NETWORKDAYS.INTL": "_xlfn.",
    "SORT": "_xlfn._xlws.", "FILTER": "_xlfn._xlws.",
}

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

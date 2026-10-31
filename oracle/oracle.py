"""The LibreOffice oracle's two steps that are not running programs: composing the workbook abax
writes, and comparing what abax and LibreOffice made of it.

    python3 -I oracle.py compose <data.tsv> <cases.txt> <book.tsv>
    python3 -I oracle.py compare <data.tsv> <cases.txt> <abax dump> <LibreOffice's .xlsx, dumped>
                                 <LibreOffice's CSV> <LibreOffice version> <tests_oracle.sysl>

LibreOffice's values are the ones it caches in the .xlsx it writes after recalculating abax's: those
are raw numbers and Excel's error codes, which is what Excel would read, where its CSV export shows
cells formatted (50% as "50%") and its own error numbers (Err:502). The CSV export is read too, for
the numbers and logicals in it: where it differs from LibreOffice's own .xlsx, the test says so.

`oracle.zsh` runs both, with abax-cli and soffice between them. `compare` prints every disagreement
and writes the agreed values -- and the documented differences, with what abax gives -- as sysl
tests, so `sysl test .` checks them without LibreOffice.
"""

import sys

COLUMN = 9  # J: where the formulas go, clear of the data in A:H

# Where LibreOffice and abax disagree and abax is right by Excel's documented behaviour: the formula,
# what abax (and Excel) give, as dump text, and why. Each one is checked against abax's value, and
# LibreOffice's is named in the test.
LOGICAL = "Excel's logicals are a kind of their own; LibreOffice's TRUE is the number 1"
IN_RANGE = "Excel's aggregates skip a logical in a range, as they skip text; LibreOffice counts A3's TRUE as 1"

DIFFERENCES = {
    "=TRUE>1": ("TRUE", "Excel orders every logical after every number; " + LOGICAL),
    "=(-8)^(1/3)": ("#NUM!", "Excel refuses a negative number to a fractional power; LibreOffice takes the odd root"),
    "=0^0": ("#NUM!", "Excel documents #NUM! for 0^0"),
    "=A3&\"\"": ("'TRUE", "Excel's TRUE joins as the word; " + LOGICAL),
    "=\"x\"&TRUE": ("xTRUE", "Excel's TRUE joins as the word; " + LOGICAL),
    "=CONCATENATE(1,TRUE)": ("1TRUE", "Excel's TRUE joins as the word; " + LOGICAL),
    "=1E+20&\"\"": ("'1E+20", "Excel writes a two-digit exponent; LibreOffice writes three"),
    "=CONCAT(A1,A3,1.5)": ("1TRUE1.5", "Excel's TRUE joins as the word; " + LOGICAL),
    "=#N/A+1/0": ("#N/A", "Excel answers the first error from the left; LibreOffice the right operand's"),
    "=SUM(A1:A7)": ("2.5", IN_RANGE),
    "=PRODUCT(A2:A4)": ("0", IN_RANGE + ", and Excel's PRODUCT of no numbers is 0"),
    "=AVERAGE(A1:A7)": ("0.8333333333333334", IN_RANGE),
    "=AVERAGE(A2:A4)": ("#DIV/0!", IN_RANGE + ", and an AVERAGE of no numbers is #DIV/0!"),
    "=COUNT(A1:A8)": ("3", IN_RANGE),
    "=VLOOKUP(3,F1:H5,4)": ("#REF!", "Excel documents #REF! for a column past the table's last"),
    "=SUM(1,\"2\",TRUE)": ("4", "Excel reads text typed as an argument as a number; LibreOffice's default "
                           "refuses to convert text"),
    "=POWER(0,-1)": ("#DIV/0!", "Excel documents #DIV/0! for zero to a negative power"),
    "=SQRT(-1)": ("#NUM!", "Excel documents #NUM!; LibreOffice's own Err:502 is cached as #VALUE!"),
    "=#REF!": ("#REF!", "LibreOffice reads a formula that is only #REF! as an unknown name"),
    "=ISERROR(FOO(1))": ("TRUE", "an unknown function inside a call is a syntax error (Err:509) to LibreOffice"),
}

# Relative tolerance for numbers. Both programs cache the shortest round-trip text, so most agree
# exactly; this allows for the last bits of a sum or a power worked out in another order.
TOLERANCE = 1e-12


def cases(path):
    """The formulas in order, each with the section it is under."""
    out = []
    section = "operators"
    for line in open(path, encoding="utf-8"):
        line = line.rstrip("\n")
        if line.startswith("# ---"):
            section = line[5:].strip()
        elif line.strip() and not line.startswith("#"):
            out.append((section, line))
    return out


def data_rows(path):
    return [line.rstrip("\n").split("\t") for line in open(path, encoding="utf-8")]


def compose(data, case_file, out):
    rows = data_rows(data)
    formulas = [f for _, f in cases(case_file)]
    n = max(len(rows), len(formulas))
    with open(out, "w", encoding="utf-8") as f:
        for i in range(n):
            fields = rows[i] if i < len(rows) else []
            fields = fields + [""] * (COLUMN - len(fields))
            if i < len(formulas):
                fields = fields[:COLUMN] + [formulas[i]]
            f.write("\t".join(fields).rstrip("\t") + "\n")


def unplain(t):
    out, i = [], 0
    while i < len(t):
        if t[i] == "\\" and i + 1 < len(t):
            out.append({"t": "\t", "n": "\n", "r": "\r", "\\": "\\"}[t[i + 1]])
            i += 2
        else:
            out.append(t[i])
            i += 1
    return "".join(out)


def dump(path):
    cells = {}
    for line in open(path, encoding="utf-8"):
        parts = line.rstrip("\n").split("\t")
        if len(parts) == 3:
            cells[parts[0]] = (unplain(parts[1]), unplain(parts[2]))
    return cells


def agrees(theirs, ours):
    if theirs[0] != ours[0]:
        return False
    if theirs[0] == "number":
        x, y = theirs[1], ours[1]
        return x == y or abs(x - y) <= TOLERANCE * max(abs(x), abs(y))
    return theirs[1] == ours[1]


def address(i):
    return "J" + str(i + 1)


def sysl_string(t):
    return '"' + t.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\t", "\\t") + '"'


def sysl_value(kind, v):
    if kind == "number":
        r = repr(float(v))
        if "e" not in r and "." not in r and "inf" not in r:
            r += ".0"
        return "Number(" + r + ")"
    if kind == "bool":
        return "Bool(" + ("true" if v == "TRUE" else "false") + ")"
    if kind == "error":
        return "Error(fault_named(" + sysl_string(v) + ").unwrap())"
    if kind == "empty":
        return "Empty"
    return "Text(" + sysl_string(v) + ")"


ERRORS = ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#CIRC!")


def dump_kind(text):
    """A value as abax-cli dump writes it -- as it would be typed, text with a ' where it would read
    as something else -- as (kind, value)."""
    if text == "":
        return ("empty", "")
    if text.startswith("'"):
        return ("text", text[1:])
    if text in ("TRUE", "FALSE"):
        return ("bool", text)
    if text in ERRORS:
        return ("error", text)
    try:
        return ("number", float(text))
    except ValueError:
        return ("text", text)


def same_input(mine, theirs):
    """Whether LibreOffice's .xlsx gives back the input abax wrote. LibreOffice writes a logical as the
    formula TRUE() or FALSE() and an error constant as a formula of that error, whether it was typed
    as a constant or inside a formula, so those spellings are taken as the same."""
    def norm(t):
        t = t.replace("TRUE()", "TRUE").replace("FALSE()", "FALSE")
        return t[1:] if t[1:] in ("TRUE", "FALSE") or t[1:] in ERRORS else t
    return norm(mine) == norm(theirs)


def csv_view(path):
    """LibreOffice's CSV export of the same workbook, as {address: text}. The export writes cells as
    formatted, so it is only compared where its text can be read unambiguously (see `shows`)."""
    import csv
    out = {}
    with open(path, encoding="utf-8", newline="") as f:
        for r, row in enumerate(csv.reader(f)):
            if COLUMN < len(row):
                out[address(r)] = row[COLUMN]
    return out


def shows(text, theirs):
    """Whether the CSV export's text is the value LibreOffice cached: the same text, the same logical,
    or the same number. A percentage, an error number (Err:502) or an empty field says nothing."""
    kind, v = theirs
    if text == "" or text.endswith("%") or text.startswith("Err:"):
        return True
    if kind == "text":
        return text == v
    if kind == "bool":
        return text == v
    if kind == "number":
        try:
            return agrees(("number", float(text)), theirs)
        except ValueError:
            return False
    return True


def compare(data, case_file, abax_dump, trip_dump, csv_file, version, out):
    formulas = cases(case_file)
    mine = dump(abax_dump)
    trip = dump(trip_dump)
    shown = csv_view(csv_file)
    agreed, documented, bugs, trip_bad, torn = 0, 0, [], [], []
    sections = []

    for i, (section, formula) in enumerate(formulas):
        a = address(i)
        got = mine.get(a, ("", ""))
        t = trip.get(a, ("", ""))
        theirs = dump_kind(t[1])
        ours = dump_kind(got[1])

        if not sections or sections[-1][0] != section:
            sections.append((section, []))

        if formula in DIFFERENCES:
            want, why = DIFFERENCES[formula]
            documented += 1
            if not agrees(dump_kind(want), ours):
                bugs.append(f"{a} {formula}: abax {got[1]!r}, expected {want!r} ({why})")
            elif agrees(theirs, ours):
                bugs.append(f"{a} {formula}: listed as a difference, but LibreOffice agrees now")
            sections[-1][1].append((a, formula, sysl_value(*dump_kind(want)), f"LibreOffice gives {t[1]}; {why}"))
        elif agrees(theirs, ours):
            agreed += 1
            note = None
            if a in shown and not shows(shown[a], theirs):
                torn.append(f"{a} {formula}: LibreOffice caches {t[1]!r} in its .xlsx, its CSV export shows {shown[a]!r}")
                note = (f"LibreOffice caches this in the .xlsx it writes, but its CSV export of the same load "
                        f"shows {shown[a]}, counting the logical in A3 as the number 1 (Excel does not)")
            sections[-1][1].append((a, formula, sysl_value(*theirs), note))
        else:
            bugs.append(f"{a} {formula}: abax {got[1]!r}, LibreOffice {t[1]!r}")

        if a not in trip:
            trip_bad.append(f"{a} {formula}: missing from LibreOffice's .xlsx")
        elif not same_input(got[0], t[0]):
            trip_bad.append(f"{a} {formula}: abax wrote {got[0]!r}, LibreOffice's .xlsx reads back {t[0]!r}")

    # The data cells round-trip through LibreOffice's .xlsx too, inputs and values.
    for a, (inp, val) in mine.items():
        if not a.startswith("J"):
            t = trip.get(a)
            if t is None or not same_input(inp, t[0]) or not agrees(dump_kind(val), dump_kind(t[1])):
                trip_bad.append(f"{a}: abax wrote {inp!r} = {val!r}, LibreOffice's .xlsx reads back {t!r}")

    for b in bugs:
        print("DISAGREE", b)
    for b in trip_bad:
        print("ROUNDTRIP", b)
    for b in torn:
        print("LIBREOFFICE", b)
    print(f"oracle: {len(formulas)} formulas compared, {agreed} agree, {documented} documented differences, "
          f"{len(bugs)} disagreements, {len(trip_bad)} round-trip problems; "
          f"{len(torn)} cells where LibreOffice's CSV export differs from its own .xlsx")

    write_tests(data, sections, version, out)
    return 1 if bugs or trip_bad else 0


def write_tests(data, sections, version, out):
    rows = data_rows(data)
    lines = [
        "module sh.sysl.abax",
        "@tests",
        "",
        f"// What LibreOffice {version} works out for each formula in oracle/cases.txt, on the sheet",
        "// oracle/data.tsv sets up: written by oracle/oracle.zsh, which has abax write the workbook as",
        "// .xlsx, LibreOffice recalculate it, and every cell compared. Where the two differ and abax",
        "// follows Excel, the test says what LibreOffice gives and why. Regenerate it; do not edit it.",
        "",
        "private oracle_sheet() -> Sheet",
        "    var s = sheet(\"Sheet1\")",
        "",
    ]
    for r, fields in enumerate(rows):
        for c, f in enumerate(fields):
            if f:
                lines.append(f"    s.set({sysl_string(column(c) + str(r + 1))}, {sysl_string(f)}).unwrap()")
    for section, checks in sections:
        for a, formula, _, _ in checks:
            lines.append(f"    s.set({sysl_string(a)}, {sysl_string(formula)}).unwrap()")
    lines += ["    s", "end oracle_sheet", "",
              "// \"\" when the cell's value is `want` (a number to within one part in 10^12), else what differs.",
              "private differs(s: Sheet, addr: string, want: Value) -> string",
              "    val got = s.value(addr)",
              "    val same = want match",
              "        Number(x) ->",
              "            got match",
              "                Number(y) -> (x - y).abs() <= 1e-12 * x.abs().max(y.abs())",
              "                _ -> false",
              "        _ -> got == want",
              "",
              "    if same then \"\" else s\"${addr} ${s.input(addr)}: abax ${got.display()}, LibreOffice ${want.display()}\"",
              "end differs", ""]
    for n, (section, checks) in enumerate(sections):
        name = "".join(ch if ch.isalnum() else "_" for ch in section.lower()).strip("_")
        lines.append(f"@test(\"LibreOffice agrees with abax: {section.lower()}\")")
        lines.append(f"oracle_{name}()")
        lines.append("    val s = oracle_sheet()")
        lines.append("")
        for a, formula, value, note in checks:
            if note:
                lines.append(f"    // {formula}: {note}")
            lines.append(f"    assert_eq(differs(s, {sysl_string(a)}, {value}), \"\")")
        lines.append("")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines).rstrip("\n") + "\n")


def column(c):
    name = ""
    c += 1
    while c > 0:
        c, rem = divmod(c - 1, 26)
        name = chr(65 + rem) + name
    return name


def uncached(src, dst):
    """A copy of the .xlsx `src` with every formula's cached value taken out, so that LibreOffice has
    to work out every formula itself: given a cache, it may write a cell it never evaluated back with
    the value it was handed."""
    import re
    import zipfile
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("xl/worksheets/"):
                text = data.decode("utf-8")
                text = re.sub(r'<c r="([A-Z]+[0-9]+)"( t="[a-z]+")?><f>(.*?)</f>(<v>.*?</v>)?</c>',
                              lambda m: f'<c r="{m.group(1)}"><f>{m.group(3)}</f></c>', text)
                data = text.encode("utf-8")
            zout.writestr(info.filename, data)


if __name__ == "__main__":
    if sys.argv[1] == "compose":
        compose(*sys.argv[2:5])
    elif sys.argv[1] == "uncached":
        uncached(*sys.argv[2:4])
    elif sys.argv[1] == "compare":
        sys.exit(compare(*sys.argv[2:9]))

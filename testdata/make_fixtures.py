"""Writes the .xlsx fixtures that are not LibreOffice's, byte for byte as Excel lays them out:

    python3 -I testdata/make_fixtures.py testdata

- `excel_style.xlsx`: what Excel writes and LibreOffice does not -- shared formulas (a column and a
  block), `_xlfn.` prefixes, rich text with a phonetic guide, `_xHHHH_` escapes, inline strings, cells
  with no `r`, a formula with no cached value, an array formula, a sheet reached through an absolute
  relationship target, and a second sheet.
- the malformed packages the reader must refuse: `not_xlsx.xlsx` (a zip of something else),
  `no_sheet_part.xlsx` (the workbook names a part that is not there), `bad_xml.xlsx` (the sheet does
  not close a tag), `out_of_range.xlsx` (a cell past XFD).

The two made by LibreOffice (headless, with a private profile, as oracle/oracle.zsh runs it):

- `libreoffice.xlsx` from `libreoffice.fods` -- rich text, entities, typed formula results, a width,
  three sheets: `soffice --headless --convert-to "xlsx:Calc MS Excel 2007 XML" libreoffice.fods`.
- `excel_style.libreoffice.xlsx`, LibreOffice's reading of `excel_style.xlsx` written again: its
  shared formulas expanded by LibreOffice, so abax's expansion is checked against another reader.
"""

import os
import sys
import zipfile

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

CONTENT_TYPES = HEAD + (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    '</Types>')

ROOT_RELS = HEAD + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>')

WORKBOOK = HEAD + (
    f'<workbook xmlns="{MAIN}" xmlns:r="{R}"><sheets>'
    '<sheet name="Budget" sheetId="1" r:id="rId7"/>'
    '<sheet name="Notes" sheetId="2" r:id="rId8"/>'
    '</sheets></workbook>')

WORKBOOK_RELS = HEAD + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/>'
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com/" TargetMode="External"/>'
    '</Relationships>')

STRINGS = HEAD + (
    f'<sst xmlns="{MAIN}" count="4" uniqueCount="4">'
    '<si><t>Item</t></si>'
    '<si><r><t>rich </t></r><r><rPr><b/></rPr><t>text</t></r><rPh sb="0" eb="1"><t>PHONETIC</t></rPh></si>'
    '<si><t xml:space="preserve">line_x000D_break and _x005F_x0041_ kept</t></si>'
    '<si><t>12</t></si>'
    '</sst>')

# The first sheet, Budget: A holds numbers, B a shared formula down a column, C:D a shared formula
# over a block, E text of every kind, F the rest.
BUDGET = HEAD + (
    f'<worksheet xmlns="{MAIN}" xmlns:r="{R}">'
    '<dimension ref="A1:F4"/>'
    '<cols><col min="1" max="1" width="9.140625" customWidth="1"/><col min="2" max="3" width="20.7109375" customWidth="1"/>'
    '<col min="4" max="4" width="30" customWidth="0"/></cols>'
    '<sheetData>'
    '<row r="1">'
    '<c r="A1"><v>1.5E-3</v></c>'
    '<c r="B1"><f t="shared" ref="B1:B4" si="0">A1*2</f><v>3E-3</v></c>'
    '<c r="C1"><f t="shared" ref="C1:D2" si="1">$A$1+A1</f><v>3E-3</v></c>'
    '<c r="D1"><f t="shared" si="1"/><v>6E-3</v></c>'
    '<c r="E1" t="s"><v>0</v></c>'
    '<c r="F1" t="str"><f>_xlfn.CONCAT(E1,"!")</f><v>Item!</v></c>'
    '</row>'
    '<row r="2">'
    '<c r="A2"><v>2</v></c>'
    '<c r="B2"><f t="shared" si="0"/><v>4</v></c>'
    '<c r="C2"><f t="shared" si="1"/><v>2.0015</v></c>'
    '<c r="D2"><f t="shared" si="1"/><v>4.0045</v></c>'
    '<c r="E2" t="s"><v>1</v></c>'
    '<c r="F2" t="b"><v>1</v></c>'
    '</row>'
    '<row r="3">'
    '<c r="A3"><v>3</v></c>'
    '<c r="B3"><f t="shared" si="0"/><v>6</v></c>'
    '<c><v>7</v></c>'
    '<c r="E3" t="s"><v>2</v></c>'
    '<c r="F3" t="e"><v>#N/A</v></c>'
    '</row>'
    '<row r="4">'
    '<c r="A4" t="inlineStr"><is><r><t>in</t></r><r><t>line</t></r></is></c>'
    '<c r="B4" t="e"><f t="shared" si="0"/><v>#VALUE!</v></c>'
    '<c r="C4"><f>SUM(A1:A3)</f></c>'
    '<c r="D4"><f t="array" ref="D4:D5">A1:A2*2</f><v>0.003</v></c>'
    '<c r="E4" t="s"><v>3</v></c>'
    '<c r="F4" s="3"/>'
    '</row>'
    '</sheetData></worksheet>')

NOTES = HEAD + (
    f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>99</v></c></row></sheetData></worksheet>')


def package(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in parts:
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)


def good():
    return [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("xl/workbook.xml", WORKBOOK),
        ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
        ("xl/sharedStrings.xml", STRINGS),
        ("xl/worksheets/sheet1.xml", NOTES),
        ("xl/worksheets/sheet2.xml", BUDGET),
    ]


def main(out):
    package(os.path.join(out, "excel_style.xlsx"), good())
    package(os.path.join(out, "not_xlsx.xlsx"), [("readme.txt", "just a zip")])
    package(os.path.join(out, "no_sheet_part.xlsx"),
            [p for p in good() if p[0] != "xl/worksheets/sheet2.xml"])
    package(os.path.join(out, "bad_xml.xlsx"),
            [(n, t if n != "xl/worksheets/sheet2.xml" else t.replace("</row>", "", 1)) for n, t in good()])
    package(os.path.join(out, "out_of_range.xlsx"),
            [(n, t if n != "xl/worksheets/sheet2.xml" else t.replace('r="F4"', 'r="XFE4"')) for n, t in good()])


if __name__ == "__main__":
    main(sys.argv[1])

# abax — design

Named for the *abax* (ἄβαξ), the Greek counting board — a table of cells you reckon on, and the word
"abacus" comes from.

## What it is for

1. **A spreadsheet a person can actually use**: a grid, a formula bar, a menu, the keys everyone
   already knows, and files that open in Excel and LibreOffice.
2. **A stressor for sysl and syslUI**: a dynamically-typed value model inside a statically-typed
   language, a recursive formula tree, a dependency graph recalculated in topological order with
   cycles reported rather than looped on, a sparse store of a million addressable rows, shortest
   round-trip number formatting, and a grid view that draws only what is on screen. **syslui,
   syslui-sdl and sdl3 are enhanced wherever abax needs them** — a gap found there is fixed there,
   released, and consumed.

## Shape

- `sh.sysl.abax` — the library, at the root. Values, cell addresses and ranges, the formula parser
  (on `sh.sysl.parsing`'s Pratt loop), the evaluator and its function table, the `Sheet` and
  `Workbook` with recalculation, editing operations (fill, copy, insert/delete rows and columns with
  references adjusted), undo, and the file formats. **No OS, no SDL** — it reads and writes bytes.
- `app/` — `abax-app`, the application, a workspace member on syslui-sdl, with a `--shot` mode that
  draws one frame to a PNG with no window, as amber's viewer does.
- `cli/` — `abax-cli`, which loads a file, recalculates, and prints cells or converts between formats;
  it is how the formula engine is driven by a script and by the oracle.

## The model

- **`Value`**: `Empty`, `Number(real)`, `Text(string)`, `Bool(bool)`, `Error(Fault)`, where `Fault` is
  the spreadsheet errors every program agrees on: `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#N/A`,
  `#NUM!`, `#NULL!`, and a circular reference.
- **A cell holds what was typed and what it means**: the input string, and either a constant or a
  parsed formula. What a formula last evaluated to is cached beside it.
- **Addresses** are A1 (`B7`, `$B$7`, `AA100`), a range is two of them (`A1:C10`), whole columns and
  rows (`A:A`, `3:3`) included. Limits are Excel's: 1,048,576 rows by 16,384 columns (`XFD`).
  A reference may name another sheet (`Sheet2!A1`, `'My sheet'!A1:B2`).
- **The store is sparse**: only cells holding something exist.
- **Recalculation**: every formula's precedents are recorded when it is parsed; changing a cell
  recalculates exactly what depends on it, in topological order. A cycle makes every cell on it a
  circular-reference error rather than a hang.
- **Coercion follows Excel**: text that reads as a number is one in arithmetic, `TRUE` is 1, an empty
  cell is 0 or "" by context, and a function given a range skips text where Excel does (`SUM`) and
  counts it where Excel does (`COUNTA`).
- **Editing follows Excel's reference rules**: a copied formula moves its relative parts and keeps
  its `$` parts; inserting or deleting rows and columns moves every reference past them, `$` or not,
  grows or shrinks a range, and turns a reference to a deleted cell into `#REF!`; a cut and paste
  moves the cells and every reference to them.
- **Undo records inputs, not inverse operations**: a step is, for each cell and column width it
  touched, what was there before and after. Undo puts the befores back through the same path as any
  edit. An inverse cannot bring back a deleted row, or the formulas that became `#REF!` when it went;
  a record of inputs can, at the cost of what the edit touched.
- **The system clipboard carries displayed values**, tab-separated, as Excel gives other programs;
  abax's own copy keeps the inputs, and a paste uses them while the clipboard still holds the text
  that copy put there.
- **Numbers are shown in General format** — up to eleven significant figures, switching to scientific
  where Excel does — and written to files as the shortest text that reads back to the same `real`.

## Oracles

- **LibreOffice, headless**: a workbook written by abax is recalculated by LibreOffice
  (`soffice --headless --convert-to`) and every cell's value compared with abax's. This is the oracle
  for the formula engine and for the file format at once.
- Excel's documented semantics where LibreOffice differs, said in the test.

## Roadmap

Each is a piece of work of its own, in this order.

1. **The formula engine**: values, addresses, ranges, the parser (literals, references, ranges,
   operators with Excel's precedence including `&`, `%`, `^` and comparison, function calls), the
   evaluator, the core functions (arithmetic and statistics, logic, text, rounding, lookups:
   `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`), the sheet with dependency-ordered recalculation and cycle
   detection, General-format display.
2. **The application**: syslui-sdl window with a menu bar (File, Edit, Insert, Format, Help), a name
   box and a formula bar, the grid drawn on demand with frozen column and row headers, selection and
   range selection with the mouse and the keyboard, typing into a cell, Enter/Tab/arrows, column
   widths dragged, scrolling. `--shot`.
3. **Editing operations**: copy, cut and paste with relative references adjusted, fill down and
   right, insert and delete rows and columns with every reference adjusted (and `#REF!` where one is
   deleted), undo and redo, the system clipboard as tab-separated text.
4. **Files**: CSV in and out; **`.xlsx`** read and written through `sysl-lang/zip` and a
   `sysl-lang/xml` package (written for this), with the LibreOffice oracle. File dialogs (SDL3's).
5. **Formatting**: number formats (fixed, percent, currency, dates), bold/italic, alignment, fill
   colour — kept in the model and in `.xlsx`.
6. **Several sheets**: tabs, cross-sheet references, rename, insert, delete.
7. **Scripting in slate**: functions a cell can call (`=TAX(B2)`), macros on the menu that read and
   change the workbook, and `abax-cli run <script>` for batch work. slate is taken **as a sysl
   library**, not through its C API — slate is enhanced with a sysl embedding surface (load a script,
   call by name, values across the boundary without a C hop, host functions, a step budget,
   cancellation, and a sandbox with no file or network access, since a script arrives inside a file
   someone else sent). Measured: the cost of a user function called across a hundred thousand cells.
8. **More functions**: dates and times, `SUMIF`/`COUNTIF`/`AVERAGEIF` and their `S` forms, text
   functions, `XLOOKUP`, financial basics.
9. **Large sheets**: a million rows loaded and scrolled at frame rate; recalculation measured.

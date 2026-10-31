# abax

A spreadsheet in [sysl](https://github.com/sysl-lang/sysl): a formula engine that recalculates in
dependency order, and an application on [syslui-sdl](https://github.com/sysl-lang/syslui-sdl) with a
menu, a formula bar and the grid. Named for the *abax*, the Greek counting board the word "abacus"
comes from.

[DESIGN.md](DESIGN.md) has the shape and the roadmap.

## The formula engine

`sh.sysl.abax` is the library: values with Excel's coercion rules, A1 addresses up to `XFD1048576`
(with `$`, whole columns and rows, and sheet names), a formula parser on
[`sh.sysl.parsing`](https://github.com/sysl-lang/parsing)'s Pratt loop with Excel's precedence (`-2^2`
is 4, `2^3^2` is 64), a printer that writes a formula back in canonical form, about forty functions
(`SUM`, `AVERAGE`, `IF`, `VLOOKUP`, `INDEX`, `MATCH`, the text functions and the rest), and a sparse
`Sheet` that recalculates exactly what depends on a change, in dependency order, with a cycle shown as
a circular-reference error rather than looped on. Numbers are shown in General format. It needs
nothing but an allocator.

```sysl
import sh.sysl.abax.sheet

var s = sheet("Sheet1")

s.set("A1", "4").unwrap()
s.set("A2", "=A1^2/2").unwrap()
s.set("A3", "=IF(A2>5,\"big\",\"small\")&\" \"&SUM(A1:A2)").unwrap()

print(s.display("A3"))      // big 12
s.set("A1", "1").unwrap()
print(s.display("A3"))      // small 1.5
print(s.formula("A2").unwrap())     // =A1^2/2
```

```
dependencies {
  abax { git = "github.com/sysl-lang/abax", version = "0.1.0" }
}
```

## The application

![abax with the budget example open, a range selected and its sum on the bottom line](shot.png)

`app/` is `abax-app`, a window on [syslui-sdl](https://github.com/sysl-lang/syslui-sdl): a menu bar, a
name box and a formula bar, the grid, and a line that shows the average, count and sum of the
selection, or why a formula was refused. The grid draws only the cells on screen, so all of
`A1:XFD1048576` scrolls at the same speed.

```
sysl build app
app/abax-app                    an empty sheet
app/abax-app --demo             a small budget to look at
app/abax-app --demo --select B4:D8 --shot shot.png
```

`--shot` draws one frame with no window and writes it to a PNG; `--cell B2==SUM(A1:A9)` (repeatable)
and `--select` set the sheet up first, so the picture can be checked by a program.

| Key | Does |
|---|---|
| arrows, Tab, Shift+Tab, Enter, Shift+Enter | move one cell |
| Shift + an arrow, Shift + a click, a drag | stretch the selection |
| Ctrl + an arrow | to the edge of the block of filled cells |
| Home, Ctrl+Home, Ctrl+End | the row's first cell, A1, the last cell used |
| Page Up, Page Down | a screenful |
| typing | starts the cell over; Enter or Tab puts it in, Escape throws it away |
| F2, a double click | change what is in the cell; F2 again makes the arrows move out of it |
| Delete, Backspace | empty the selection |
| a click on a column letter or row number, or the corner | the whole column, row or sheet |
| a drag on a column letter's right edge | the column's width |
| Ctrl+N, Ctrl+Q, Ctrl+A | new sheet, quit, select all (⌘ on a Mac) |

Typing an address or a range (`B7`, `B2:D9`, `C:E`, `3:5`) into the name box and pressing Enter
selects it. A formula typed into the formula bar goes into the active cell on Enter. Saving and
opening files, copy and paste, and undo are not written yet; their menu items are there, disabled or
saying so.

## Licence

ISC — see [LICENSE](LICENSE).

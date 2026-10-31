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

## Licence

ISC — see [LICENSE](LICENSE).

#!/bin/zsh
# The LibreOffice oracle for abax's formula engine and its .xlsx files.
#
#     zsh oracle/oracle.zsh
#
# 1. composes a sheet from oracle/data.tsv and the formulas in oracle/cases.txt (one to a row, in
#    column J), and has abax-cli read it and write it as .xlsx, with abax's values cached;
# 2. has LibreOffice -- headless, in a private profile set to recalculate every .xlsx it loads, so it
#    never touches the user's own and can run beside an open LibreOffice -- recalculate the workbook
#    and write it again as .xlsx, with its own values cached;
# 3. reads LibreOffice's .xlsx back with abax-cli: its cached values are LibreOffice's answers, and
#    its formulas show that every one survived the trip;
# 4. compares every cell, prints each disagreement, and writes sh/sysl/abax/tests_oracle.sysl from
#    the agreed values, so `sysl test .` checks them without LibreOffice.
#
# It exits non-zero when anything disagrees.

set -e

here=${0:A:h}
root=${here:h}
soffice=/Applications/LibreOffice.app/Contents/MacOS/soffice
work=$(mktemp -d -t abax-oracle)
cli=$root/cli/abax-cli

nice sysl build $root/cli

python3 -I $here/oracle.py compose $here/data.tsv $here/cases.txt $work/book.tsv
$cli convert --formulas $work/book.tsv $work/book.xlsx
$cli dump $work/book.xlsx > $work/abax.txt

mkdir -p $work/profile/user
print -r -- '<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>' > $work/profile/user/registrymodifications.xcu

profile="-env:UserInstallation=file://$work/profile"

# LibreOffice gets the workbook with abax's cached values taken out, so that every value it writes is
# one it worked out: handed a cache, it writes some cells back with the value they came with.
mkdir -p $work/bare
python3 -I $here/oracle.py uncached $work/book.xlsx $work/bare/book.xlsx

$soffice $profile --headless --convert-to "xlsx:Calc MS Excel 2007 XML" --outdir $work/trip $work/bare/book.xlsx > $work/soffice.log 2>&1
$cli dump $work/trip/book.xlsx > $work/trip.txt

# The same workbook as CSV: separator, quote, UTF-8, from line 1, no column formats, language, quote
# text, detect special numbers, values not as shown, values not formulas, no trimming, first sheet.
$soffice $profile --headless --convert-to "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,true,true,false,false,false,1" --outdir $work/values $work/bare/book.xlsx >> $work/soffice.log 2>&1
values=($work/values/book*.csv)

version=$(defaults read /Applications/LibreOffice.app/Contents/Info.plist CFBundleShortVersionString)

python3 -I $here/oracle.py compare $here/data.tsv $here/cases.txt $work/abax.txt $work/trip.txt $values[1] "$version" $root/sh/sysl/abax/tests_oracle.sysl
print "oracle: work files in $work"

# DuckSheets

A free spreadsheet for Windows that opens and saves Excel workbooks (.xlsx, .xlsm) and CSV files.

- **Excel files that stay Excel files.** Formulas, formats, column widths, merged cells, frozen panes, filters and charts are read and written as Excel keeps them. Anything DuckSheets doesn't edit yet stays in the file unchanged: conditional formatting, comments, pictures, other kinds of chart, print settings.
- **Formulas.** About 260 of Excel's functions, checked against the real Excel:
  - SUM, AVERAGE, COUNT, MIN/MAX, ROUND, SUBTOTAL;
  - IF, IFS, IFERROR, AND/OR, LET;
  - the conditional family: SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS/MINIFS;
  - lookups: VLOOKUP/HLOOKUP, XLOOKUP, INDEX/MATCH, LOOKUP, OFFSET, INDIRECT;
  - dynamic arrays: FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, TAKE, DROP, VSTACK, HSTACK, CHOOSECOLS;
  - text: LEFT/RIGHT/MID, CONCAT, TEXTJOIN, TEXT, SUBSTITUTE, FIND/SEARCH, TEXTBEFORE/TEXTAFTER, TEXTSPLIT;
  - dates: TODAY, DATE, EDATE/EOMONTH, DATEDIF, NETWORKDAYS, WEEKDAY, ISOWEEKNUM, YEARFRAC;
  - maths and trigonometry: SIN/COS/TAN and their inverses, GCD, LCM, MROUND, COMBIN;
  - statistics: MEDIAN, STDEV, PERCENTILE, RANK, CORREL, SLOPE/INTERCEPT, FORECAST, NORM.DIST/NORM.INV, BINOM.DIST;
  - money: PMT, FV, PV, NPV, IPMT/PPMT, RATE, NPER, IRR, XNPV/XIRR;
  - engineering: DEC2BIN/HEX2DEC and the other number bases, BITAND/BITXOR, and complex numbers (IMSUM, IMPRODUCT, IMABS…);
  - and more.
- **References.** Relative, `$` absolute, other sheets and named ranges. Arguments can be separated with `,` or `;`, as in Norwegian Excel.
- **Typing like Excel:**
  - `10%` becomes a percentage and `20.10.2026` a date, in your own number and date format;
  - Enter moves down, Tab right, F2 edits;
  - while typing a formula, clicking cells or using the arrow keys puts their addresses in;
  - F4 switches `$`.
- **The fill handle.** Drag to copy formulas and continue series: 1, 2, 3…, Week 1, Week 2…, months, weekdays, dates.
- **Rows and columns.** Insert and delete them, hide them, and resize them (double-click an edge to fit). Formulas adjust everywhere.
- **Formatting.**
  - Text: font, size (also Alt+scroll wheel), bold/italic/underline/strike and colours.
  - Cells: fill, borders, alignment, wrap text, merge and centre.
  - Numbers: number formats, decimals.
- **Dynamic arrays.** FILTER, SORT, SORTBY, UNIQUE, SEQUENCE and TRANSPOSE spill their results into the cells around them, as in Excel 365, and are saved as Excel's own dynamic arrays.
- **Conditional formatting.** Highlight cells greater than, between or equal to a value, text that contains something, duplicates, the top 10 or those above average, or where a formula is true; data bars and colour scales.
- **Data validation.** A drop-down list to choose from (Alt+Down), or limits on numbers; values that don't fit are refused.
- **Data tools.** AutoSum, sort, filter (with search), remove duplicates, text to columns, goal seek, freeze panes, find and replace.
- **Typing help.** AutoComplete finishes words from the column, Format cells (Ctrl+1), format painter, Paste special (values, formulas, formats, transpose) and Show formulas (Ctrl+`).
- **Links.** Ctrl+K links a cell to a web or e-mail address or a place in the workbook; a click follows it, and HYPERLINK works too.
- **Notes.** Notes on cells, as in Excel (Shift+F2, or right-click): a red corner shows where they are, and the mouse over the cell shows the note.
- **Printing.** Ctrl+P prints the sheet (or the selection), with its formats; choose "Microsoft Print to PDF" to save a PDF.
- **Charts.** Column, bar, line, area, pie and doughnut charts, saved as real Excel charts linked to the cells.
- **Copy and paste** with Excel and other programs.

## Keys

| | |
|---|---|
| Enter / Tab | Move down / right (within a selected block) |
| F2 | Edit the cell |
| Ctrl+Enter | Put what's typed into every selected cell |
| Alt+Enter | New line in the cell |
| Ctrl+arrow | To the edge of the data |
| Ctrl+Home / Ctrl+End | First / last cell |
| Ctrl+A | The data around the cell, then everything |
| Shift+Space / Ctrl+Space | The whole row / column |
| Ctrl++ / Ctrl+- | Insert / delete rows or columns |
| Ctrl+D / Ctrl+R | Fill down / right |
| Ctrl+B / I / U / 5 | Bold / italic / underline / strikethrough |
| Ctrl+Shift+1 / 3 / 4 / 5 | Number / date / currency / percentage |
| Ctrl+; / Ctrl+Shift+; | Today's date / the time |
| Alt+= | AutoSum |
| Ctrl+Shift+L | Filter |
| Ctrl+F / Ctrl+H | Find / replace |
| Ctrl+1 | Format cells |
| Ctrl+Alt+V | Paste special |
| Ctrl+` | Show formulas |
| Alt+Down | Open a cell's drop-down list |
| Ctrl+P | Print |
| Ctrl+K | Link |
| Shift+F2 | Note |
| Ctrl+PageUp / PageDown | Previous / next sheet |
| Shift+F11 | New sheet |
| Ctrl+scroll wheel | Zoom |
| Alt+scroll wheel | Bigger / smaller text |

## Building

```
npm install
npm test
npm run tauri build -- --no-bundle
npm run smoke
```

With Excel installed, `EXCEL_CHECK=1 npm test` and `EXCEL_CHECK=1 npm run smoke` also check that Excel opens DuckSheets' files as intended, and that DuckSheets reads Excel's.

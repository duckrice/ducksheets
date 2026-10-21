# DuckSheets

A free spreadsheet for Windows that opens and saves Excel workbooks (.xlsx, .xlsm) and CSV files.

- **Excel files that stay Excel files.** Formulas, formats, column widths, merged cells, frozen panes, filters and charts are read and written as Excel keeps them. Anything DuckSheets doesn't edit yet stays in the file unchanged: conditional formatting, comments, pictures, other kinds of chart, print settings.
- **Formulas.** About 110 of Excel's functions:
  - SUM, AVERAGE, COUNT, MIN/MAX, ROUND;
  - IF, IFS, IFERROR, AND/OR;
  - the conditional family: SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS/MINIFS;
  - lookups: VLOOKUP/HLOOKUP, XLOOKUP, INDEX/MATCH, OFFSET, INDIRECT;
  - text: LEFT/RIGHT/MID, CONCAT, TEXTJOIN, TEXT, SUBSTITUTE, FIND/SEARCH;
  - dates: TODAY, DATE, EDATE/EOMONTH, DATEDIF, NETWORKDAYS, WEEKDAY, ISOWEEKNUM;
  - statistics: MEDIAN, STDEV, PERCENTILE, RANK;
  - money: PMT, FV, PV, NPV;
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
- **Data tools.** AutoSum, sort, filter (with search), freeze panes, find and replace.
- **Charts.** Column, bar, line and pie charts, saved as real Excel charts linked to the cells.
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

# Scripts: Office Scripts

    sheets script book.xlsx report.ts               # also .js, -e 'code', or stdin
    sheets script book.xlsx report.ts --arg 2026     # main(workbook, 2026)

    function main(workbook: ExcelScript.Workbook) {
      const ws = workbook.getWorksheet("Sales")!;
      const rows = ws.getUsedRange()!.getValues();          // computed values
      const late = rows.slice(1).filter(r => (r[4] as number) > 30);
      ws.getRange("H1").setValue("Late");
      ws.getRange("H2").setFormula("=COUNTIF(E:E,\">30\")");
      return { late: late.length, check: ws.getRange("H2").getValue() };
    }

- For logic that commands do not express: loops, conditions, reshaping,
  reconciling. TypeScript or JavaScript, in Office Scripts' API
  (ExcelScript), as in Excel's Automate tab. Without `main`, plain code
  runs with the global `workbook`. `console.log` prints; what `main`
  returns is printed as JSON.
- Reads give what the cells hold now: values, and formula results the
  sheets calculator computes, the script's own changes included (null,
  with a warning, for a result it cannot compute). Dates read as serial
  numbers, as in Excel; `getTexts` gives what Excel shows. Writes take
  values as typed input: `"=SUM(A1:A9)"` is a formula, `"2026-07-01"` a
  date; a Date is written as its UTC date and time; null leaves a cell as
  it is.
- Saved once at the end, as one undo step, like `do`; a script that
  throws saves nothing and the error names its line. A read after writes
  costs a save in memory, so read what you need first, then write.
- What it has: `workbook.getWorksheet(s)`, `getActiveWorksheet`,
  `addWorksheet`, `getTable(s)`, `getNamedItem`; on a sheet `getRange`
  ("A1:C9", "B:B"), `getRangeByIndexes`, `getCell`, `getUsedRange`,
  `getFreezePanes`, `getAutoFilter`, `setName`, `copy`, `delete`; on a
  range values, formulas, texts, number formats and value types (get and
  set), `getCell`, `getRow`, `getColumn`, `getOffsetRange`,
  `getResizedRange`, `getEntireRow`, `getSurroundingRegion`, `clear`,
  `delete` and `insert` of whole rows or columns, `copyFrom`, `moveTo`,
  `getSort().apply`, `removeDuplicates`, `find`, `setHyperlink`, and
  `getFormat()` with font, fill, alignment, wrap, borders, widths. Tables
  are Excel tables, or a sheet's tables named as `query` names them
  (`Sales`, `Sales_2`): `getRange`, `getHeaderRowRange`,
  `getRangeBetweenHeaderAndTotal`, `getColumns`, `getColumnByName`,
  `addRows` (at the end, as `write --append`: above the total row,
  formula columns filled), `deleteRowsAt`, `getSort`.
- Not there: creating Excel tables, charts, conditional formats and
  validation (use `chart add`, `cond`, `validate` after the script), and
  filtering rows by criteria. Positions count from 0, as in Office Scripts.
- No files or network; `--timeout` (default 1m) stops a runaway loop.

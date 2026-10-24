Build report.xlsx from sales.csv and targets.csv:

- a Sales sheet with the sales data and a Revenue column computed as
  Units × Price (a formula in each row);
- a Summary sheet with one row per region and the columns Region, Revenue
  (summing the Sales sheet with formulas, so it updates when Sales
  changes), Target (from targets.csv) and Gap (Revenue minus Target, a
  formula);
- on Summary, Gap shown in red when it is negative, and Revenue, Target
  and Gap as dollar amounts;
- on both sheets, a bold header row that stays in view when scrolling.

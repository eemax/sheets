For spreadsheet work, use Python with openpyxl and pandas. Run scripts with
`uv run --with openpyxl --with pandas python script.py` (the packages are
fetched on first use). LibreOffice is not installed. Do not use the `sheets`
tool.
Excel functions newer than Excel 2007 (such as XLOOKUP, IFS, MAXIFS, TEXTJOIN, CONCAT, UNIQUE) must be written with their `_xlfn.` prefix in openpyxl (`=_xlfn.XLOOKUP(A2,B:B,C:C)`), or Excel shows #NAME? (FILTER and SORT need `_xlfn._xlws.`; LET and LAMBDA also need `_xlpm.` before their variable names).

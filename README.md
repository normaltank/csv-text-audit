# CSV Text Audit

CSV Text Audit checks whether selected fields from an original CSV survived as text in a saved Excel workbook. Use it when a spreadsheet is part of a handoff for identifiers, product codes, postal codes, or other fields whose exact spelling matters.

A numeric cell may display `00123` because its number format adds the zeros. The audit still flags it because the workbook stores `123`. It also detects changed long integers, date or time values, scientific notation, and formulas stored in a text field. CSV fields are read as strings. Workbook formulas are inspected as stored and never evaluated.

```sh
python -m pip install .
csv-text-audit original.csv edited.xlsx --sheet Data --key row_id --column postal_code --column account_id
```

- Supply `--column` once for each field that must remain text. For example, `42` stored as the number 42 is reported as `number`.
- Row 1 must contain unique column names. Columns may be reordered. Use `--key` with an unchanged, unique text column if rows can be reordered, added, or deleted. Without a key, both inputs need the same number of data rows in the same order.
- Use `--sheet` for a workbook with multiple worksheets. The CSV delimiter is a comma. Use `--encoding` if the CSV uses an encoding other than UTF-8.
- Add `--json` for a machine-readable array. Findings identify row and column locations without printing cell contents. CSV record numbers count the header as the first record and exclude blank lines. XLSX row numbers match Excel.
- Exit code `0` indicates that the selected fields were preserved, `1` indicates differences, and `2` indicates an input or usage error.

The audit reads `.xlsx` files without opening Excel, changing either input, or sending data to a server. A finding describes the stored difference. The comparison cannot determine whether an import setting or a later edit caused the difference. Excel documents the import options used to preserve identifiers as text. [Microsoft Support](https://support.microsoft.com/en-us/excel/keeping-leading-zeros-and-large-numbers)

## Development

```sh
python -m pip install .
python -m unittest discover -s tests -v
```

MIT licensed.

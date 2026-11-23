import argparse
import csv
import json
import re
import sys
from dataclasses import asdict, dataclass
from decimal import Decimal
from pathlib import Path
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


@dataclass(frozen=True)
class Cell:
    value: object = None
    kind: str = "n"


@dataclass(frozen=True)
class Finding:
    code: str
    csv_row: int | None = None
    xlsx_row: int | None = None
    column: str | None = None


EMPTY = Cell()


def read_csv(path, encoding):
    with path.open(encoding=encoding, newline="") as handle:
        records = csv.reader(handle, strict=True)
        header = next(records, [])
        rows = []
        for row in records:
            if not row:
                continue
            number = len(rows) + 2
            if len(row) != len(header):
                raise ValueError(
                    f"CSV record {number} has {len(row)} fields, expected {len(header)}"
                )
            rows.append(row)
    return header, rows


def read_xlsx(path, sheet):
    workbook = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        if sheet is None:
            if len(workbook.worksheets) != 1:
                raise ValueError(
                    "specify --sheet for a workbook with multiple worksheets"
                )
            worksheet = workbook.worksheets[0]
        else:
            worksheet = next(
                (item for item in workbook.worksheets if item.title == sheet), None
            )
            if worksheet is None:
                raise ValueError(f"XLSX worksheet {sheet!r} does not exist")

        worksheet.reset_dimensions()
        records = worksheet.iter_rows()
        header = [cell.value for cell in next(records, ())]
        while header and header[-1] is None:
            header.pop()

        rows = []
        for number, row in enumerate(records, 2):
            if any(cell.value is not None for cell in row[len(header) :]):
                raise ValueError(f"XLSX row {number} contains data beyond the header")
            cells = tuple(
                Cell(cell.value, cell.data_type) for cell in row[: len(header)]
            )
            rows.append(cells + (EMPTY,) * (len(header) - len(cells)))
        while rows and all(cell.value is None for cell in rows[-1]):
            rows.pop()
        return header, rows
    finally:
        workbook.close()


def header_indexes(header, source):
    if not header or any(
        not isinstance(name, str) or not name.strip() for name in header
    ):
        raise ValueError(
            f"{source} requires non-empty text column names in the first row"
        )
    indexes = {name: index for index, name in enumerate(header)}
    if len(indexes) != len(header):
        raise ValueError(f"{source} column names must be unique")
    return indexes


def key_indexes(keys, source):
    indexes = {}
    for number, key in enumerate(keys, 2):
        if not isinstance(key, str) or not key.strip():
            raise ValueError(f"{source} row {number} needs a non-empty text key")
        if key in indexes:
            raise ValueError(
                f"{source} row {number} duplicates a key from row {indexes[key]}"
            )
        indexes[key] = number
    return indexes


def change_code(original, stored):
    if stored.kind == "f":
        return "formula"
    if stored.kind == "e":
        return "spreadsheet-error"
    if stored.value is None:
        return "value-missing" if original else None
    if isinstance(stored.value, str):
        return "text-changed" if original != stored.value else None
    if stored.kind == "d":
        return "date-or-time"
    if stored.kind == "b":
        return "boolean"
    if stored.kind == "n":
        number = Decimal(str(stored.value))
        if re.fullmatch(r"[+-]?0[0-9]+", original) and Decimal(original) == number:
            return "leading-zeros"
        if (
            re.fullmatch(r"[+-]?[1-9][0-9]{15,}", original)
            and Decimal(original) != number
        ):
            return "long-integer-changed"
        if re.fullmatch(r"[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)[eE][+-]?[0-9]+", original):
            try:
                if Decimal(original) == number:
                    return "scientific-notation"
            except ArithmeticError:
                pass
        return "number"
    return "cell-type-changed"


def audit(csv_path, xlsx_path, columns, sheet=None, key=None, encoding="utf-8-sig"):
    if xlsx_path.suffix.lower() != ".xlsx":
        raise ValueError("the workbook must use the .xlsx format")
    csv_header, csv_rows = read_csv(csv_path, encoding)
    xlsx_header, xlsx_rows = read_xlsx(xlsx_path, sheet)
    csv_indexes = header_indexes(csv_header, "CSV")
    xlsx_indexes = header_indexes(xlsx_header, "XLSX")
    requested = list(dict.fromkeys([*columns, key] if key else columns))
    for name in requested:
        if name not in csv_indexes or name not in xlsx_indexes:
            raise ValueError(f"column {name!r} must exist in both inputs")

    findings = []
    if key:
        csv_keys = key_indexes((row[csv_indexes[key]] for row in csv_rows), "CSV")
        key_cells = (row[xlsx_indexes[key]] for row in xlsx_rows)
        xlsx_keys = key_indexes(
            (
                cell.value if cell.kind in {"s", "inlineStr"} else None
                for cell in key_cells
            ),
            "XLSX",
        )
        pairs = [(number, xlsx_keys.get(value)) for value, number in csv_keys.items()]
        findings.extend(
            Finding("xlsx-row-added", xlsx_row=number)
            for value, number in xlsx_keys.items()
            if value not in csv_keys
        )
    else:
        if len(csv_rows) != len(xlsx_rows):
            raise ValueError("row counts differ, so use --key to match rows")
        pairs = [(number, number) for number in range(2, len(csv_rows) + 2)]

    for csv_number, xlsx_number in pairs:
        if xlsx_number is None:
            findings.append(Finding("csv-row-missing", csv_row=csv_number))
            continue
        for name in dict.fromkeys(columns):
            original = csv_rows[csv_number - 2][csv_indexes[name]]
            stored = xlsx_rows[xlsx_number - 2][xlsx_indexes[name]]
            code = change_code(original, stored)
            if code is not None:
                findings.append(Finding(code, csv_number, xlsx_number, name))
    return findings


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Check that selected CSV fields survived as text in XLSX"
    )
    parser.add_argument("csv", type=Path, help="original comma-separated CSV")
    parser.add_argument("xlsx", type=Path, help="resulting Excel workbook")
    parser.add_argument(
        "--column",
        action="append",
        required=True,
        help="text column to audit, repeatable",
    )
    parser.add_argument("--sheet", help="XLSX worksheet name")
    parser.add_argument("--key", help="unchanged, unique text column to match rows")
    parser.add_argument(
        "--encoding", default="utf-8-sig", help="CSV encoding, default utf-8-sig"
    )
    parser.add_argument("--json", action="store_true", help="print findings as JSON")
    args = parser.parse_args(argv)
    try:
        findings = audit(
            args.csv, args.xlsx, args.column, args.sheet, args.key, args.encoding
        )
    except (
        OSError,
        ValueError,
        LookupError,
        csv.Error,
        BadZipFile,
        InvalidFileException,
        ParseError,
    ) as error:
        print(f"csv-text-audit: {error}", file=sys.stderr)
        return 2

    if args.json:
        print(json.dumps([asdict(finding) for finding in findings]))
    elif findings:
        for finding in findings:
            csv_row = (
                f"CSV record {finding.csv_row}" if finding.csv_row is not None else None
            )
            xlsx_row = (
                f"XLSX row {finding.xlsx_row}" if finding.xlsx_row is not None else None
            )
            location = ", ".join(
                part for part in (csv_row, xlsx_row, finding.column) if part is not None
            )
            print(f"{location}: {finding.code}")
    else:
        print("Selected fields were preserved as text.")
    return int(bool(findings))


if __name__ == "__main__":
    raise SystemExit(main())

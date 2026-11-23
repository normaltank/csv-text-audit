import csv
import json
import subprocess
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook


class CliTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.directory = Path(temp.name)
        self.csv_path = self.directory / "source.csv"
        self.xlsx_path = self.directory / "edited.xlsx"

    def write_csv(self, *rows):
        with self.csv_path.open("w", encoding="utf-8-sig", newline="") as output:
            csv.writer(output).writerows(rows)

    def write_workbook(self, *rows):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Data"
        for row in rows:
            sheet.append(row)
        self.addCleanup(workbook.close)
        return workbook, sheet

    def run_cli(self, *options):
        return subprocess.run(
            [
                sys.executable,
                "-m",
                "csv_text_audit",
                self.csv_path,
                self.xlsx_path,
                *options,
            ],
            capture_output=True,
            text=True,
            check=False,
        )

    def test_coercions_and_text_controls(self):
        header = ["key", "postal", "account", "token", "literal", "sku", "note"]
        row1 = ["K1", "00123", "12345678901234567", "2024-03-01", "=1+1", "123E5", "a"]
        row2 = ["K2", "00007", "23456789012345678", "2024-07-09", "=2+2", "456E5", "a"]
        self.write_csv(header, row1, row2)
        workbook, sheet = self.write_workbook(
            header,
            ["K1", 123, 12345678901234500, date(2024, 3, 1), "=1+1", 12300000, "a"],
            [*row2[:-1], "b"],
        )
        sheet["B2"].number_format = "00000"
        sheet["E3"].data_type = "s"
        workbook.save(self.xlsx_path)
        options = [
            argument for column in header[1:] for argument in ("--column", column)
        ]

        result = self.run_cli("--sheet", "Data", "--key", "key", *options, "--json")

        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertEqual(
            [
                (item["column"], item["code"], item["xlsx_row"])
                for item in json.loads(result.stdout)
            ],
            [
                ("postal", "leading-zeros", 2),
                ("account", "long-integer-changed", 2),
                ("token", "date-or-time", 2),
                ("literal", "formula", 2),
                ("sku", "scientific-notation", 2),
                ("note", "text-changed", 3),
            ],
        )
        self.assertNotIn("12345678901234567", result.stdout)

    def test_key_tracks_reordered_columns_and_rows(self):
        self.write_csv(["key", "postal"], ["K1", "001"], ["K2", "002"])
        workbook, sheet = self.write_workbook(
            ["postal", "key"], ["002", "K2"], ["001", "K1"]
        )
        sheet["B6"].number_format = "@"
        workbook.save(self.xlsx_path)
        result = self.run_cli("--key", "key", "--column", "postal", "--json")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout), [])

    def test_positional_audit_keeps_trailing_empty_text_record(self):
        self.write_csv(["identifier"], ["001"], [""])
        workbook, sheet = self.write_workbook(["identifier"], ["001"], [""])
        sheet["A4"].number_format = "@"
        workbook.save(self.xlsx_path)

        result = self.run_cli("--column", "identifier", "--json")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout), [])

    def test_key_reports_both_missing_and_added_rows(self):
        self.write_csv(["key", "postal"], ["K1", "001"], ["K2", "002"])
        workbook, _ = self.write_workbook(
            ["key", "postal"], ["K2", "002"], ["K3", "003"]
        )
        workbook.save(self.xlsx_path)
        result = self.run_cli("--key", "key", "--column", "postal", "--json")
        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertEqual(
            json.loads(result.stdout),
            [
                {
                    "code": "xlsx-row-added",
                    "csv_row": None,
                    "xlsx_row": 3,
                    "column": None,
                },
                {
                    "code": "csv-row-missing",
                    "csv_row": 2,
                    "xlsx_row": None,
                    "column": None,
                },
            ],
        )

    def test_duplicate_keys_and_unequal_positional_counts_fail(self):
        self.write_csv(["key", "postal"], ["K1", "001"])
        workbook, _ = self.write_workbook(
            ["key", "postal"], ["K1", "001"], ["K1", "002"]
        )
        workbook.save(self.xlsx_path)
        for options, message in (
            (["--key", "key"], "duplicates a key"),
            ([], "row counts differ"),
        ):
            with self.subTest(options=options):
                result = self.run_cli(*options, "--column", "postal")
                self.assertEqual(result.returncode, 2)
                self.assertIn(message, result.stderr)

    def test_sheet_selection_and_simple_numeric_type_change(self):
        self.write_csv(["identifier"], ["42"])
        workbook, _ = self.write_workbook(["identifier"], [42])
        workbook.create_sheet("Another")
        workbook.save(self.xlsx_path)
        result = self.run_cli("--column", "identifier")
        self.assertEqual(result.returncode, 2)
        self.assertIn("specify --sheet", result.stderr)
        result = self.run_cli("--sheet", "Data", "--column", "identifier")
        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertIn("identifier: number", result.stdout)

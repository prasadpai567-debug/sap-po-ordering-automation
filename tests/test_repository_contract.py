from __future__ import annotations

import csv
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class RepositoryContractTests(unittest.TestCase):
    def test_required_public_files_exist(self):
        for relative_path in ("Module1.bas", "README.md", "sample_input.csv"):
            self.assertTrue((ROOT / relative_path).is_file(), relative_path)

    def test_vba_source_contract(self):
        source = (ROOT / "Module1.bas").read_text(encoding="utf-8")
        required = (
            "Option Explicit",
            "Sub SAP_Final_Solution()",
            "Function GetMD04Sub",
            "Function SafeGet",
            "Function ExtractPONumber",
        )
        for declaration in required:
            self.assertIn(declaration, source)
        self.assertNotIn('GetMD04Sub = "0750"', source)
        self.assertNotIn("The focused UI element", source)
        self.assertTrue(source.rstrip().endswith("End Function"))

    def test_sample_csv_shape(self):
        with (ROOT / "sample_input.csv").open(newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            self.assertEqual(reader.fieldnames, ["Material", "MRP Area", "PO Number"])
            rows = list(reader)
        self.assertGreaterEqual(len(rows), 1)
        for row in rows:
            self.assertTrue(row["Material"])
            self.assertTrue(row["MRP Area"])

    def test_public_repository_excludes_workbook_binaries(self):
        excluded_suffixes = {".xlsm", ".xlsb"}
        for path in ROOT.rglob("*"):
            if path.is_file() and ".git" not in path.parts:
                self.assertNotIn(path.suffix.lower(), excluded_suffixes, str(path))


if __name__ == "__main__":
    unittest.main()

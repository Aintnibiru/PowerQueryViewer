import base64
import tempfile
import unittest
import zipfile
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch

from power_query_extractor import (
    _extract_mashup_package,
    _split_query_definitions,
    ask_confirmation,
    find_power_query_files,
    main,
    sanitize_filename,
    save_power_query_files,
)


class TestPowerQueryExtractor(unittest.TestCase):
    def test_split_standard_queries(self):
        content = """
section Section1;

shared Query1 = let
    Source = 1
in
    Source;

shared Query2 = "Hello World";
"""
        queries = _split_query_definitions(content)
        self.assertEqual(len(queries), 2)
        self.assertEqual(queries[0][0], "Query1")
        self.assertTrue(queries[0][1].startswith("let"))
        self.assertEqual(queries[1][0], "Query2")
        self.assertEqual(queries[1][1], '"Hello World"')

    def test_split_quoted_identifiers(self):
        content = """
section Section1;
shared #"Sales Summary (2024)" = 100;
shared #"Report ""Special"" Version" = 200;
"""
        queries = _split_query_definitions(content)
        self.assertEqual(len(queries), 2)
        self.assertEqual(queries[0][0], "Sales Summary (2024)")
        self.assertEqual(queries[0][1], "100")
        self.assertEqual(queries[1][0], 'Report "Special" Version')
        self.assertEqual(queries[1][1], "200")

    def test_split_parameter_and_attributes(self):
        content = """
section Section1;

[IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]
shared ServerName = "db.example.com" meta [IsParameterQuery=true];

[Documentation.Name="Helper", Meta=[Deep=[A=1]]]
shared ProcessData = (x) => x * 2;
"""
        queries = _split_query_definitions(content)
        self.assertEqual(len(queries), 2)
        self.assertEqual(queries[0][0], "ServerName")
        self.assertEqual(queries[0][1], '"db.example.com" meta [IsParameterQuery=true]')
        self.assertEqual(queries[1][0], "ProcessData")
        self.assertEqual(queries[1][1], "(x) => x * 2")

    def test_split_unshared_queries(self):
        content = """
section Section1;
InternalHelper = (a, b) => a + b;
shared PublicQuery = InternalHelper(1, 2);
"""
        queries = _split_query_definitions(content)
        self.assertEqual(len(queries), 2)
        self.assertEqual(queries[0][0], "InternalHelper")
        self.assertEqual(queries[0][1], "(a, b) => a + b")
        self.assertEqual(queries[1][0], "PublicQuery")
        self.assertEqual(queries[1][1], "InternalHelper(1, 2)")

    def test_split_semicolons_in_strings_and_comments(self):
        content = """
section Section1;
// Line comment with semicolon; and "quotes"
/* Block comment with ; semicolon and // nested comment */
shared ComplexQuery = let
    TextVal = "String with ; semicolon and \"\"escaped quotes\"\"",
    ListVal = {1, 2, 3},
    RecordVal = [ColA = "val;ue", ColB = 10]
in
    RecordVal;
"""
        queries = _split_query_definitions(content)
        self.assertEqual(len(queries), 1)
        self.assertEqual(queries[0][0], "ComplexQuery")
        self.assertIn('"String with ; semicolon', queries[0][1])

    def test_split_utf8_bom_handled(self):
        content = "\ufeffsection Section1;\nshared FirstQuery = 42;"
        queries = _split_query_definitions(content)
        self.assertEqual(len(queries), 1)
        self.assertEqual(queries[0][0], "FirstQuery")
        self.assertEqual(queries[0][1], "42")

    def test_sanitize_filename(self):
        self.assertEqual(sanitize_filename("NormalName"), "NormalName")
        self.assertEqual(sanitize_filename("Invalid:<>|?*Name"), "Invalid______Name")
        self.assertEqual(sanitize_filename("  "), "PowerQuery_1")
        self.assertEqual(sanitize_filename("CON"), "_CON")
        self.assertEqual(sanitize_filename("aux.pq"), "_aux.pq")
        self.assertEqual(sanitize_filename("NUL"), "_NUL")

    def test_save_power_query_files_collision_handling(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_dir = Path(tmp_dir)
            queries = [
                ("Sales/Data", "let x = 1 in x"),
                ("Sales:Data", "let y = 2 in y"),
                ("sales_data", "let z = 3 in z"),
            ]
            saved = save_power_query_files(queries, out_dir)
            self.assertEqual(len(saved), 3)

            filenames = [p.name for p in saved]
            self.assertEqual(len({f.casefold() for f in filenames}), 3)
            self.assertEqual(filenames[0], "Sales_Data.pq")
            self.assertEqual(filenames[1], "Sales_Data (2).pq")
            self.assertEqual(filenames[2], "sales_data (3).pq")

            # Verify contents are preserved and not overwritten
            self.assertEqual(saved[0].read_text(encoding="utf-8"), "let x = 1 in x")
            self.assertEqual(saved[1].read_text(encoding="utf-8"), "let y = 2 in y")
            self.assertEqual(saved[2].read_text(encoding="utf-8"), "let z = 3 in z")

    def test_extract_mashup_package_with_header(self):
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            zf.writestr("Formulas/Section1.m", "section Section1; shared Q1 = 1;")
        zip_data = zip_buf.getvalue()

        # Construct package with 4-byte length prefix
        package_blob = len(zip_data).to_bytes(4, "little") + zip_data + b"trailing"
        extracted = _extract_mashup_package(package_blob)
        self.assertIsNotNone(extracted)
        self.assertEqual(extracted, zip_data)

    def _create_dummy_excel(self, destination: Path) -> None:
        inner_zip_buf = BytesIO()
        with zipfile.ZipFile(inner_zip_buf, "w") as inner_zf:
            m_code = (
                "section Section1;\n"
                '[IsParameterQuery=true] shared Param1 = "abc";\n'
                'shared #"Total Revenue" = 5000;\n'
            )
            inner_zf.writestr("Formulas/Section1.m", m_code.encode("utf-8-sig"))
        inner_zip_bytes = inner_zip_buf.getvalue()

        mashup_blob = (
            len(inner_zip_bytes).to_bytes(4, "little")
            + inner_zip_bytes
            + b"permissions_or_metadata"
        )
        mashup_b64 = base64.b64encode(mashup_blob).decode("ascii")

        custom_xml = (
            f'<root xmlns:d="http://schemas.microsoft.com/DataMashup">'
            f"<d:DataMashup>{mashup_b64}</d:DataMashup>"
            f"</root>"
        ).encode()

        with zipfile.ZipFile(destination, "w") as outer_zf:
            outer_zf.writestr("customXml/item1.xml", custom_xml)
            outer_zf.writestr("[Content_Types].xml", b"<Types/>")

    def test_find_power_query_files_end_to_end(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            excel_path = Path(tmp_dir) / "TestWorkbook.xlsx"
            self._create_dummy_excel(excel_path)

            queries = find_power_query_files(excel_path)
            self.assertEqual(len(queries), 2)
            query_dict = dict(queries)
            self.assertIn("Param1", query_dict)
            self.assertEqual(query_dict["Param1"], '"abc"')
            self.assertIn("Total Revenue", query_dict)
            self.assertEqual(query_dict["Total Revenue"], "5000")

    def test_ask_confirmation(self):
        with patch("builtins.input", return_value="y"):
            self.assertTrue(ask_confirmation("Continue?"))
        with patch("builtins.input", return_value=""):
            self.assertTrue(ask_confirmation("Continue?"))
        with patch("builtins.input", return_value="n"):
            self.assertFalse(ask_confirmation("Continue?"))
        with patch("builtins.input", side_effect=EOFError):
            self.assertFalse(ask_confirmation("Continue?"))

    def test_main_cli_list_and_extract(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            excel_path = Path(tmp_dir) / "Demo.xlsx"
            out_dir = Path(tmp_dir) / "out_pq"
            self._create_dummy_excel(excel_path)

            # Test --list flag (should display queries without asking confirmation or creating files)
            with (
                patch("sys.argv", ["script", str(excel_path), "--list"]),
                patch("sys.stdout", new_callable=StringIO) as mock_stdout,
            ):
                main()
                output = mock_stdout.getvalue()
                self.assertIn("Daftar query yang ditemukan (2):", output)
                self.assertIn("1. Param1", output)
                self.assertIn("2. Total Revenue", output)
                self.assertFalse(out_dir.exists())

            # Test user cancellation ('n')
            with (
                patch(
                    "sys.argv",
                    ["script", str(excel_path), "-o", str(out_dir)],
                ),
                patch("builtins.input", return_value="n"),
                patch("sys.stdout", new_callable=StringIO) as mock_stdout,
            ):
                main()
                output = mock_stdout.getvalue()
                self.assertIn("Ekstraksi dibatalkan oleh pengguna.", output)
                self.assertFalse(out_dir.exists())

            # Test user confirmation ('y')
            with (
                patch(
                    "sys.argv",
                    ["script", str(excel_path), "-o", str(out_dir)],
                ),
                patch("builtins.input", return_value="y"),
                patch("sys.stdout", new_callable=StringIO) as mock_stdout,
            ):
                main()
                output = mock_stdout.getvalue()
                self.assertIn("Daftar query yang ditemukan (2):", output)
                self.assertIn("Mengekstrak query...", output)
                self.assertIn("Jumlah file tersimpan: 2", output)
                self.assertTrue((out_dir / "Param1.pq").exists())
                self.assertTrue((out_dir / "Total Revenue.pq").exists())

            # Test auto-yes flag (-y)
            out_dir_yes = Path(tmp_dir) / "out_yes"
            with (
                patch(
                    "sys.argv",
                    ["script", str(excel_path), "-o", str(out_dir_yes), "-y"],
                ),
                patch("sys.stdout", new_callable=StringIO) as mock_stdout,
            ):
                main()
                output = mock_stdout.getvalue()
                self.assertIn("Mengekstrak query...", output)
                self.assertTrue((out_dir_yes / "Param1.pq").exists())

    def test_file_validation_errors(self):
        with self.assertRaises(FileNotFoundError):
            find_power_query_files("non_existent_file.xlsx")

        with tempfile.TemporaryDirectory() as tmp_dir:
            invalid_ext = Path(tmp_dir) / "test.txt"
            invalid_ext.write_text("hello", encoding="utf-8")
            with self.assertRaises(ValueError):
                find_power_query_files(invalid_ext)

            bad_zip = Path(tmp_dir) / "corrupt.xlsx"
            bad_zip.write_text("not a real zip", encoding="utf-8")
            with self.assertRaises(ValueError):
                find_power_query_files(bad_zip)


if __name__ == "__main__":
    unittest.main()

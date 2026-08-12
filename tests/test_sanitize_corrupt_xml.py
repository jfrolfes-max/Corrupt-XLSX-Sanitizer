import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from sanitize_corrupt_xml import process_corrupt_xml, sanitize_to_ascii


class SanitizeCorruptXmlTests(unittest.TestCase):
    def test_sanitize_to_ascii_removes_non_ascii(self):
        data = "abcé£xyz".encode("utf-8")
        self.assertEqual(sanitize_to_ascii(data), "abcxyz")

    def test_process_corrupt_xml_creates_xlsx_in_sanitized_data_folder(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            xml_path = tmp_path / "input.xml"
            output_dir = tmp_path / "sanitized data"

            xml_bytes = (
                b"<rows><row><name>J\xc3\xb8hn</name><city>Ci\xc3\xa9ty</city></row>"
                b"<row><name>Ann\xffa</name><city>LA</city></row></rows>"
            )
            xml_path.write_bytes(xml_bytes)

            output_file = process_corrupt_xml(xml_path, output_dir)

            self.assertTrue(output_file.exists())
            self.assertEqual(output_file.parent.name, "sanitized data")

            with ZipFile(output_file, "r") as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")

            self.assertIn("name", sheet_xml)
            self.assertIn("city", sheet_xml)
            self.assertIn("Jhn", sheet_xml)
            self.assertIn("City", sheet_xml)
            self.assertIn("Anna", sheet_xml)


if __name__ == "__main__":
    unittest.main()

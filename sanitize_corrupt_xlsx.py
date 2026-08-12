from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile


def sanitize_to_ascii(raw_bytes: bytes) -> str:
    decoded = raw_bytes.decode("utf-8", errors="ignore")
    return "".join(ch for ch in decoded if ord(ch) < 128)


def _parse_xml_part(raw_bytes: bytes, part_name: str) -> ET.Element:
    sanitized_xml = sanitize_to_ascii(raw_bytes)
    try:
        return ET.fromstring(sanitized_xml)
    except ET.ParseError as exc:
        raise ValueError(
            f"The XLSX part {part_name!r} could not be parsed after sanitization."
        ) from exc


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _shared_strings(root: ET.Element) -> list[str]:
    return [
        "".join(node.itertext()).strip()
        for node in root
        if _local_name(node.tag) == "si"
    ]


def _column_index(cell_reference: str) -> int:
    letters = re.match(r"[A-Z]+", cell_reference.upper())
    if letters is None:
        return 1

    index = 0
    for letter in letters.group():
        index = index * 26 + ord(letter) - ord("A") + 1
    return index


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = next(
        (child for child in cell if _local_name(child.tag) == "v"), None
    )
    value = "" if value_node is None else (value_node.text or "")

    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (IndexError, ValueError):
            return ""
    if cell_type == "inlineStr":
        text_node = next(
            (child for child in cell if _local_name(child.tag) == "is"), None
        )
        return "" if text_node is None else "".join(text_node.itertext()).strip()
    return value.strip()


def _worksheet_rows(root: ET.Element, shared_strings: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for row_node in root.iter():
        if _local_name(row_node.tag) != "row":
            continue

        cells: dict[int, str] = {}
        for cell in row_node:
            if _local_name(cell.tag) != "c":
                continue
            column = _column_index(cell.attrib.get("r", ""))
            cells[column] = _cell_value(cell, shared_strings)

        if cells:
            rows.append([cells.get(column, "") for column in range(1, max(cells) + 1)])
    return rows


def _col_name(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _xlsx_sheet_xml(rows: Iterable[Iterable[str]]) -> str:
    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    lines.append(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    )

    for row_index, row in enumerate(rows, start=1):
        lines.append(f'<row r="{row_index}">')
        for col_index, value in enumerate(row, start=1):
            ref = f"{_col_name(col_index)}{row_index}"
            text = escape(str(value))
            lines.append(
                f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'
            )
        lines.append("</row>")

    lines.append("</sheetData></worksheet>")
    return "".join(lines)


def write_xlsx(rows: list[list[str]], destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    safe_rows = rows if rows else [[""]]

    with ZipFile(destination, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>',
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            '</Relationships>',
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="SanitizedData" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
            '</Relationships>',
        )
        archive.writestr(
            "xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
            '<cellXfs count="1"><xf xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>',
        )
        archive.writestr("xl/worksheets/sheet1.xml", _xlsx_sheet_xml(safe_rows))


def process_xlsx(source_file: Path, output_dir: Path) -> Path:
    try:
        with ZipFile(source_file, "r") as archive:
            worksheet_root = _parse_xml_part(
                archive.read("xl/worksheets/sheet1.xml"),
                "xl/worksheets/sheet1.xml",
            )
            shared_strings = []
            if "xl/sharedStrings.xml" in archive.namelist():
                shared_strings = _shared_strings(
                    _parse_xml_part(
                        archive.read("xl/sharedStrings.xml"),
                        "xl/sharedStrings.xml",
                    )
                )
    except (BadZipFile, KeyError) as exc:
        raise ValueError(
            "The input file is not a valid XLSX file with xl/worksheets/sheet1.xml."
        ) from exc

    rows = _worksheet_rows(worksheet_root, shared_strings)
    destination = output_dir / f"{source_file.stem}_sanitized.xlsx"
    write_xlsx(rows, destination)
    return destination


def process_xlsx_folder(
    source_dir: Path, output_dir: Path | None = None
) -> list[Path]:
    if not source_dir.is_dir():
        raise ValueError(f"Input path is not a folder: {source_dir}")

    xlsx_files = sorted(
        path
        for path in source_dir.iterdir()
        if path.is_file() and path.suffix.lower() == ".xlsx"
    )
    if not xlsx_files:
        raise ValueError(f"No .xlsx files found in folder: {source_dir}")

    destination_dir = output_dir or source_dir.parent / "sanitized data"
    return [process_xlsx(source_file, destination_dir) for source_file in xlsx_files]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sanitize every XLSX file in a folder and export the results."
    )
    parser.add_argument(
        "input_folder", type=Path, help="Folder containing corrupt XLSX files"
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Directory where sanitized XLSX files will be written; defaults to a sibling folder named 'sanitized data'",
    )
    args = parser.parse_args()

    output_paths = process_xlsx_folder(args.input_folder, args.output_dir)
    for output_path in output_paths:
        print(f"Sanitized XLSX written to: {output_path}")


if __name__ == "__main__":
    main()

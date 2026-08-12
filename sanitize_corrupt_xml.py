from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


def sanitize_to_ascii(raw_bytes: bytes) -> str:
    decoded = raw_bytes.decode("utf-8", errors="ignore")
    return "".join(ch for ch in decoded if ord(ch) < 128)


def _collect_table_rows(root: ET.Element) -> list[list[str]]:
    children = list(root)
    if not children:
        return [[root.tag, (root.text or "").strip()]]

    repeated = len({child.tag for child in children}) == 1
    structured = all(len(list(child)) > 0 for child in children)

    if repeated and structured:
        headers: list[str] = []
        seen: set[str] = set()

        for row_node in children:
            for cell in list(row_node):
                if cell.tag not in seen:
                    seen.add(cell.tag)
                    headers.append(cell.tag)

        rows: list[list[str]] = [headers]
        for row_node in children:
            values = {
                cell.tag: "".join(cell.itertext()).strip() for cell in list(row_node)
            }
            rows.append([values.get(header, "") for header in headers])
        return rows

    rows = [["path", "value"]]

    def walk(node: ET.Element, path: str) -> None:
        text = (node.text or "").strip()
        if text:
            rows.append([path, text])
        for child in list(node):
            walk(child, f"{path}/{child.tag}")

    walk(root, root.tag)
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


def process_corrupt_xml(source_file: Path, output_dir: Path) -> Path:
    raw_bytes = source_file.read_bytes()
    sanitized_xml = sanitize_to_ascii(raw_bytes)

    try:
        root = ET.fromstring(sanitized_xml)
    except ET.ParseError as exc:
        raise ValueError(
            "The file could not be parsed as XML after non-ASCII sanitization."
        ) from exc

    rows = _collect_table_rows(root)
    destination = output_dir / f"{source_file.stem}_sanitized.xlsx"
    write_xlsx(rows, destination)
    return destination


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sanitize a corrupt XML file and export its data to XLSX."
    )
    parser.add_argument("input_xml", type=Path, help="Path to the corrupt XML file")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("sanitized data"),
        help="Directory where sanitized XLSX output will be written",
    )
    args = parser.parse_args()

    output_path = process_corrupt_xml(args.input_xml, args.output_dir)
    print(f"Sanitized XLSX written to: {output_path}")


if __name__ == "__main__":
    main()

# Corrupt-XLSX-Sanitizer

This repository includes a Python script that sanitizes every `.xlsx` file in a folder by removing non-ASCII characters from worksheet XML and shared strings, then exporting each result to a new `.xlsx` file.

## Usage

```bash
python sanitize_corrupt_xml.py "/absolute/path/to/xlsx files"
```

By default, output files are written to a new sibling folder named `sanitized data`. For example, files in `/data/xlsx files` are written to `/data/sanitized data`.

Optional output directory:

```bash
python sanitize_corrupt_xml.py "/absolute/path/to/xlsx files" --output-dir "/absolute/path/to/sanitized data"
```

The input folder must contain XLSX ZIP workbooks with `xl/worksheets/sheet1.xml`. Every `.xlsx` file directly inside the folder is processed. Quote paths that contain spaces, such as:

```bash
python sanitize_corrupt_xml.py "Corrupt XML files"
```

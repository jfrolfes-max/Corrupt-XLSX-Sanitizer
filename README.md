# Corrupt-XML-Parser

This repository includes a Python script that sanitizes corrupt XML files by removing non-ASCII characters and exporting the parsed data to a new `.xlsx` file.

## Usage

```bash
python sanitize_corrupt_xml.py /absolute/path/to/corrupt.xml
```

By default, output files are written to a folder named `sanitized data`.

Optional output directory:

```bash
python sanitize_corrupt_xml.py /absolute/path/to/corrupt.xml --output-dir "/absolute/path/to/sanitized data"
```

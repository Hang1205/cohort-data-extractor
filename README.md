# Cohort Data Extractor

An offline desktop GUI and command-line tool for extracting every matching record for a selected set of IDs from multiple CSV, TSV and Excel tables. Use it for research cohorts, customer/account subsets, device logs, inventories or other datasets with a shared entity identifier.

The selected-ID column and each source file's ID column can have different names. No patient-specific column name, diagnosis, cohort size or filename is required. There is no default expected-count restriction. `record_id` is only a configurable starting value.

## Download and run

**Windows users: download `Cohort_Data_Extractor_Windows_Desktop_v1.0.0.zip` from [Releases](https://github.com/Hang1205/cohort-data-extractor/releases/latest).** Right-click → Extract All → open `Cohort_Data_Extractor.exe`. Keep `_internal` beside the EXE. The desktop ZIP includes Python, Tkinter, Excel libraries, complete editable source and synthetic examples. You do not need to install Python.

GitHub's **Code → Download ZIP** downloads the editable source, which requires Python 3.12 x64 with Tkinter on Windows. Launch it with `Start_Windows.cmd` or `python cohort_extract_gui.py`. Excel libraries are bundled in `vendor.zip` (or `vendor/`), so running the extracted source does not need pip or internet. On macOS/Linux, install a Python version with Tkinter and run `python3 cohort_extract_gui.py`; actual execution on those operating systems has not been verified.

## GUI workflow

1. Choose a CSV/TSV/`.xlsx` selected-ID list. Read headers and select its ID column. Use automatic or manual header rows; choose an Excel worksheet if needed.
2. Optional: choose a cohort-label column and explicit accepted values to select a subset from a larger list. Matching label values is exact after trimming surrounding spaces. Leave expected count blank, or set a count to validate the selection.
3. Choose the folder containing source tables. Under **Files & ID columns**, click **Read files**. Include/exclude individual files. Double-click a file, choose its worksheet, header row, ID column position, encoding and delimiter, then Save. Duplicate headings can be distinguished by column position.
4. Test ID matching on the first 10,000 rows per included file. A zero sample does not establish absence in later rows. Exact ID matching is the default. Explicit optional modes ignore case and/or remove a numeric integer `.0` suffix. They preserve leading zeros and reject modes that merge different selected IDs.
5. Check inputs, choose a new output folder outside the source folder, and Extract. Processing is streamed and the GUI reports progress. Existing destinations and source files are not overwritten.

Other choices include recursive subfolders, include/exclude filename patterns, strict/tolerant unrelated headings, and explicit skipping of files with unusable headers. All skips/exclusions are audited. Malformed records stop the run; rows are never silently discarded.

## Outputs and matching diagnostics

- `data/`: matching source records, preserving column order, duplicate records and row order.
- `file_summary.csv`: per-file counts and status.
- `matching_diagnostics.csv`: exact, case-insensitive and integer-suffix match counts per configured ID column.
- `id_coverage.csv`: selected IDs and matched-record totals.
- `missing_ids.csv`: selected IDs absent from every processed file.
- `audit.json`: selected count, per-file mapping, matching mode, input settings, skips and completion status.

CSV/TSV values are retained, although quoting/encoding can change. Excel `.xlsx` sources become CSV: literal cell values/formulas are serialized; dates use ISO text; workbook formatting/layout and cached formula results are not retained. Formula/error cells in the ID column are rejected. Convert legacy `.xls` or encrypted workbooks locally before use. Excel and CSV sources with colliding output filenames are rejected rather than overwritten.

IDs that look alike can differ as stored text. `00123` remains different from `123`. Integer `.0` normalization is optional; scientific notation and approximate/fuzzy joins are not supported. Correct column mapping and a shared ID domain remain the user's responsibility. This tool is a table-subsetting tool, not a relational join engine or anonymizer. No records are uploaded or fetched online.

## Command line

```bash
python extract_cli.py --source examples/source --ids examples/selected_ids.csv --output /path/to/new/output --id-column record_id --selected-id-column record_id --expected 5
```

For multiple mappings, use `--file-options file_options.example.json`. Other options include `--sheet`, `--selection-header-row`, `--source-header-row`, `--filter-column`, `--filter-values`, `--id-mode`, `--include`, `--exclude`, `--recursive`, and `--skip-invalid-files`. Run `python extract_cli.py --help` for details. `--expected 0` disables the count check (default).

## Fictional demo, tests and rebuild

`examples.zip` contains the fictional data and is unpacked locally by Load fictional demo if the examples folder is absent.

**Load fictional demo** selects five IDs from ten source entities and three tables. Expected output: nine records. No real data is included.

```bash
python -m unittest discover -s . -p test_extractor.py
python desktop_app.py --self-test --report self_test.json
```

The original 38 synthetic tests cover mixed ID columns, per-file worksheets, exact/optional normalization, collisions, sample matching, header detection, encodings, exclusions, malformed records, cancellation and failure without publishing incomplete output. Additional general-purpose tests cover unrestricted cohort sizes and the customer-style demo. Packaged Windows tests run with external Python paths disabled.

To build the Windows x64 EXE: create a Python 3.12 virtual environment, install `requirements_build_windows.lock.txt`, then run `python build_windows.py`. Installation/rebuilding needs a package source; the distributed desktop app runs offline. Packaging and source archive scripts are in `pack_release.py`.

## License and citation

Application code is MIT licensed, copyright Hang Xu. Vendored openpyxl (3.1.5) and et_xmlfile (2.0.0) carry their own notices/licenses. Desktop distributions include Python and packaging/runtime license notices. See `THIRD_PARTY_NOTICES.md` and `LICENSE`.

Citation metadata is in `CITATION.cff`; this software release has no claimed DOI. Suggested citation: Hang Xu (2026). *Cohort Data Extractor*, version 1.0.0. https://github.com/Hang1205/cohort-data-extractor

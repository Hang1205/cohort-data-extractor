# Cohort Data Extractor

An offline desktop GUI and command-line tool for extracting every matching record for a selected set of IDs from multiple CSV, TSV and Excel tables. Use it for research cohorts, customer/account subsets, device logs, inventories or other datasets with a shared entity identifier.

The selected-ID column and each source file's ID column can have different names. No patient-specific column name, diagnosis, cohort size or filename is required. There is no default expected-count restriction. `record_id` is only a configurable starting value.

## Download and run

**Windows users: download `Cohort_Data_Extractor_Windows_Desktop_v1.1.0.zip` from [Releases](https://github.com/Hang1205/cohort-data-extractor/releases/latest).** Right-click → Extract All → open `Cohort_Data_Extractor.exe`. Keep `_internal` beside the EXE. The desktop ZIP includes Python, Tkinter, Excel libraries, complete editable source and synthetic examples. You do not need to install Python.

GitHub's **Code → Download ZIP** downloads the editable source, which requires Python 3.12 x64 with Tkinter on Windows. Launch it with `Start_Windows.cmd` or `python cohort_extract_gui.py`. Excel libraries are bundled in `vendor.zip` (or `vendor/`), so running the extracted source does not need pip or internet. On macOS/Linux, install a Python version with Tkinter and run `python3 cohort_extract_gui.py`; actual execution on those operating systems has not been verified.

## GUI workflow

1. **Choose your ID list.** Select the CSV or Excel file and its ID column.
2. **Choose source files.** Choose a source folder, then use **Add files** or **Add another folder** for additional sources. All file choices are on the main screen. Select rows and choose **Filter selected**, **Copy whole selected**, or **Exclude selected**. Double-click a row to choose its ID column, Excel sheet or other import settings.
3. **Choose output and Extract.** Use a new folder outside every source folder. Open the Results tab for per-file progress and Open completed output when finished.

Files missing a usable ID column start **excluded** when first read. You may choose the correct column and Filter by ID, keep Exclude, or explicitly Copy whole file. The GUI also skips invalid headers by default and records every skip in the audit. Malformed data records still stop the run.

**Copy whole file copies every byte and every record, without patient filtering.** These files go into `supporting_files/`, carry an unfiltered flag in the audit and do not contribute to ID coverage. This is useful for provider lookup tables. Patient-specific extraction from files linked only through encounter/order/provider IDs needs a linking table; the app does not infer that relationship.

Files from additional folders retain separate folder namespaces, so equal filenames do not overwrite each other. Added folders are scanned when added; refresh your selection to include newly created files. Column settings and matching options are optional; start with the main screen. Use **Check matches** to compare the first 10,000 rows against your IDs. Zero sample matches do not establish absence later in the file. Exact text matching preserves leading zeros; optional case/`.0` normalization is explicit and detects selection collisions.

Other choices include recursive subfolders, include/exclude filename patterns, strict/tolerant unrelated headings, and explicit skipping of files with unusable headers. All skips/exclusions are audited. Malformed records stop the run; rows are never silently discarded.

## Outputs and matching diagnostics

- `data/`: matching source records, preserving column order, duplicate records and row order.
- `supporting_files/`: explicitly selected whole files, unfiltered and byte-for-byte.
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

For multiple mappings, repeat `--source-file /path/to/additional.csv` and use `--file-options file_options.example.json`. Other options include `--sheet`, `--selection-header-row`, `--source-header-row`, `--filter-column`, `--filter-values`, `--id-mode`, `--include`, `--exclude`, `--recursive`, and `--skip-invalid-files`. Run `python extract_cli.py --help` for details. `--expected 0` disables the count check (default).

## Fictional demo, tests and rebuild

`examples.zip` contains the fictional data and is unpacked locally by Load fictional demo if the examples folder is absent.

**Load fictional demo** selects five IDs from ten source entities and three tables. Expected output: nine records. No real data is included.

```bash
python -m unittest discover -s . -p test_extractor.py
python desktop_app.py --self-test --report self_test.json
```

The original 38 synthetic tests cover mixed ID columns, per-file worksheets, exact/optional normalization, collisions, sample matching, header detection, encodings, exclusions, malformed records, cancellation and failure without publishing incomplete output. Additional tests cover unrestricted cohort sizes, the demo, multiple source folders, duplicate filenames, whole-copy separation and coverage, and missing-ID skipping. 43 engine tests pass. Packaged Windows tests run with external Python paths disabled.

To build the Windows x64 EXE: create a Python 3.12 virtual environment, install `requirements_build_windows.lock.txt`, then run `python build_windows.py`. Installation/rebuilding needs a package source; the distributed desktop app runs offline. Packaging and source archive scripts are in `pack_release.py`.

## License and citation

Application code is MIT licensed, copyright Hang Xu. Vendored openpyxl (3.1.5) and et_xmlfile (2.0.0) carry their own notices/licenses. Desktop distributions include Python and packaging/runtime license notices. See `THIRD_PARTY_NOTICES.md` and `LICENSE`.

Citation metadata is in `CITATION.cff`; this software release has no claimed DOI. Suggested citation: Hang Xu (2026). *Cohort Data Extractor*, version 1.1.0. https://github.com/Hang1205/cohort-data-extractor

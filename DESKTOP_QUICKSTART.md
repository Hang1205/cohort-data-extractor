# Start here — v1.1.0

Extract ALL files from the Windows Desktop ZIP. Open Cohort_Data_Extractor.exe and keep _internal beside it. Python installation and internet are not required.

1. Choose your CSV/Excel ID list, worksheet and ID column.
2. Choose your main data folder. Add files or another folder as needed.
3. Select file rows: Filter selected uses the chosen ID column; Copy whole selected copies all records; Exclude selected omits the file. Double-click to choose a different ID column for each file.
4. Choose a new output folder outside every source folder. Extract.

Files without a usable ID start excluded. To include Providers.csv without patient IDs, explicitly choose Copy whole selected. It goes into supporting_files/ and contains all providers, not a patient-filtered subset. A linking-file join is not performed.

Results tab shows progress. Filtered records are in data/. Whole files are separate in supporting_files/. Review file_summary.csv and audit.json for exclusions/skips. ID coverage counts only filtered rows. All data stays on your computer.

For 842 SVP patients, select your already-filtered 842-ID list and IP_PATIENT_ID. Choose IP_PATIENT_ID separately for each patient table if the default record_id is absent. Missing patient IDs are never inferred.

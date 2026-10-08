# v1.1.0

- File choices moved to the main screen with Filter by ID, Copy whole file and Exclude.
- Add chosen CSV/Excel files or another folder; configure a separate column/sheet for each.
- Missing/unusable IDs start excluded; GUI defaults to audited skips of invalid headers.
- Whole files copied in cancellable chunks to supporting_files/, marked unfiltered and excluded from cohort coverage.
- Additional source folders have distinct output namespaces; all input files and the selected list must remain stable during extraction.
- 43 engine tests, GUI demo and frozen portable tests.

# Changelog

## 1.0.0 — 2026-10-08

First general-purpose release, derived from the selected-ID extraction workflow. Removes the fixed 842-patient requirement and domain-specific terminology. Supports arbitrary selected-ID columns, per-file CSV/Excel mapping and inclusion, explicit cohort filtering, header detection, streamed extraction, matching samples, safe optional normalization and audited outputs. Includes desktop GUI, CLI, synthetic examples, editable code and Windows packaging scripts.

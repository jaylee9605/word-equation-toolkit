# Validation examples

Run commands from the project directory. The reviewed `equation_sample.docx` is the unchanged v0.1 baseline: three pages, native line numbers 1–30, ten numbered display equations, four inline math objects and a five-row, two-column matrix in equation (7). Its expected result is exit code **0**, no detected errors or warnings, and SHA-256 `e8a9658a9a26184319ba0e0d2a673d4c36549c9fbffbbcfde6ce79c5d207ca84`.

```bash
python src/equation_audit.py examples/equation_sample.docx --json reports/my_sample_report.json
```

The files under `negative/` contain only generated elementary mathematics. Each deliberately exercises one defect category. They are not intended for writing or visual comparison.

| Fixture | Expected diagnostic | Exit code |
| --- | --- | --- |
| `negative/broken_bookmark.docx` | `UNMATCHED_BOOKMARK` | 1 |
| `negative/missing_ref_target.docx` | `MISSING_REF_TARGET` | 1 |
| `negative/malformed_table.docx` | `MALFORMED_EQUATION_TABLE` (and related width/row diagnostics) | 1 |
| `negative/missing_seq.docx` | `MISSING_SEQ` | 1 |
| `negative/visible_border.docx` | `VISIBLE_EQUATION_BORDER` | 1 |

```bash
python src/equation_audit.py examples/negative/broken_bookmark.docx --json reports/my_negative_report.json
```

Use `reports/examples/negative_cases.json` for expected diagnostics and fixture hashes. Reports distinguish stored fields/results from unverified live Word/visual checks. Running validation does not fix a negative fixture and must not change any input bytes. Regenerate synthetic fixtures/reports with `python tools/generate_examples.py` after installing development requirements.

# Word Equation Toolkit v0.1.0

First release of a read-only Python CLI for native Word equation documents.

## Included

- Native OMML inventory, matrix dimensions, SEQ Eq numbering, REF targets and bookmark integrity checks.
- Simple/complex fields, split instruction runs, nested fields and stored numeric cache checks.
- Three-column equation-table geometry, alignment and explicit border validation. Visible table, row-exception and cell borders report errors; uncertain inherited styles report unverified warnings.
- Readable summaries and JSON reports, documented exit codes, input hash checks and output-alias protection.
- Editable neutral DOCX sample, genuine DOTX template, five synthetic negative cases and example reports.
- 80 automated tests, Python 3.10–3.13 CI configuration and deterministic ZIP packaging with a SHA-256 file manifest.
- MIT License.

## Verification and limitations

Recorded local tests pass on Python 3.10.11, 3.11.9, 3.12.14 and 3.13.9. See `reports/compatibility.json` and the GitHub Actions tab for separate local and hosted evidence.

The sample/template retain their validated bytes: 14 native math objects, ten numbered equations, ten SEQ and REF fields, ten bookmark pairs, the matrix in equation (7) and continuous line numbering 1–30.

Static OOXML checks do not render Word, recalculate live fields or certify actual pagination, font appearance or the complete border/style cascade. Follow README manual Word checks after document edits or environment changes. No automatic repair, Word add-in, AI integration or cloud service is included.

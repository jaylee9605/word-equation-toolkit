# Word Equation Toolkit

A small, read-only Python CLI for checking native Word equations, equation numbering, cross-references and borderless equation tables. It reports stored OOXML structure and cached field results. It does not render Word documents, update live fields or repair files.

## Requirements and quick start

Compatibility target: **Python 3.10–3.13**. Runtime dependency: `lxml`. Development/test dependency: `pytest`. Use a virtual environment if appropriate.

```bash
python -m pip install -r requirements.txt
python src/equation_audit.py examples/equation_sample.docx
python src/equation_audit.py examples/equation_sample.docx --json reports/my_audit.json
```

Use `--json -` for JSON on stdout and a readable summary on stderr. Exit codes:

| Code | Meaning |
| --- | --- |
| 0 | No demonstrated static error; warnings/unverified checks may remain |
| 1 | One or more static errors detected |
| 2 | Input or output could not be processed |

`PASS (static OOXML only)` is not live Word certification. Warnings are visibly labeled `PASS WITH WARNINGS`; check each table's `borders.status` and all unverified checks. JSON reports must not overwrite/alias the input or use a Word document/template extension. Reports use atomic replacement in a trusted output directory. Validation compares the input SHA-256 before and after inspection and never saves it.

## Checks implemented

- Native `m:oMath` and display `m:oMathPara` inventory; matrix row/column consistency.
- Simple and complex Word fields, split instruction runs and nested fields; instruction text stays separate from cached results.
- `SEQ Eq` and `REF` fields; unique bookmark names, matched start/end markers, missing targets and numeric reference caches.
- Three-column numbered equation tables: unmerged cells, stored widths, symmetric outer columns, equation/number cell placement and paragraph alignment.
- Explicit table, row-exception and cell border declarations. Known visible line styles such as `single` and `double` produce `VISIBLE_EQUATION_BORDER` and exit 1.
- Informational section line-numbering configuration, including continuous numbering.
- Basic OPC container, relationship-target and XML safety checks. External relationships are flagged for review.

Supported stored numbering switches are ordinary increments, `\n`, `\r <integer>`, `\c`, `\h` and `\* ARABIC`. Unsupported/conflicting controls, hidden/missing caches and uncertain states remain unverified. Each story is evaluated separately; this is not a complete Word field evaluator.

### Border validation scope

The checker enforces explicitly borderless equation-table declarations. A static border pass requires all six physical table borders (`top`, `left`, `bottom`, `right`, `insideH`, `insideV`) to be explicitly `nil`/`none`, with no explicitly visible row/cell borders or unresolved style influence. Known unconditional inherited table borders can be suppressed by those direct overrides.

Absent/partial declarations, missing/cyclic styles, inherited cell/conditional borders, and unknown values yield `BORDER_STYLE_UNVERIFIED`, a warning and per-table status `unverified`. The checker does **not** implement Word's complete style cascade or border-conflict algorithm. It does not infer actual painted visibility from line width, color, themes, logical edges or renderer behavior. Conflicting explicit visible declarations are flagged even if a renderer might hide them. Verify inherited/ambiguous cases in Word.

## Reuse the Word sample

`examples/equation_sample.docx` is a three-page editable example with elementary mathematics: **14 OMML objects, ten numbered display equations, ten SEQ fields, ten REF fields and ten bookmark pairs**. Continuous native line numbering **1–30 is intentional**. Its equation (7) is a five-row, two-column native math matrix. Do not change it merely to accommodate artifacts from another renderer.

All nine equation tables use a symmetric 1,310 / 6,740 / 1,310-twip grid, invisible borders, centered equations and right-aligned numbers. One table has two numbered rows. `templates/equation-template.dotx` is a genuine OOXML template, not a renamed DOCX.

Edit a working copy or create a document from the template. Preserve native math, table cells and field codes. Add a real `SEQ Eq` field and a new unique bookmark for each added equation; use `REF` fields for references. Copying a row does not guarantee unique bookmarks. Update fields in Word and run the CLI after edits.

## Tests, CI and examples

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q --color=no
python tools/generate_examples.py
python src/equation_audit.py examples/negative/visible_border.docx --json -
```

The last command intentionally exits **1**. Five synthetic negative examples cover broken bookmarks, missing REF targets, malformed tables, missing SEQ fields and a visible top border. `reports/examples/` includes positive/negative reports and expected diagnostics. Generation never edits the reviewed example. See `examples/README.md`.

[GitHub Actions configuration](.github/workflows/ci.yml) runs pytest and the sample CLI on Python **3.10, 3.11, 3.12 and 3.13**, using read-only repository permissions and no deployment/publishing steps. It follows the official [setup-python](https://github.com/actions/setup-python) and [checkout](https://github.com/actions/checkout) actions. Hosted workflow results are available in the repository Actions tab when executed. Recorded local results do not establish hosted CI success. Actual local interpreter results are in `reports/compatibility.json`; `reports/test_results.txt` records the test run. Target versions are not blanket guarantees for every patch, platform or renderer.

## Build a distribution ZIP

```bash
python tools/package_release.py
# Or choose an explicit local destination:
python tools/package_release.py --output ../word-equation-toolkit-v0.1.0.zip
```

The default destination is `dist/word-equation-toolkit-v0.1.0.zip`. Source, tests, CI, documentation, neutral examples/template and curated technical evidence are included. Git data, caches, environments and temporary Word files are excluded. Report inclusion is allowlisted. Each included file has a SHA-256 in `RELEASE_MANIFEST.json` (the manifest excludes itself).

ZIP entry order, creator system, timestamps and permissions are normalized; repeated builds of identical content in the tested environment are byte-identical. Packaging does not run tests, certify privacy or publish anything. See `reports/release_readiness.md` for recorded verification and limitations.

## Manual Word verification

The unchanged sample was previously reviewed in Word 16.0, including a native field insertion/deletion exercise. This evidence is specific to the sample hash in `reports/word_sample_verification.md`; the CLI does not perform it.

After editing a document or moving to another Word/font environment, use a disposable copy:

1. Update main-story fields with **Ctrl+A**, **F9**, repeating if references retain old caches. Some keyboards require **Fn+F9**. Header/footer fields may need a separate update.
2. For the provided sample, confirm three pages and line labels **1–11 / 12–20 / 21–30**, equation numbers **(1)–(10)** and no reference-source errors.
3. Insert a native numbered equation after equation 1, with a real SEQ field and unique bookmark. Update twice: expect eleven numbers and existing later references shifted by one. Delete it and update twice: expect the original numbers/references again.
4. Inspect equation centering, right-aligned numbers, invisible borders and editable black symbols in the five-row matrix. Close the disposable copy without overwriting the distributed baseline.

## Limits

- Static XML cannot establish displayed page/line numbers, actual placement, fonts/glyphs, painted border appearance or live field updates. These remain explicitly unverified by the CLI.
- Inherited/conditional border resolution is deliberately conservative; see the border scope above. Cross-renderer matrix compatibility is not guaranteed.
- This is not full OOXML/schema validation or a semantic privacy scanner. Metadata inventory and relationship warnings do not certify another document free of identifying or embedded content.
- Table checks target the documented three-cell layout. Arbitrary math-bearing data tables, merged/floating/nested layouts, multi-column text, conditional style semantics, tracked changes, text boxes and Strict OOXML are outside the validated layout scope.
- Packages over 100 MiB uncompressed or 10,000 parts, DTDs, encrypted/duplicate ZIP parts and macro-enabled main types are rejected. External XML entity resolution and network loading are disabled. The CLI makes no network/API calls.
- Python 3.14+, other implementations and platforms beyond recorded results are not claimed tested. No Word installation is needed for the CLI or synthetic tests.

## License

MIT License; see `LICENSE` and `LICENSING.md`. Dependencies are installed separately under their own licenses.

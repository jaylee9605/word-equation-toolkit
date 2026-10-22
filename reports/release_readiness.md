# Release verification

This report records local verification of v0.1.0. Hosted CI results are separate and must be checked in the repository Actions tab.

| Acceptance criterion | Result |
| --- | --- |
| Explicit visible equation-table borders | Error for known visible table, row-exception and cell borders; independent top-border regression covered |
| Conservative inheritance | Missing/partial or unresolved declarations produce an unverified warning; complete Word style/conflict resolution is outside scope |
| Full local tests | 80 passed on each of Python 3.10.11, 3.11.9, 3.12.14 and 3.13.9; zero failures, skips or warnings |
| CI | GitHub Actions configured for 3.10–3.13, read-only permissions; hosted Ubuntu runs not executed |
| Baseline | SHA-256 unchanged; 14 OMML objects, ten numbered equations, ten SEQ and REF fields, ten bookmark pairs; matrix and continuous line-number configuration preserved |
| CLI | Reviewed sample: zero errors and warnings; all five synthetic negative cases return exit 1 with the expected diagnostic |
| Read-only behavior | Input byte/hash comparisons in tests and separate extracted-package checks |
| Public distribution | Curated reports only; no private source documents, private absolute paths, credentials or identifying document metadata found in included files |
| Packaging | Independently extracted, manifest hashes checked, full suite and example CLI rerun; two builds byte-identical in the tested environment |
| Separate final review | Implementation, regression expectations, documentation, CI permissions and package contents checked against these criteria, beyond test outcomes |

Evidence: `compatibility.json`, `test_results.txt`, `examples/` reports and `word_sample_verification.md`. The final archive contains an integrity manifest. Local compatibility evidence is Windows x64; hosted Linux results are not established by these local runs.

The CLI validates stored OOXML and caches. It cannot certify actual Word pagination, visible line labels, font rendering, full inherited borders or live field recalculation. The unchanged sample retains prior Word 16.0 verification; this update did not repeat a live Word session. For edited documents or another Word environment, follow the README's manual Word checklist using a disposable copy. No known unresolved implementation blocker remains within v0.1 scope.

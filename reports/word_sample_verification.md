# Reviewed Word sample identity

The distributed sample/template are retained byte-for-byte. This is sample-specific technical evidence, separate from the Python CLI's static results.

| Artifact | SHA-256 |
| --- | --- |
| `examples/equation_sample.docx` | `e8a9658a9a26184319ba0e0d2a673d4c36549c9fbffbbcfde6ce79c5d207ca84` |
| `templates/equation-template.dotx` | `995931a60733088bda6f7d6abbd6c5d6f64b3e2e2d2b409c3c46f5a787887a5c` |

The sample contains 14 native OMML objects (ten numbered display expressions and four inline), ten `SEQ Eq` fields, ten `REF` fields, ten matched unique bookmark pairs and nine three-column borderless tables. The matrix in equation (7) is five rows by two columns. Its native line-number configuration is continuous, count-by-one; it has not been rewritten as typed numbers.

An earlier native Word 16.0 review confirmed three pages, visible line labels 1–11 / 12–20 / 21–30, editable formulas and correct matrix appearance. Inserting a numbered native equation in a disposable copy produced eleven numbers and shifted the later references; deleting it and updating restored ten numbers and all references. Creating a new document from the genuine template also succeeded. This review applies to the above unchanged bytes, not arbitrary input documents or all renderers.

Core creator/last-modified-by values are empty. The sample/template contain neutral elementary mathematics, no embedded media/attachments, comments, tracked revisions, custom XML or external relationship targets. Distribution checks re-open the package and confirm this identity; they do not claim the CLI is a complete privacy scanner.

No new visual Word certification is implied by a passing static CLI result. Follow README's manual checks after edits or on a different Word/font environment. Another renderer's artifacts are not grounds to alter the reviewed matrix.

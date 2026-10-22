"""Build a deterministic distribution ZIP. Does not publish or contact services."""
from __future__ import annotations

import argparse
import ast
import hashlib
import json
import os
from pathlib import Path
import tempfile
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

REQUIRED = [
    "README.md", "LICENSE", "LICENSING.md", "RELEASE_NOTES.md",
    "requirements.txt", "requirements-dev.txt", "src/equation_audit.py",
    "tests/test_equation_audit.py", "tests/test_release_acceptance.py",
    "examples/equation_sample.docx", "templates/equation-template.dotx",
    ".github/workflows/ci.yml", "tests/test_equation_borders.py",
    "reports/release_readiness.md", "reports/word_sample_verification.md",
    "reports/test_results.txt", "reports/test_environment.json", "reports/compatibility.json",
]
ROOT_FILES = {".gitignore", "README.md", "LICENSE", "LICENSING.md", "RELEASE_NOTES.md",
              "requirements.txt", "requirements-dev.txt", ".github/workflows/ci.yml"}
FOLDERS = {"src", "tests", "tools", "examples", "templates", "reports"}
EXCLUDE = {".git", ".venv", ".pytest_cache", "__pycache__", "dist"}
PUBLIC_REPORTS = {
    "reports/release_readiness.md", "reports/word_sample_verification.md",
    "reports/test_results.txt", "reports/test_environment.json", "reports/compatibility.json",
    "reports/examples/valid_sample.json", "reports/examples/negative_cases.json",
    "reports/examples/broken_bookmark.json", "reports/examples/missing_ref_target.json",
    "reports/examples/malformed_table.json", "reports/examples/missing_seq.json",
    "reports/examples/visible_border.json",
}


def release_files(root):
    missing = [p for p in REQUIRED if not (root / p).is_file()]
    if missing:
        raise ValueError("Required release files missing: " + ", ".join(missing))
    files = []
    for path in root.rglob("*"):
        relative = path.relative_to(root)
        if any(p in EXCLUDE for p in relative.parts) or path.name.startswith(("~$", ".equation-audit-")):
            continue
        if path.is_symlink():
            raise ValueError("Release input must not contain symbolic links.")
        if not path.is_file():
            continue
        allowed = relative.as_posix() in ROOT_FILES or (
            relative.parts[0] in FOLDERS and path.suffix in {".py", ".md", ".txt", ".json", ".docx", ".dotx"}
        )
        if relative.parts[0] == "reports":
            allowed = relative.as_posix() in PUBLIC_REPORTS
        if allowed:
            files.append(path)
    return sorted(files, key=lambda p: p.relative_to(root).as_posix())


def build_archive(root: Path, output: Path):
    root, output = root.resolve(), output.resolve()
    if output.suffix.lower() != ".zip":
        raise ValueError("Release output must use a .zip filename.")
    files = release_files(root)
    if output in files:
        raise ValueError("Archive must not overwrite a release input.")
    tree = ast.parse((root / "src/equation_audit.py").read_text(encoding="utf-8"))
    version = next(ast.literal_eval(e.value) for e in tree.body if isinstance(e, ast.Assign)
                   and any(isinstance(t, ast.Name) and t.id == "VERSION" for t in e.targets))
    contents = {p.relative_to(root).as_posix(): p.read_bytes() for p in files}
    manifest = {"project": "Word Equation Toolkit", "version": version,
                "files": {name: hashlib.sha256(data).hexdigest() for name, data in contents.items()}}
    manifest_bytes = (json.dumps(manifest, indent=2) + "\n").encode("utf-8")
    contents["RELEASE_MANIFEST.json"] = manifest_bytes
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".zip.tmp", delete=False) as stream:
            temporary = Path(stream.name)
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
            for name in sorted(contents):
                info = ZipInfo("word-equation-toolkit/" + name, date_time=(1980, 1, 1, 0, 0, 0))
                info.create_system = 3
                info.compress_type = ZIP_DEFLATED
                info.external_attr = 0o644 << 16
                archive.writestr(info, contents[name])
        with ZipFile(temporary) as archive:
            if archive.testzip() is not None:
                raise ValueError("ZIP integrity verification failed.")
            for name, data in contents.items():
                if archive.read("word-equation-toolkit/" + name) != data:
                    raise ValueError("Archive content differs from the release snapshot.")
        os.replace(temporary, output)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    (root / "RELEASE_MANIFEST.json").write_bytes(manifest_bytes)
    return manifest


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    root = Path(__file__).resolve().parents[1]
    parser.add_argument("--output", type=Path, default=root / "dist/word-equation-toolkit-v0.1.0.zip")
    args = parser.parse_args(argv)
    try:
        manifest = build_archive(root, args.output)
        print(f"Created {args.output.name}; {len(manifest['files'])} files plus integrity manifest. No upload performed.")
        return 0
    except (OSError, ValueError, SyntaxError, StopIteration) as exc:
        print(f"Packaging failed: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

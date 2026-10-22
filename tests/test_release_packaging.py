"""Local packaging tests; no publication or remote operations."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
from zipfile import ZipFile

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from package_release import REQUIRED, build_archive


def release_tree(tmp_path):
    root = tmp_path / "project"
    for name in REQUIRED:
        p = root / name; p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text("synthetic packaging fixture\n", encoding="utf-8")
    (root / "src/equation_audit.py").write_text('VERSION = "0.1.0"\n', encoding="utf-8")
    return root


def test_zip_is_deterministic_and_manifest_matches(tmp_path):
    root = release_tree(tmp_path)
    first, second = tmp_path / "first.zip", tmp_path / "second.zip"
    build_archive(root, first)
    build_archive(root, second)
    assert first.read_bytes() == second.read_bytes()
    with ZipFile(first) as z:
        assert z.testzip() is None
        assert all(i.create_system == 3 and i.date_time == (1980, 1, 1, 0, 0, 0) for i in z.infolist())
        manifest = json.loads(z.read("word-equation-toolkit/RELEASE_MANIFEST.json"))
        for name, expected in manifest["files"].items():
            assert hashlib.sha256(z.read("word-equation-toolkit/" + name)).hexdigest() == expected


def test_zip_excludes_git_caches_and_environment(tmp_path):
    root = release_tree(tmp_path)
    for name in [".git/config", "tests/__pycache__/private.py", ".venv/module.py", "reports/.pytest_cache/token.json"]:
        p = root / name; p.parent.mkdir(parents=True, exist_ok=True); p.write_text("not for distribution")
    output = tmp_path / "review.zip"
    build_archive(root, output)
    with ZipFile(output) as z:
        assert not any(".git/" in n or "__pycache__" in n or ".venv" in n or ".pytest_cache" in n for n in z.namelist())


def test_packager_refuses_nonzip_output_without_touching_file(tmp_path):
    root = release_tree(tmp_path); original = root / "examples/equation_sample.docx"; before = original.read_bytes()
    with pytest.raises(ValueError, match=".zip"):
        build_archive(root, original)
    assert original.read_bytes() == before


def test_missing_release_material_is_reported(tmp_path):
    root = release_tree(tmp_path); (root / "RELEASE_NOTES.md").unlink()
    with pytest.raises(ValueError, match="RELEASE_NOTES"):
        build_archive(root, tmp_path / "review.zip")


def test_synthetic_examples_regenerate_deterministically(tmp_path):
    root = tmp_path / "copy"
    for name in ["tools/generate_examples.py", "tests/test_equation_audit.py", "src/equation_audit.py", "examples/equation_sample.docx"]:
        target = root / name; target.parent.mkdir(parents=True, exist_ok=True); shutil.copyfile(ROOT / name, target)
    baseline = (root / "examples/equation_sample.docx").read_bytes()
    snapshots = []
    for _ in range(2):
        run = subprocess.run([sys.executable, str(root / "tools/generate_examples.py")], capture_output=True, text=True)
        assert run.returncode == 0, run.stderr
        snapshots.append({p.relative_to(root).as_posix(): p.read_bytes()
                          for folder in [root / "examples/negative", root / "reports/examples"] for p in folder.rglob("*") if p.is_file()})
    assert snapshots[0] == snapshots[1]
    assert len(list((root / "examples/negative").glob("*.docx"))) == 5
    assert (root / "examples/equation_sample.docx").read_bytes() == baseline


def test_public_zip_omits_internal_reports_but_keeps_ci_and_evidence(tmp_path):
    root = release_tree(tmp_path)
    for name in ["reports/internal_review.json", "reports/word_verification.json", "reports/initial_audit.md",
                 ".github/workflows/unrelated_publish.yml"]:
        p = root / name; p.parent.mkdir(parents=True, exist_ok=True); p.write_text("internal only")
    output = tmp_path / "public.zip"
    build_archive(root, output)
    with ZipFile(output) as z:
        names = z.namelist()
        assert "word-equation-toolkit/.github/workflows/ci.yml" in names
        assert "word-equation-toolkit/reports/release_readiness.md" in names
        assert "word-equation-toolkit/reports/word_sample_verification.md" in names
        assert not any("internal_review" in n or "word_verification.json" in n or "initial_audit" in n or "unrelated_publish" in n for n in names)


def test_ci_declares_all_supported_python_versions_without_publish_steps():
    workflow = (ROOT / ".github/workflows/ci.yml").read_text(encoding="utf-8")
    assert 'python-version: ["3.10", "3.11", "3.12", "3.13"]' in workflow
    assert "python -m pytest -q" in workflow
    assert "python -m pip install -r requirements-dev.txt" in workflow
    assert "contents: read" in workflow and "persist-credentials: false" in workflow
    assert "pull_request_target" not in workflow and "secrets." not in workflow
    assert "twine" not in workflow and "gh release" not in workflow

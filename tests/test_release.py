import importlib.util
from pathlib import Path
from zipfile import ZipFile

import pytest

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("build_release", ROOT / "scripts" / "build_release.py")
release = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(release)


def test_release_contains_only_allowlisted_reusable_files(tmp_path):
    root = tmp_path / "source"
    root.mkdir()
    for name in release.FILES:
        target = root / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("1.1.0" if name == "VERSION" else "synthetic", encoding="utf-8")
    for name in ("manifest.json", "ticket-private.json", "screenshot.png", "customer.xlsx",
                 ".env", ".git/config", "runs/private.json", "tests/__pycache__/private.pyc"):
        target = root / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("PRIVATE SYNTHETIC CONTENT", encoding="utf-8")
    output = tmp_path / "release.zip"
    result = release.build(output, root)
    assert result["version"] == "1.1.0"
    with ZipFile(output) as archive:
        assert set(archive.namelist()) == {
            "foundry-quota-increase/" + name for name in release.FILES
        }
        assert not any(b"PRIVATE SYNTHETIC CONTENT" in archive.read(name)
                       for name in archive.namelist())
    with pytest.raises(FileExistsError):
        release.build(output, root)


def test_release_refuses_missing_files(tmp_path):
    with pytest.raises(ValueError, match="Missing or unsafe"):
        release.build(tmp_path / "release.zip", tmp_path)
    assert not (tmp_path / "release.zip").exists()


def test_release_includes_template_docs_license_and_matching_version(tmp_path):
    output = tmp_path / "skill.zip"
    release.build(output)
    with ZipFile(output) as archive:
        for name in ("README.md", "LICENSE"):
            assert archive.read("foundry-quota-increase/" + name) == (
                ROOT / name
            ).read_bytes()
        assert archive.read("foundry-quota-increase/assets/quota_increase_request_template.xlsx") == (
            ROOT / "assets" / "quota_increase_request_template.xlsx"
        ).read_bytes()
        version = archive.read("foundry-quota-increase/VERSION").decode().strip()
        assert f'version: "{version}"' in archive.read(
            "foundry-quota-increase/SKILL.md"
        ).decode("utf-8")

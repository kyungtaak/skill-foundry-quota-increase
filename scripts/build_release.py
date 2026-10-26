"""Package reusable skill files only; never include local request runs or captures."""

import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[1]
FILES = (
    ".gitignore", ".github/workflows/tests.yml", "README.md", "LICENSE",
    "SKILL.md", "VERSION",
    "requirements.txt", "requirements-dev.txt",
    "assets/quota_increase_request_template.xlsx",
    "references/form-map.json", "references/interactive-intake.json",
    "scripts/browser_flow.js", "scripts/quota_workbook.py",
    "scripts/make_blank_template.py", "scripts/build_release.py",
    "tests/test_quota_workbook.py", "tests/test_template.py", "tests/test_release.py",
    "tests/make_browser_smoke.py", "tests/fixtures/form.html",
)


def build(output, root=ROOT):
    root = Path(root).resolve()
    output = Path(output).resolve()
    if output.suffix.lower() != ".zip":
        raise ValueError("The release output must be a ZIP file.")
    if output.exists():
        raise FileExistsError(output)
    for name in FILES:
        source = root / name
        if source.is_symlink() or not source.is_file() or not source.resolve().is_relative_to(root):
            raise ValueError("Missing or unsafe release file: " + name)
    with ZipFile(output, "x", ZIP_DEFLATED) as archive:
        for name in FILES:
            archive.write(root / name, "foundry-quota-increase/" + name)
    return {"output": str(output), "version": (root / "VERSION").read_text().strip(),
            "files": len(FILES)}


if __name__ == "__main__":
    import json
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    print(json.dumps(build(args.output)))

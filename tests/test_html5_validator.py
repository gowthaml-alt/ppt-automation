from pathlib import Path

import pytest

from utils.exceptions import OutputValidationError
from validator.html5 import validate_ispring_output


def _package(root: Path) -> Path:
    (root / "data").mkdir(parents=True)
    (root / "index.html").write_text("<html>ok</html>", encoding="utf-8")
    (root / "data" / "app.js").write_text("console.log(1)", encoding="utf-8")
    return root


def test_valid_package_returns_manifest(tmp_path):
    manifest = validate_ispring_output(_package(tmp_path / "output"))
    names = {row["path"] for row in manifest}
    assert "index.html" in names
    assert "data/app.js" in names


def test_missing_index_is_error(tmp_path):
    output = tmp_path / "output"
    output.mkdir()
    (output / "data").mkdir()
    (output / "data" / "app.js").write_text("x", encoding="utf-8")
    with pytest.raises(OutputValidationError):
        validate_ispring_output(output)


def test_no_asset_subdirectory_is_error(tmp_path):
    output = tmp_path / "output"
    output.mkdir()
    (output / "index.html").write_text("<html>stub</html>", encoding="utf-8")
    with pytest.raises(OutputValidationError):
        validate_ispring_output(output)

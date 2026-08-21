import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from ppt_mcp.config import Settings  # noqa: E402


@pytest.fixture
def settings(tmp_path: Path) -> Settings:
    return Settings(output_dir=(tmp_path / "out").resolve(), template_dir=tmp_path)


@pytest.fixture
def tiny_png(tmp_path: Path) -> Path:
    import base64
    path = tmp_path / "tiny.png"
    path.write_bytes(base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="))
    return path

"""PPT(.pptx) 생성·편집 MCP 서버."""

from .config import Settings
from .server import build_server

__version__ = "0.1.0"
__all__ = ["build_server", "Settings", "__version__"]

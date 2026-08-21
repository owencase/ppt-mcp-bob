"""이미지 입력(경로/base64/URL)을 바이트 스트림으로 바꾼다.

URL 다운로드는 기본으로 꺼 둔다. 도구 인자는 LLM이 채우기 때문에, 프롬프트
인젝션으로 임의 주소를 호출하는 통로가 되지 않게 명시적 opt-in을 요구한다.
"""

from __future__ import annotations

import base64
import binascii
from io import BytesIO
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from .config import IMAGE_SUFFIXES, Settings, resolve_read_path
from .models import ImageSource

MAX_IMAGE_BYTES = 20 * 1024 * 1024
_FETCH_TIMEOUT = 15


def load_image(source: ImageSource, settings: Settings) -> BytesIO:
    if source.path:
        path = resolve_read_path(source.path, settings, suffixes=IMAGE_SUFFIXES)
        data = path.read_bytes()
    elif source.base64:
        raw = source.base64.strip()
        if raw.startswith("data:"):  # data URI 접두사 제거
            _, _, raw = raw.partition(",")
        try:
            data = base64.b64decode(raw, validate=True)
        except (binascii.Error, ValueError) as exc:
            raise ValueError(f"base64 이미지를 해석할 수 없습니다: {exc}") from exc
    else:
        data = _fetch(source.url or "", settings)

    if not data:
        raise ValueError("이미지 데이터가 비어 있습니다.")
    if len(data) > MAX_IMAGE_BYTES:
        raise ValueError(f"이미지가 너무 큽니다({len(data) / 1e6:.1f}MB). 최대 {MAX_IMAGE_BYTES / 1e6:.0f}MB.")
    return BytesIO(data)


def _fetch(url: str, settings: Settings) -> bytes:
    if not settings.allow_remote_images:
        raise PermissionError(
            "URL 이미지는 기본적으로 차단돼 있습니다. 파일을 내려받아 path로 넘기거나, "
            "서버에 PPT_MCP_ALLOW_REMOTE_IMAGES=1을 설정하세요."
        )
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"}:
        raise ValueError(f"http/https URL만 허용합니다: {url!r}")
    request = Request(url, headers={"User-Agent": "ppt-mcp/0.1"})
    with urlopen(request, timeout=_FETCH_TIMEOUT) as response:  # noqa: S310 - 스킴 검증 완료
        content_type = (response.headers.get("Content-Type") or "").split(";")[0].strip()
        if content_type and not content_type.startswith("image/"):
            raise ValueError(f"이미지가 아닌 응답입니다(Content-Type: {content_type}).")
        return response.read(MAX_IMAGE_BYTES + 1)


def guess_extension(name: str) -> str:
    suffix = ("." + name.rsplit(".", 1)[-1]).lower() if "." in name else ""
    return suffix if suffix in IMAGE_SUFFIXES else ".png"


__all__ = ["load_image", "MAX_IMAGE_BYTES", "guess_extension"]

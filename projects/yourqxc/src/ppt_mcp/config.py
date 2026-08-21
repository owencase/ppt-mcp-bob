"""환경변수 기반 설정과 경로 안전장치.

쓰기는 기본적으로 작업 디렉터리(output_dir) 안으로 제한한다. MCP 도구 인자는
LLM이 채우기 때문에, 프롬프트 인젝션으로 임의 경로에 파일을 쓰는 일을 막는다.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

#: 읽기를 허용할 확장자. 이 목록 밖의 파일은 도구가 열지 않는다.
DECK_SUFFIXES = frozenset({".pptx", ".potx", ".pptm"})
IMAGE_SUFFIXES = frozenset({".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".webp", ".emf", ".wmf"})


class PathNotAllowed(ValueError):
    """설정된 작업 영역 밖의 경로를 요청했을 때."""


def _env_bool(name: str, default: bool = False) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _env_path(name: str) -> Path | None:
    raw = os.environ.get(name)
    if not raw or not raw.strip():
        return None
    return Path(raw).expanduser()


@dataclass(frozen=True)
class Settings:
    """서버 실행 설정."""

    output_dir: Path
    template_dir: Path | None = None
    default_template: Path | None = None
    default_theme: str = "carbon_light"
    allow_any_path: bool = False
    allow_remote_images: bool = False
    max_slides: int = 200

    @classmethod
    def from_env(cls) -> "Settings":
        output_dir = _env_path("PPT_MCP_OUTPUT_DIR") or (Path.cwd() / "output")
        return cls(
            output_dir=output_dir.resolve(),
            template_dir=_env_path("PPT_MCP_TEMPLATE_DIR"),
            default_template=_env_path("PPT_MCP_DEFAULT_TEMPLATE"),
            default_theme=os.environ.get("PPT_MCP_DEFAULT_THEME", "carbon_light").strip() or "carbon_light",
            allow_any_path=_env_bool("PPT_MCP_ALLOW_ANY_PATH"),
            allow_remote_images=_env_bool("PPT_MCP_ALLOW_REMOTE_IMAGES"),
            max_slides=int(os.environ.get("PPT_MCP_MAX_SLIDES", "200")),
        )

    def ensure_output_dir(self) -> Path:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        return self.output_dir


def _is_within(child: Path, parent: Path) -> bool:
    try:
        child.relative_to(parent)
    except ValueError:
        return False
    return True


def resolve_write_path(raw: str | None, settings: Settings, *, default_stem: str = "deck") -> Path:
    """저장 경로를 작업 디렉터리 기준으로 해석한다.

    상대 경로는 output_dir 기준으로 붙이고, 절대 경로는 output_dir 안일 때만 허용한다
    (`PPT_MCP_ALLOW_ANY_PATH=1`이면 제한 해제).
    """
    settings.ensure_output_dir()
    if not raw or not raw.strip():
        path = settings.output_dir / f"{default_stem}.pptx"
    else:
        candidate = Path(raw).expanduser()
        path = candidate if candidate.is_absolute() else settings.output_dir / candidate
    path = path if path.suffix else path.with_suffix(".pptx")
    path = path.resolve()

    if path.suffix.lower() not in DECK_SUFFIXES:
        raise PathNotAllowed(f"저장 확장자는 {sorted(DECK_SUFFIXES)} 중 하나여야 합니다: {path.name}")
    if not settings.allow_any_path and not _is_within(path, settings.output_dir):
        raise PathNotAllowed(
            f"'{path}'는 작업 디렉터리 밖입니다. 작업 디렉터리: {settings.output_dir} "
            "(제한을 풀려면 PPT_MCP_ALLOW_ANY_PATH=1)"
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def resolve_read_path(raw: str, settings: Settings, *, suffixes: frozenset[str]) -> Path:
    """읽기 경로 해석. 상대 경로는 output_dir → template_dir → cwd 순으로 찾는다."""
    candidate = Path(raw).expanduser()
    if candidate.is_absolute():
        path = candidate.resolve()
    else:
        roots = [settings.output_dir, settings.template_dir, Path.cwd()]
        path = next(
            (r / candidate for r in roots if r and (r / candidate).exists()),
            settings.output_dir / candidate,
        ).resolve()

    if path.suffix.lower() not in suffixes:
        raise PathNotAllowed(f"허용되지 않는 확장자입니다: {path.name} (허용: {sorted(suffixes)})")
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")
    return path

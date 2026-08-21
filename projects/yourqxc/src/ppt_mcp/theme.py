"""테마(색·폰트·치수)와 슬라이드 기본 배치 규격.

여기 담긴 팔레트는 가독성 위주의 중립 팔레트다. 공식 사내 브랜드 규정을 따라야
하면 테마 대신 실제 템플릿 파일(.potx/.pptx)을 `template` 인자로 넘기면 된다.
"""

from __future__ import annotations

from dataclasses import dataclass, replace

from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


def hex_to_rgb(value: str) -> RGBColor:
    """'#0F62FE' 또는 '0F62FE' → RGBColor."""
    cleaned = value.strip().lstrip("#")
    if len(cleaned) == 3:  # #abc 축약형
        cleaned = "".join(ch * 2 for ch in cleaned)
    if len(cleaned) != 6:
        raise ValueError(f"색상은 6자리 hex여야 합니다: {value!r}")
    return RGBColor.from_string(cleaned.upper())


@dataclass(frozen=True)
class Geometry:
    """16:9 슬라이드의 안전 영역. 단위는 인치."""

    slide_w: float = 13.333
    slide_h: float = 7.5
    margin_x: float = 0.9
    title_top: float = 0.55
    content_top: float = 1.72
    margin_bottom: float = 0.72
    gutter: float = 0.4

    @property
    def content_w(self) -> float:
        return self.slide_w - 2 * self.margin_x

    @property
    def content_h(self) -> float:
        return self.slide_h - self.content_top - self.margin_bottom

    @property
    def column_w(self) -> float:
        return (self.content_w - self.gutter) / 2

    def emu(self, inches: float):
        return Inches(inches)


@dataclass(frozen=True)
class Theme:
    """덱 전체에 적용되는 시각 토큰."""

    name: str
    background: str
    surface: str
    text: str
    text_muted: str
    accent: str
    accent_alt: str
    positive: str
    negative: str
    divider: str
    font_latin: str = "Arial"
    font_ea: str = "맑은 고딕"

    # 포인트 단위 타이포 스케일
    size_cover_title: int = 44
    size_cover_subtitle: int = 20
    size_section_title: int = 36
    size_slide_title: int = 30
    size_heading: int = 20
    size_body: int = 18
    size_small: int = 14
    size_caption: int = 12
    size_kpi_value: int = 54
    size_kpi_label: int = 14

    geometry: Geometry = Geometry()

    def with_fonts(self, latin: str | None, ea: str | None) -> "Theme":
        return replace(self, font_latin=latin or self.font_latin, font_ea=ea or self.font_ea)

    def with_overrides(self, overrides: dict[str, str] | None) -> "Theme":
        """{'accent': '#FF0000'} 형태로 색상 일부만 교체."""
        if not overrides:
            return self
        allowed = {"background", "surface", "text", "text_muted", "accent",
                   "accent_alt", "positive", "negative", "divider"}
        bad = set(overrides) - allowed
        if bad:
            raise ValueError(f"덮어쓸 수 없는 색상 키: {sorted(bad)} (가능: {sorted(allowed)})")
        for key, value in overrides.items():
            hex_to_rgb(value)  # 형식 검증
        return replace(self, **overrides)

    # 자주 쓰는 변환 --------------------------------------------------
    @property
    def rgb_text(self) -> RGBColor:
        return hex_to_rgb(self.text)

    @property
    def rgb_muted(self) -> RGBColor:
        return hex_to_rgb(self.text_muted)

    @property
    def rgb_accent(self) -> RGBColor:
        return hex_to_rgb(self.accent)

    @property
    def is_dark(self) -> bool:
        r, g, b = (int(self.background.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
        return (0.299 * r + 0.587 * g + 0.114 * b) < 128


_CARBON_LIGHT = Theme(
    name="carbon_light",
    background="#FFFFFF", surface="#F4F4F4", text="#161616", text_muted="#525252",
    accent="#0F62FE", accent_alt="#8A3FFC", positive="#198038", negative="#DA1E28",
    divider="#E0E0E0", font_latin="IBM Plex Sans", font_ea="IBM Plex Sans KR",
)

_THEMES: dict[str, Theme] = {
    "carbon_light": _CARBON_LIGHT,
    "carbon_dark": Theme(
        name="carbon_dark",
        background="#161616", surface="#262626", text="#F4F4F4", text_muted="#A8A8A8",
        accent="#4589FF", accent_alt="#BE95FF", positive="#42BE65", negative="#FA4D56",
        divider="#393939", font_latin="IBM Plex Sans", font_ea="IBM Plex Sans KR",
    ),
    "minimal": Theme(
        name="minimal",
        background="#FFFFFF", surface="#F6F6F4", text="#1A1A1A", text_muted="#6B6B6B",
        accent="#1A1A1A", accent_alt="#8C8C8C", positive="#2E7D32", negative="#C62828",
        divider="#E3E3E0", font_latin="Arial", font_ea="맑은 고딕",
    ),
    "vivid": Theme(
        name="vivid",
        background="#FFFFFF", surface="#F3F0FF", text="#14142B", text_muted="#5A5A72",
        accent="#5B21B6", accent_alt="#DB2777", positive="#059669", negative="#DC2626",
        divider="#E5E0F5", font_latin="Arial", font_ea="맑은 고딕",
    ),
}

DEFAULT_THEME = "carbon_light"


def get_theme(name: str | None) -> Theme:
    key = (name or DEFAULT_THEME).strip().lower()
    if key not in _THEMES:
        raise ValueError(f"알 수 없는 테마 '{name}'. 사용 가능: {sorted(_THEMES)}")
    return _THEMES[key]


def list_theme_names() -> list[str]:
    return sorted(_THEMES)


def describe_themes() -> list[dict[str, str]]:
    return [
        {
            "name": t.name,
            "background": t.background,
            "accent": t.accent,
            "font_latin": t.font_latin,
            "font_ea": t.font_ea,
            "mode": "dark" if t.is_dark else "light",
        }
        for t in (_THEMES[k] for k in sorted(_THEMES))
    ]


__all__ = ["Theme", "Geometry", "get_theme", "list_theme_names", "describe_themes",
           "hex_to_rgb", "DEFAULT_THEME", "Pt", "Inches"]

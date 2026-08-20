"""테마 팔레트.

색은 레포 규약대로 '#' 없는 RRGGBB 문자열입니다.
레퍼런스 구현(packages/ppt-bridge)과 같은 이름·같은 값을 씁니다. 학생끼리
결과물을 비교할 때 테마가 달라서 생기는 잡음을 없애려는 것입니다.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Theme:
    name: str
    bg: str       # 슬라이드 배경
    text: str     # 본문 텍스트
    accent: str   # 제목·강조
    muted: str    # 캡션·각주처럼 덜 중요한 텍스트

    def as_dict(self) -> dict[str, str]:
        return {"name": self.name, "bg": self.bg, "text": self.text,
                "accent": self.accent, "muted": self.muted}


THEMES: dict[str, Theme] = {
    "minimal_dark":   Theme("minimal_dark",   "1E1E2E", "FFFFFF", "CBA6F7", "9399B2"),
    "minimal_light":  Theme("minimal_light",  "FFFFFF", "1E1E2E", "3B82F6", "6B7280"),
    "tech_blue":      Theme("tech_blue",      "0F172A", "E2E8F0", "38BDF8", "94A3B8"),
    "marketing_warm": Theme("marketing_warm", "FFF7ED", "1C1917", "F97316", "78716C"),
}

DEFAULT_THEME = "minimal_light"


def get_theme(name: str) -> Theme:
    """테마를 찾습니다. 없으면 고를 수 있는 목록을 담아 실패합니다.

    모델에게 'unknown theme' 만 돌려주면 다음에 또 틀립니다. 무엇을 쓸 수
    있는지 같이 줘야 스스로 고칠 수 있습니다.
    """
    try:
        return THEMES[name]
    except KeyError:
        raise ValueError(
            f"'{name}' 테마는 없습니다. 사용 가능: {', '.join(sorted(THEMES))}"
        ) from None

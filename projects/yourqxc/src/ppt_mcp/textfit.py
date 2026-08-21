"""텍스트 넘침 방지용 글자 크기 추정.

python-pptx에는 '넘치면 줄이기'가 없어서, 실제 렌더링 없이 폭·높이를 어림잡아
들어갈 만한 크기를 고른다. 정확한 조판이 아니라 넘침을 막는 게 목적이다.
"""

from __future__ import annotations

import math

# 전각으로 취급할 유니코드 구간(한글/한자/가나/전각 기호).
_WIDE_RANGES = (
    (0x1100, 0x115F), (0x2E80, 0xA4CF), (0xAC00, 0xD7A3),
    (0xF900, 0xFAFF), (0xFE30, 0xFE6F), (0xFF00, 0xFF60), (0xFFE0, 0xFFE6),
)


def is_wide(ch: str) -> bool:
    code = ord(ch)
    return any(lo <= code <= hi for lo, hi in _WIDE_RANGES)


def display_width(text: str) -> float:
    """한글·한자는 2, 그 외는 1로 센 문자열 폭."""
    return sum(2.0 if is_wide(ch) else 1.0 for ch in text)


def has_cjk(text: str) -> bool:
    return any(is_wide(ch) for ch in text)


def estimate_line_count(text: str, box_w_in: float, size_pt: float, indent_in: float = 0.0) -> int:
    """주어진 폭에서 몇 줄로 접힐지 어림."""
    usable_pt = max(0.4, box_w_in - indent_in) * 72.0
    # 반각 한 글자의 평균 진행폭 ≈ 글자 크기의 0.52배 (Plex/Arial 계열 실측 근사)
    chars_per_line = max(4.0, usable_pt / (size_pt * 0.52))
    explicit = text.split("\n") or [""]
    return sum(max(1, math.ceil(display_width(part) / chars_per_line)) for part in explicit)


def estimate_block_height_pt(
    lines: list[tuple[str, int]],
    box_w_in: float,
    size_pt: float,
    *,
    line_spacing: float = 1.18,
    para_gap_pt: float = 6.0,
    indent_per_level_in: float = 0.34,
) -> float:
    """(텍스트, 들여쓰기 단계) 목록이 차지할 높이를 포인트로 어림."""
    total = 0.0
    for text, level in lines:
        count = estimate_line_count(text, box_w_in, size_pt, indent_per_level_in * level)
        total += count * size_pt * line_spacing + para_gap_pt
    return total


def fit_font_size(
    lines: list[tuple[str, int]],
    box_w_in: float,
    box_h_in: float,
    *,
    base_pt: int,
    min_pt: int,
    line_spacing: float = 1.18,
    para_gap_pt: float = 6.0,
    indent_per_level_in: float = 0.34,
) -> int:
    """(텍스트, 들여쓰기 단계) 목록이 상자에 들어가는 최대 크기를 고른다.

    base_pt에서 시작해 1pt씩 줄이며, 예상 높이가 상자 높이 이하가 되면 멈춘다.
    끝까지 못 맞추면 min_pt를 돌려준다. 그래도 넘치는 경우는 호출부에서
    PowerPoint 자동 축소 비율(normAutofit)을 따로 계산해 넣는다.
    """
    if not lines:
        return base_pt
    available_pt = box_h_in * 72.0
    for size in range(base_pt, min_pt - 1, -1):
        if estimate_block_height_pt(lines, box_w_in, size, line_spacing=line_spacing,
                                    para_gap_pt=para_gap_pt,
                                    indent_per_level_in=indent_per_level_in) <= available_pt:
            return size
    return min_pt


def fit_single_line(text: str, box_w_in: float, *, base_pt: int, min_pt: int) -> int:
    """제목처럼 한 줄로 두고 싶은 텍스트의 크기."""
    usable_pt = box_w_in * 72.0
    for size in range(base_pt, min_pt - 1, -1):
        if display_width(text) * size * 0.52 <= usable_pt:
            return size
    return min_pt


__all__ = ["display_width", "has_cjk", "is_wide", "estimate_line_count",
           "estimate_block_height_pt", "fit_font_size", "fit_single_line"]

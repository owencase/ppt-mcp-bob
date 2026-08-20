"""레이아웃 계산.

이 모듈이 이 MCP 서버의 핵심입니다.

레퍼런스 구현은 `add_text_box(left_cm, top_cm, width_cm, height_cm, ...)` 처럼
좌표를 호출자에게 맡깁니다. 그러면 LLM 이 좌표를 직접 계산하게 되고, 계산이
틀리면 글자가 슬라이드 밖으로 나가거나 상자를 넘칩니다. 실제로 자주 그럽니다.

여기서는 **서버가 레이아웃을 책임집니다.** 호출자는 "제목 슬라이드", "불릿
슬라이드" 같은 의도만 말하고, 좌표와 폰트 크기는 이 모듈이 정합니다.

python-pptx 를 import 하지 않는 순수 함수만 둡니다. 파일을 만들지 않고도
테스트할 수 있어야, 레이아웃 규칙을 빠르고 촘촘하게 검증할 수 있습니다.
"""
from __future__ import annotations

from dataclasses import dataclass

# 16:9 와이드스크린. 단위는 레포 규약대로 cm 입니다.
SLIDE_W = 33.87
SLIDE_H = 19.05

MARGIN_X = 2.54
MARGIN_TOP = 1.9
MARGIN_BOTTOM = 1.6

CONTENT_W = SLIDE_W - 2 * MARGIN_X

# 폰트 크기 하한. 이보다 작아지면 화면에서 안 읽힙니다. 줄이는 대신 실패시켜서
# 호출자가 글을 줄이도록 만듭니다.
MIN_BODY_PT = 14.0
MIN_TITLE_PT = 20.0


@dataclass(frozen=True)
class Box:
    """cm 단위 사각형. python-pptx 에 그대로 넘길 수 있는 형태."""
    left: float
    top: float
    width: float
    height: float

    @property
    def bottom(self) -> float:
        return self.top + self.height

    @property
    def right(self) -> float:
        return self.left + self.width


def within_slide(box: Box, tolerance: float = 0.01) -> bool:
    """상자가 슬라이드 안에 있는지. 렌더 직전 마지막 방어선입니다."""
    return (box.left >= -tolerance and box.top >= -tolerance
            and box.right <= SLIDE_W + tolerance
            and box.bottom <= SLIDE_H + tolerance)


def estimate_lines(text: str, font_pt: float, width_cm: float) -> int:
    """주어진 폭에서 이 텍스트가 몇 줄이 될지 추정합니다.

    정확한 값은 폰트 메트릭이 있어야 알 수 있지만, 여기서는 넘침을 '막는' 것이
    목적이라 보수적인 근사로 충분합니다. 글자 평균 폭을 font_pt 의 0.6 배로
    잡습니다 (한글은 거의 1.0 이지만, 섞여 있을 때 과소추정하면 넘치므로
    한글 비중을 반영해 올려 잡습니다).
    """
    if not text:
        return 1
    hangul = sum(1 for ch in text if "가" <= ch <= "힣")
    ratio = 0.6 + 0.4 * (hangul / len(text))          # 0.6(라틴) ~ 1.0(한글)
    char_w_cm = font_pt * 0.03528 * ratio             # 1pt = 0.03528cm
    per_line = max(1, int(width_cm / char_w_cm))
    lines = 0
    for paragraph in text.split("\n"):
        lines += max(1, -(-len(paragraph) // per_line))   # 올림 나눗셈
    return lines


def text_height(text: str, font_pt: float, width_cm: float, line_spacing: float = 1.3) -> float:
    """줄 수 × 줄 높이 = 텍스트가 실제로 차지하는 높이(cm)."""
    return estimate_lines(text, font_pt, width_cm) * font_pt * 0.03528 * line_spacing


def fit_font_size(text: str, box: Box, start_pt: float, min_pt: float) -> float:
    """상자에 들어갈 때까지 폰트를 줄입니다. 못 줄이면 실패합니다.

    조용히 잘리게 두지 않는 게 핵심입니다. 잘린 PPT 는 인쇄하고 나서야
    발견되는데, 그때는 이미 늦습니다.
    """
    size = start_pt
    while size >= min_pt:
        if text_height(text, size, box.width) <= box.height:
            return size
        size -= 1.0
    raise ValueError(
        f"{min_pt:.0f}pt 로 줄여도 상자에 안 들어갑니다 "
        f"(상자 {box.width:.1f}×{box.height:.1f}cm, 글자 {len(text)}자). "
        f"텍스트를 줄이거나 슬라이드를 나누세요."
    )


# ── 슬라이드 종류별 레이아웃 ────────────────────────────────────────────────
# 각 함수는 Box 들을 돌려줍니다. 좌표 계산은 전부 여기 모여 있어서,
# 디자인을 바꾸고 싶으면 이 아래만 고치면 됩니다.

def title_slide() -> dict[str, Box]:
    """표지: 제목을 시각적 중앙보다 살짝 위에 둡니다(광학적 중심)."""
    return {
        "title": Box(MARGIN_X, 6.6, CONTENT_W, 4.2),
        "subtitle": Box(MARGIN_X, 11.2, CONTENT_W, 2.2),
    }


def section_slide() -> dict[str, Box]:
    """구분 슬라이드: 큰 글씨 하나만."""
    return {"title": Box(MARGIN_X, 7.6, CONTENT_W, 3.8)}


def bullets_slide(count: int) -> dict[str, Box]:
    """제목 + 불릿. 불릿 개수에 따라 영역 높이가 정해집니다."""
    title = Box(MARGIN_X, MARGIN_TOP, CONTENT_W, 2.6)
    body_top = title.bottom + 0.9
    body_h = SLIDE_H - MARGIN_BOTTOM - body_top
    boxes = {"title": title, "body": Box(MARGIN_X, body_top, CONTENT_W, body_h)}
    if count > 0:
        # 불릿 하나가 쓸 수 있는 높이. 넘침 검사는 fit_font_size 가 합니다.
        boxes["bullet_slot"] = Box(MARGIN_X, body_top, CONTENT_W, body_h / count)
    return boxes


def chart_slide() -> dict[str, Box]:
    """제목 + 차트. 차트는 가로를 꽉 채우지 않고 양옆에 여백을 더 줍니다."""
    title = Box(MARGIN_X, MARGIN_TOP, CONTENT_W, 2.6)
    top = title.bottom + 0.6
    return {
        "title": title,
        "chart": Box(MARGIN_X + 1.2, top, CONTENT_W - 2.4, SLIDE_H - MARGIN_BOTTOM - top),
    }

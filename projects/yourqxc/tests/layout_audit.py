"""생성된 덱의 레이아웃 불변식 검사(테스트 전용).

렌더링 없이 도형 좌표만으로 '슬라이드 밖으로 나감 / 글자 넘침 / 텍스트 겹침'을 잡는다.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from ppt_mcp.textfit import estimate_block_height_pt

EMU = 914400.0
_BOUNDS_SLACK = 0.02
_OVERLAP_SLACK = 0.08


_NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _font_scale(text_frame) -> float:
    """<a:normAutofit fontScale="..."/>에 저장된 축소 비율(0~1). 없으면 1.0."""
    autofit = text_frame._txBody.bodyPr.find(f"{_NS_A}normAutofit")
    if autofit is None:
        return 1.0
    raw = autofit.get("fontScale")
    return int(raw) / 100000.0 if raw else 1.0


def audit(presentation) -> list[str]:
    slide_w = presentation.slide_width / EMU
    slide_h = presentation.slide_height / EMU
    problems: list[str] = []

    for number, slide in enumerate(presentation.slides, start=1):
        text_boxes: list[tuple[float, float, float, float, str]] = []
        for shape in slide.shapes:
            if shape.left is None:
                continue
            x, y = shape.left / EMU, shape.top / EMU
            w, h = (shape.width or 0) / EMU, (shape.height or 0) / EMU

            if (x < -_BOUNDS_SLACK or y < -_BOUNDS_SLACK
                    or x + w > slide_w + _BOUNDS_SLACK or y + h > slide_h + _BOUNDS_SLACK):
                problems.append(
                    f"s{number}: 슬라이드 밖 ({x:.2f},{y:.2f} {w:.2f}x{h:.2f})")

            if not shape.has_text_frame or not shape.text_frame.text.strip():
                continue

            # PowerPoint가 열 때 적용할 자동 축소 비율을 반영해야 실제 렌더 높이가 나온다.
            # 높이 추정은 빌더와 같은 함수를 써서 둘이 어긋나지 않게 한다.
            scale = _font_scale(shape.text_frame)
            needed_pt = 0.0
            for para in shape.text_frame.paragraphs:
                text = "".join(run.text for run in para.runs)
                if not text:
                    continue
                size = next((r.font.size.pt for r in para.runs if r.font.size), 18.0) * scale
                needed_pt += estimate_block_height_pt([(text, 0)], w, size)
            if needed_pt > h * 72 + 4:
                problems.append(
                    f"s{number}: 텍스트 넘침 {shape.text_frame.text[:20]!r} "
                    f"필요 {needed_pt / 72:.2f}in > 상자 {h:.2f}in")
            text_boxes.append((x, y, w, h, shape.text_frame.text[:18]))

        for i in range(len(text_boxes)):
            for j in range(i + 1, len(text_boxes)):
                x1, y1, w1, h1, t1 = text_boxes[i]
                x2, y2, w2, h2, t2 = text_boxes[j]
                overlap_x = min(x1 + w1, x2 + w2) - max(x1, x2)
                overlap_y = min(y1 + h1, y2 + h2) - max(y1, y2)
                if overlap_x > _OVERLAP_SLACK and overlap_y > _OVERLAP_SLACK:
                    problems.append(f"s{number}: 텍스트 겹침 {t1!r} × {t2!r}")

    return problems

"""
ppt-bridge/bridge.py
────────────────────
Python bridge between the MCP Server and Microsoft PowerPoint.

Protocol
--------
stdin  → single JSON line: {"action": "<name>", "params": {...}}
stdout → single JSON line: {"success": true/false, "message": "...", "data": {...}}

All measurements that arrive from the MCP server are in centimetres.
python-pptx works in EMUs (1 cm = 360_000 EMU), so we convert internally.

Dependencies
------------
  pip install -r requirements.txt
"""

from __future__ import annotations

import io
import json
import sys
from typing import Any

# Windows에서 stdout/stdin을 UTF-8로 강제 설정
sys.stdin  = io.TextIOWrapper(sys.stdin.buffer,  encoding="utf-8", errors="replace")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# ── python-pptx imports ────────────────────────────────────────────────────
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
CM_TO_EMU = 360_000

# Predefined themes  {name: {bg, text, accent1, accent2}}
THEMES: dict[str, dict[str, str]] = {
    "minimal_dark":   {"bg": "1E1E2E", "text": "FFFFFF", "accent1": "CBA6F7", "accent2": "89DCEB"},
    "minimal_light":  {"bg": "FFFFFF", "text": "1E1E2E", "accent1": "3B82F6", "accent2": "8B5CF6"},
    "tech_blue":      {"bg": "0F172A", "text": "E2E8F0", "accent1": "38BDF8", "accent2": "22D3EE"},
    "marketing_warm": {"bg": "FFF7ED", "text": "1C1917", "accent1": "F97316", "accent2": "EAB308"},
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def cm(value: float) -> int:
    """Convert centimetres to EMU."""
    return int(value * CM_TO_EMU)


def hex_to_rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def ok(message: str = "OK", data: Any = None) -> dict:
    result: dict[str, Any] = {"success": True, "message": message}
    if data is not None:
        result["data"] = data
    return result


def err(message: str) -> dict:
    return {"success": False, "error": message}


def load_prs(file_path: str) -> Presentation:
    return Presentation(file_path)


def save_prs(prs: Presentation, file_path: str) -> None:
    prs.save(file_path)

# ---------------------------------------------------------------------------
# Action handlers
# ---------------------------------------------------------------------------

def handle_create_presentation(params: dict) -> dict:
    file_path: str = params["file_path"]
    width_cm: float = params.get("width_cm", 33.867)   # 16:9 widescreen
    height_cm: float = params.get("height_cm", 19.05)

    prs = Presentation()
    prs.slide_width = Emu(cm(width_cm))
    prs.slide_height = Emu(cm(height_cm))
    save_prs(prs, file_path)
    return ok(f"Presentation created: {file_path}")


def handle_add_slide(params: dict) -> dict:
    file_path: str = params["file_path"]
    layout_index: int = int(params.get("layout_index", 6))

    prs = load_prs(file_path)
    layout = prs.slide_layouts[layout_index]
    prs.slides.add_slide(layout)
    save_prs(prs, file_path)
    slide_count = len(prs.slides)
    return ok(f"Slide added. Total slides: {slide_count}", {"slide_index": slide_count - 1})


def handle_add_text_box(params: dict) -> dict:
    file_path: str = params["file_path"]
    slide_index: int = int(params["slide_index"])
    text: str = params["text"]
    font_size_pt: float = float(params.get("font_size_pt", 24))
    bold: bool = bool(params.get("bold", False))
    color_hex: str = str(params.get("color_hex", "000000"))
    align_str: str = str(params.get("align", "left"))

    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    alignment = align_map.get(align_str, PP_ALIGN.LEFT)

    prs = load_prs(file_path)
    slide = prs.slides[slide_index]

    txBox = slide.shapes.add_textbox(
        cm(float(params["left_cm"])),
        cm(float(params["top_cm"])),
        cm(float(params["width_cm"])),
        cm(float(params["height_cm"])),
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None

    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    run.font.color.rgb = hex_to_rgb(color_hex)

    save_prs(prs, file_path)
    return ok("Text box added")


def handle_add_image(params: dict) -> dict:
    file_path: str = params["file_path"]
    slide_index: int = int(params["slide_index"])
    image_path: str = params["image_path"]

    prs = load_prs(file_path)
    slide = prs.slides[slide_index]
    slide.shapes.add_picture(
        image_path,
        cm(float(params["left_cm"])),
        cm(float(params["top_cm"])),
        cm(float(params["width_cm"])),
        cm(float(params["height_cm"])),
    )
    save_prs(prs, file_path)
    return ok("Image added")


def handle_set_background_color(params: dict) -> dict:
    file_path: str = params["file_path"]
    slide_index: int = int(params["slide_index"])
    color_hex: str = str(params["color_hex"])

    prs = load_prs(file_path)
    slide = prs.slides[slide_index]

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(color_hex)

    save_prs(prs, file_path)
    return ok(f"Background set to #{color_hex}")


def handle_add_shape(params: dict) -> dict:
    file_path: str = params["file_path"]
    slide_index: int = int(params["slide_index"])
    shape_type: str = str(params.get("shape_type", "rectangle"))
    fill_hex: str = str(params.get("fill_color_hex", "4472C4"))
    line_hex: str | None = params.get("line_color_hex")

    prs = load_prs(file_path)
    slide = prs.slides[slide_index]

    shape_id = (
        MSO_SHAPE.ROUNDED_RECTANGLE
        if shape_type == "rounded_rectangle"
        else MSO_SHAPE.RECTANGLE
    )

    shape = slide.shapes.add_shape(
        shape_id,
        cm(float(params["left_cm"])),
        cm(float(params["top_cm"])),
        cm(float(params["width_cm"])),
        cm(float(params["height_cm"])),
    )

    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)

    if line_hex:
        shape.line.color.rgb = hex_to_rgb(line_hex)
    else:
        shape.line.fill.background()  # no border

    save_prs(prs, file_path)
    return ok(f"{shape_type} shape added")


def handle_apply_theme(params: dict) -> dict:
    file_path: str = params["file_path"]
    theme_name: str = params["theme"]

    if theme_name not in THEMES:
        return err(f"Unknown theme '{theme_name}'. Available: {list(THEMES.keys())}")

    theme = THEMES[theme_name]
    prs = load_prs(file_path)

    for slide in prs.slides:
        # background
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(theme["bg"])

    save_prs(prs, file_path)
    return ok(f"Theme '{theme_name}' applied to all {len(prs.slides)} slides", {"theme": theme})


def handle_save_presentation(params: dict) -> dict:
    file_path: str = params["file_path"]
    prs = load_prs(file_path)
    save_prs(prs, file_path)
    return ok(f"Saved: {file_path}")


def handle_get_presentation_info(params: dict) -> dict:
    file_path: str = params["file_path"]
    prs = load_prs(file_path)

    slides_info = []
    for i, slide in enumerate(prs.slides):
        shapes_info = [
            {"name": s.name, "shape_type": str(s.shape_type), "left_cm": round(s.left / CM_TO_EMU, 2),
             "top_cm": round(s.top / CM_TO_EMU, 2)}
            for s in slide.shapes
        ]
        slides_info.append({"slide_index": i, "shapes": shapes_info})

    data = {
        "slide_count": len(prs.slides),
        "width_cm": round(prs.slide_width / CM_TO_EMU, 2),
        "height_cm": round(prs.slide_height / CM_TO_EMU, 2),
        "slides": slides_info,
    }
    return ok("Presentation info retrieved", data)

# ---------------------------------------------------------------------------
# Dispatch table
# ---------------------------------------------------------------------------
HANDLERS = {
    "create_presentation":    handle_create_presentation,
    "add_slide":              handle_add_slide,
    "add_text_box":           handle_add_text_box,
    "add_image":              handle_add_image,
    "set_background_color":   handle_set_background_color,
    "add_shape":              handle_add_shape,
    "apply_theme":            handle_apply_theme,
    "save_presentation":      handle_save_presentation,
    "get_presentation_info":  handle_get_presentation_info,
}

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    raw = sys.stdin.read().strip()
    try:
        request = json.loads(raw)
    except json.JSONDecodeError as exc:
        print(json.dumps(err(f"Invalid JSON input: {exc}")))
        return

    action: str = request.get("action", "")
    params: dict = request.get("params", {})

    handler = HANDLERS.get(action)
    if handler is None:
        print(json.dumps(err(f"Unknown action '{action}'. Available: {list(HANDLERS.keys())}")))
        return

    try:
        result = handler(params)
    except Exception as exc:  # noqa: BLE001
        result = err(f"Unhandled exception in '{action}': {type(exc).__name__}: {exc}")

    print(json.dumps(result))


if __name__ == "__main__":
    main()

from __future__ import annotations

import json
import os
import platform
import re
import shutil
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from .planner import create_plan
from .routing import resolve_template
from .semantic_qa import inspect_grounding

EMU = 914400
RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
PROTECTED_PLACEHOLDERS = {
    PP_PLACEHOLDER.DATE,
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.SLIDE_NUMBER,
}
PROTECTED_NAME_TOKENS = (
    "footer", "slide number", "slidenumber", "date", "copyright", "logo", "watermark",
)


@dataclass
class TextSlot:
    shape: Any
    name: str
    left: float
    top: float
    width: float
    height: float
    font_size: float
    placeholder_type: int | None
    original_text: str


class ComUnavailableError(RuntimeError):
    pass


def _ensure_com_environment() -> None:
    if platform.system() != "Windows":
        raise ComUnavailableError("COM template editing requires Windows and Microsoft PowerPoint")
    try:
        import pythoncom  # noqa: F401
        import pywintypes  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError as exc:
        raise ComUnavailableError("Install the Windows extra: pip install -e '.[windows]'") from exc


def _hex_color(color) -> str | None:
    try:
        if color.type and color.rgb:
            return f"#{color.rgb}"
    except (AttributeError, TypeError, ValueError):
        return None
    return None


def _walk_shapes_pptx(shapes, prefix: str = ""):
    for index, shape in enumerate(shapes, 1):
        path = f"{prefix}/{index}:{shape.name}" if prefix else f"{index}:{shape.name}"
        yield path, shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk_shapes_pptx(shape.shapes, path)


def _first_run_style(shape) -> tuple:
    if not getattr(shape, "has_text_frame", False):
        return (None, None, None, None, None)
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            return (
                run.font.name,
                round(run.font.size.pt, 2) if run.font.size else None,
                run.font.bold,
                run.font.italic,
                _hex_color(run.font.color),
            )
    return (None, None, None, None, None)


def _shape_signature(path: str, shape) -> tuple:
    placeholder = None
    if getattr(shape, "is_placeholder", False):
        try:
            placeholder = int(shape.placeholder_format.type)
        except Exception:
            placeholder = None
    fill = None
    line = None
    try:
        fill = _hex_color(shape.fill.fore_color)
    except Exception:
        pass
    try:
        line = _hex_color(shape.line.color)
    except Exception:
        pass
    return (
        path,
        int(shape.shape_type),
        int(shape.left), int(shape.top), int(shape.width), int(shape.height),
        round(float(shape.rotation or 0), 2),
        placeholder,
        fill,
        line,
        _first_run_style(shape),
    )


def design_fingerprint(path: str | Path) -> dict:
    """Semantic design fingerprint: ignores text content but tracks visual structure/style."""
    prs = Presentation(str(path))
    slides = []
    for slide_index, slide in enumerate(prs.slides, 1):
        signatures = [_shape_signature(shape_path, shape) for shape_path, shape in _walk_shapes_pptx(slide.shapes)]
        bg = None
        try:
            bg = _hex_color(slide.background.fill.fore_color)
        except Exception:
            pass
        slides.append({"slide": slide_index, "background": bg, "shapes": signatures})
    return {
        "slide_count": len(prs.slides),
        "slide_size": [int(prs.slide_width), int(prs.slide_height)],
        "slides": slides,
    }


def _design_diff(before: dict, after: dict) -> list[str]:
    issues: list[str] = []
    if before["slide_count"] != after["slide_count"]:
        issues.append(f"slide_count changed: {before['slide_count']} -> {after['slide_count']}")
    if before["slide_size"] != after["slide_size"]:
        issues.append("slide size changed")
    for left, right in zip(before["slides"], after["slides"]):
        if left["background"] != right["background"]:
            issues.append(f"slide {left['slide']}: background changed")
        if len(left["shapes"]) != len(right["shapes"]):
            issues.append(f"slide {left['slide']}: shape count changed")
            continue
        for a, b in zip(left["shapes"], right["shapes"]):
            if a != b:
                issues.append(f"slide {left['slide']}: visual/style signature changed at {a[0]}")
                if len(issues) >= 20:
                    return issues
    return issues


def _validate_template_source(template: Path, output: Path) -> None:
    if template.resolve() == output.resolve():
        raise ValueError("template_path and output_path must be different; the original template is never overwritten")
    if output.suffix.lower() not in {".pptx", ".pptm"}:
        raise ValueError("COM edited output must end in .pptx or .pptm")
    output.parent.mkdir(parents=True, exist_ok=True)


def _com_retry(callable_, *args, attempts: int = 5, delay: float = 1.0, **kwargs):
    try:
        import pywintypes  # type: ignore
    except ImportError as exc:
        raise ComUnavailableError("pywin32 is required for COM template editing") from exc
    last = None
    for attempt in range(attempts):
        try:
            return callable_(*args, **kwargs)
        except pywintypes.com_error as exc:  # type: ignore[attr-defined]
            last = exc
            hresult = getattr(exc, "hresult", None)
            if hresult not in {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER} or attempt == attempts - 1:
                raise
            time.sleep(delay)
    raise last  # pragma: no cover


def _connect_powerpoint(visible: bool):
    _ensure_com_environment()
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise ComUnavailableError("Install the Windows extra: pip install -e '.[windows]'") from exc
    pythoncom.CoInitialize()
    launched = False
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        launched = True
    if visible:
        _com_retry(setattr, app, "Visible", True)
    return pythoncom, app, launched


def _iter_com_shapes(shapes) -> Iterable[Any]:
    # msoGroup = 6
    for i in range(1, int(shapes.Count) + 1):
        shape = shapes.Item(i)
        if int(shape.Type) == 6:
            yield from _iter_com_shapes(shape.GroupItems)
        else:
            yield shape


def _placeholder_type(shape) -> int | None:
    # msoPlaceholder = 14
    if int(shape.Type) != 14:
        return None
    try:
        return int(shape.PlaceholderFormat.Type)
    except Exception:
        return None


def _safe_text_slot(shape) -> TextSlot | None:
    try:
        if int(shape.HasTextFrame) != -1:
            return None
    except Exception:
        return None
    name = str(shape.Name or "")
    lowered = name.lower()
    if any(token in lowered for token in PROTECTED_NAME_TOKENS):
        return None
    placeholder = _placeholder_type(shape)
    # COM numeric constants align with PowerPoint placeholder enum:
    # 13 slide number, 14 header, 15 footer, 16 date.
    if placeholder in {13, 14, 15, 16}:
        return None
    try:
        text = str(shape.TextFrame.TextRange.Text or "").strip()
    except Exception:
        text = ""
    # Preserve tiny numeric/ornamental tokens such as page indexes or "01".
    if text and len(text) <= 5 and re.fullmatch(r"[\d\s./|—–-]+", text):
        return None
    try:
        font_size = float(shape.TextFrame.TextRange.Font.Size)
        if font_size < 0:
            font_size = 0.0
    except Exception:
        font_size = 0.0
    return TextSlot(
        shape=shape,
        name=name,
        left=float(shape.Left), top=float(shape.Top),
        width=float(shape.Width), height=float(shape.Height),
        font_size=font_size,
        placeholder_type=placeholder,
        original_text=text,
    )


def _collect_slots(slide) -> tuple[TextSlot | None, list[TextSlot]]:
    slots = [slot for shape in _iter_com_shapes(slide.Shapes) if (slot := _safe_text_slot(shape))]
    if not slots:
        return None, []
    title = next((slot for slot in slots if slot.placeholder_type in {1, 3, 5}), None)
    if title is None:
        # Favor a large, high text element. Font size dominates; top position breaks ties.
        title = max(slots, key=lambda s: (s.font_size, -s.top, s.width))
    bodies = [slot for slot in slots if slot is not title]
    bodies.sort(key=lambda s: (s.top, s.left))
    return title, bodies


def _content_units(spec) -> list[str]:
    values: list[str] = []
    if spec.subtitle.strip():
        values.append(spec.subtitle.strip())
    for item in spec.items:
        heading = item.heading.strip()
        body = item.body.strip()
        if heading and body:
            values.append(f"{heading}\r{body}")
        elif heading or body:
            values.append(heading or body)
    return values


def _replace_text(slot: TextSlot, text: str) -> None:
    # PowerPoint COM preserves the existing shape, geometry, fill, line and the
    # text range's inherited formatting. TextFrame2 AutoSize=2 shrinks text to
    # fit the existing box instead of resizing the box and breaking the design.
    text_range = slot.shape.TextFrame.TextRange
    _com_retry(setattr, text_range, "Text", text)
    try:
        _com_retry(setattr, slot.shape.TextFrame2, "AutoSize", 2)
    except Exception:
        pass


def _edit_with_com(output: Path, plan, visible: bool) -> list[dict[str, Any]]:
    pythoncom = None
    app = None
    pres = None
    launched = False
    operations: list[dict[str, Any]] = []
    try:
        pythoncom, app, launched = _connect_powerpoint(visible)
        # ykuwai/ppt-mcp inspired safety: hold an explicit presentation object
        # for the selected file and never edit ActivePresentation implicitly.
        pres = _com_retry(app.Presentations.Open, str(output), 0, 0, -1 if visible else 0)
        opened = Path(str(pres.FullName)).resolve()
        if opened != output.resolve():
            raise RuntimeError(f"PowerPoint opened an unexpected target: {opened}")
        if int(pres.Slides.Count) != len(plan.slides):
            raise RuntimeError(
                f"template slide count changed before edit: COM={pres.Slides.Count}, plan={len(plan.slides)}"
            )
        for index, spec in enumerate(plan.slides, 1):
            slide = pres.Slides(index)
            title, bodies = _collect_slots(slide)
            slide_ops = []
            if title is not None:
                _replace_text(title, spec.title)
                slide_ops.append({"shape": title.name, "role": "title", "text": spec.title})
            units = _content_units(spec)
            for body_index, slot in enumerate(bodies):
                replacement = units[body_index] if body_index < len(units) else ""
                _replace_text(slot, replacement)
                slide_ops.append({"shape": slot.name, "role": "body", "text": replacement})
            operations.append({
                "slide": index,
                "title_shape": title.name if title else None,
                "editable_body_slots": len(bodies),
                "content_units": len(units),
                "operations": slide_ops,
            })
        _com_retry(pres.Save)
    finally:
        if pres is not None:
            try:
                _com_retry(pres.Close)
            except Exception:
                pass
        if app is not None and launched:
            try:
                _com_retry(app.Quit)
            except Exception:
                pass
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    return operations


def edit_template_with_com(*, topic: str, template_name: str, output_path: str,
                           audience: str = "", purpose: str = "", language: str = "ko",
                           research_text: str | None = None, source_urls: list[str] | None = None,
                           research_required: bool = True,
                           research_documents: list[dict[str, str]] | None = None,
                           template_dir: str | None = None,
                           visible: bool = False) -> dict[str, Any]:
    """Copy a /template file and replace only text through PowerPoint COM."""
    if not topic or not topic.strip():
        raise ValueError("topic must not be empty")
    _ensure_com_environment()
    template = resolve_template(template_name, template_dir)
    output = Path(output_path).expanduser().resolve()
    _validate_template_source(template, output)

    # COM opens PowerPoint templates differently. Convert POTX/POTM to a working
    # presentation using PowerPoint itself by copying only when extension is already
    # a presentation. For template extensions, PowerPoint can open and SaveAs to pptx.
    before_source = template
    if template.suffix.lower() in {".pptx", ".pptm"}:
        shutil.copy2(template, output)
    else:
        # PowerPoint handles POTX/POTM creation semantics more reliably than python-pptx.
        if platform.system() != "Windows":
            raise ComUnavailableError(".potx/.potm COM template editing requires Windows PowerPoint")
        pythoncom = app = pres = None
        launched = False
        try:
            pythoncom, app, launched = _connect_powerpoint(False)
            pres = _com_retry(app.Presentations.Open, str(template), 0, 1, 0)
            # ppSaveAsOpenXMLPresentation = 24
            _com_retry(pres.SaveAs, str(output), 24)
        finally:
            if pres is not None:
                try: _com_retry(pres.Close)
                except Exception: pass
            if app is not None and launched:
                try: _com_retry(app.Quit)
                except Exception: pass
            if pythoncom is not None:
                try: pythoncom.CoUninitialize()
                except Exception: pass
        before_source = output

    before = design_fingerprint(before_source)
    slide_count = before["slide_count"]
    if not 3 <= slide_count <= 30:
        raise ValueError("template must contain between 3 and 30 slides")
    plan = create_plan(
        topic=topic, audience=audience, purpose=purpose, slide_count=slide_count,
        language=language, content_json=None, research_text=research_text,
        source_urls=source_urls, research_required=research_required,
        style_preference=None, research_documents=research_documents,
    )
    semantic_issues = inspect_grounding(plan, topic, research_required)
    semantic_errors = [issue for issue in semantic_issues if issue.severity == "error"]
    if semantic_errors:
        raise RuntimeError("Semantic QA failed: " + "; ".join(x.message for x in semantic_errors))

    operations = _edit_with_com(output, plan, visible)
    after = design_fingerprint(output)
    design_issues = _design_diff(before, after)
    if design_issues:
        raise RuntimeError(
            "COM edit changed template design; output rejected: " + "; ".join(design_issues[:8])
        )

    manifest_dir = output.with_name(output.stem + "_qa")
    manifest_dir.mkdir(parents=True, exist_ok=True)
    (manifest_dir / "template-com-plan.json").write_text(plan.model_dump_json(indent=2), encoding="utf-8")
    manifest = {
        "mode": "template_com",
        "engine": "PowerPoint COM (pywin32)",
        "template_path": str(template),
        "output_path": str(output),
        "slide_count": slide_count,
        "design_preserved": True,
        "design_issues": [],
        "semantic_qa": {
            "passed": not semantic_errors,
            "issues": [issue.model_dump() for issue in semantic_issues],
        },
        "operations": operations,
        "research_sources": [source.model_dump() for source in plan.research_sources],
        "plan_path": str(manifest_dir / "template-com-plan.json"),
    }
    (manifest_dir / "template-com-manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return manifest

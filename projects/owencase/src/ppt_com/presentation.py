"""Presentation-level operations for PowerPoint COM automation.

Create, open, save, close, and query PowerPoint presentations.
"""

import glob as glob_mod
import json
import logging
import os
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.color import int_to_hex
from utils.com_wrapper import ppt
from utils.tool_result import execute_json
from utils.onedrive import resolve_local_path
from ppt_com.constants import (
    msoTrue,
    msoFalse,
    ppSaveAsOpenXMLPresentation,
    ppSaveAsPDF,
    ppSaveAsPNG,
    ppSaveAsJPG,
    ppSaveAsDefault,
)
from utils.units import (
    SLIDE_WIDTH_16_9,
    SLIDE_HEIGHT_16_9,
    SLIDE_WIDTH_4_3,
    SLIDE_HEIGHT_4_3,
)

logger = logging.getLogger(__name__)

# Mapping of format aliases to PpSaveAsFileType constants
SAVE_FORMAT_MAP = {
    "pptx": ppSaveAsOpenXMLPresentation,
    "pdf": ppSaveAsPDF,
    "png": ppSaveAsPNG,
    "jpg": ppSaveAsJPG,
    "default": ppSaveAsDefault,
}

# Mapping of preset names to (width, height) in points
SLIDE_SIZE_PRESETS = {
    "16:9": (SLIDE_WIDTH_16_9, SLIDE_HEIGHT_16_9),
    "4:3": (SLIDE_WIDTH_4_3, SLIDE_HEIGHT_4_3),
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class CreatePresentationInput(BaseModel):
    """Input for creating a new presentation."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    template_path: Optional[str] = Field(
        default=None,
        description=(
            "Absolute path to a reusable PowerPoint source (.pptx, .pptm, "
            ".potx, .potm, .pot). "
            "When provided, creates a new presentation based on this template, "
            "preserving all slides, layouts, themes, and custom data. "
            "Size parameters (preset, slide_width, slide_height) are ignored "
            "when a template is used."
        ),
    )
    slide_width: Optional[float] = Field(
        default=None,
        description=(
            "Slide width in points (72 points = 1 inch). "
            "Ignored if preset or template_path is provided."
        ),
    )
    slide_height: Optional[float] = Field(
        default=None,
        description=(
            "Slide height in points (72 points = 1 inch). "
            "Ignored if preset or template_path is provided."
        ),
    )
    preset: Optional[str] = Field(
        default=None,
        description=(
            "Slide size preset. Supported values: '16:9' (widescreen, 960x540 pt), "
            "'4:3' (standard, 720x540 pt). Overrides slide_width/slide_height. "
            "Ignored if template_path is provided."
        ),
    )
    activate: bool = Field(
        default=True,
        description=(
            "If true (default), the newly created presentation becomes the "
            "session target — subsequent tool calls operate on it regardless "
            "of which window the user clicks into. Set false to keep an "
            "existing session target."
        ),
    )


class OpenPresentationInput(BaseModel):
    """Input for opening an existing presentation."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    file_path: str = Field(
        ...,
        description="Full path to the presentation file (.pptx, .pptm, .ppt, .potx, etc.).",
    )
    read_only: bool = Field(
        default=False,
        description="If true, open in read-only mode.",
    )
    with_window: bool = Field(
        default=True,
        description="If true, open with a visible window. Set false for background processing.",
    )
    activate: bool = Field(
        default=True,
        description=(
            "If true (default), the opened presentation becomes the session "
            "target — subsequent tool calls operate on it regardless of which "
            "window the user clicks into. Set false to keep an existing "
            "session target."
        ),
    )


class SavePresentationInput(BaseModel):
    """Input for saving the active presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to save. "
            "If omitted, saves the active presentation."
        ),
    )
    presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "Name of the presentation (e.g. 'MySlides.pptx'). "
            "Alternative to presentation_index. If omitted, uses the locked session target."
        ),
    )


class SavePresentationAsInput(BaseModel):
    """Input for SaveAs operation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_path: str = Field(
        ...,
        description="Target file path for saving.",
    )
    format: Optional[str] = Field(
        default=None,
        description=(
            "Output format: 'pptx', 'pdf', 'png', 'jpg', or 'default'. "
            "If omitted, PowerPoint infers from file extension."
        ),
    )
    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to save. "
            "If omitted, saves the active presentation."
        ),
    )
    presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "Name of the presentation (e.g. 'MySlides.pptx'). "
            "Alternative to presentation_index. If omitted, uses the locked session target."
        ),
    )


class ClosePresentationInput(BaseModel):
    """Input for closing a presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    save_changes: bool = Field(
        default=False,
        description="If true, save the presentation before closing. If false, discard unsaved changes.",
    )
    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to close. "
            "If omitted, closes the active presentation."
        ),
    )
    presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "Name of the presentation (e.g. 'MySlides.pptx'). "
            "Alternative to presentation_index. If omitted, uses the locked session target."
        ),
    )


class GetPresentationInfoInput(BaseModel):
    """Input for getting presentation info."""
    model_config = ConfigDict(str_strip_whitespace=True)

    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to query. "
            "If omitted, uses the locked session target."
        ),
    )
    presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "Name of the presentation (e.g. 'MySlides.pptx'). "
            "Alternative to presentation_index. If omitted, uses the locked session target."
        ),
    )


class ActivatePresentationInput(BaseModel):
    """Input for activating a specific presentation as the MCP target."""
    model_config = ConfigDict(str_strip_whitespace=True)

    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to activate. "
            "Use ppt_list_presentations to see available indices."
        ),
    )
    presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "File name of the presentation to activate (e.g. 'demo.pptx'). "
            "Alternative to presentation_index."
        ),
    )


class ListTemplatesInput(BaseModel):
    """Input for listing available PowerPoint templates."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    templates_dir: Optional[str] = Field(
        default=None,
        description=(
            "Directory to scan for .pptx, .pptm, .potx, .potm, and .pot files. "
            "If omitted, PPT_TEMPLATES_DIR must be configured."
        ),
    )


# ---------------------------------------------------------------------------
# Helper to resolve a presentation by explicit identifier or locked target
# ---------------------------------------------------------------------------
def _resolve_presentation(
    app,
    presentation_index: Optional[int] = None,
    presentation_name: Optional[str] = None,
):
    """Return a Presentation COM object by index, name, or locked session target.

    Args:
        app: PowerPoint Application COM object.
        presentation_index: 1-based index of the presentation.
        presentation_name: File name of the presentation (e.g. 'MySlides.pptx').

    Raises:
        ValueError: If both parameters are provided, or if the name matches
            zero or multiple presentations.
        RuntimeError: If no presentations are open.
    """
    if presentation_index is not None and presentation_name is not None:
        raise ValueError(
            "Specify either presentation_index or presentation_name, not both"
        )

    if presentation_index is not None:
        count = app.Presentations.Count
        if presentation_index < 1 or presentation_index > count:
            raise ValueError(
                f"Presentation index {presentation_index} out of range (1-{count})"
            )
        return app.Presentations(presentation_index)

    if presentation_name is not None:
        count = app.Presentations.Count
        if count == 0:
            raise RuntimeError(
                "No presentation is open. "
                "Use ppt_create_presentation or ppt_open_presentation first."
            )
        matches = []
        available = []
        for i in range(1, count + 1):
            pres = app.Presentations(i)
            name = pres.Name
            available.append(f"  [{i}] {name}")
            if name == presentation_name:
                matches.append((i, pres))
        if len(matches) == 1:
            return matches[0][1]
        if len(matches) > 1:
            match_list = ", ".join(
                f"[{idx}] {presentation_name}" for idx, _ in matches
            )
            raise ValueError(
                f"Multiple presentations match name '{presentation_name}': "
                f"{match_list}. Use presentation_index to disambiguate."
            )
        raise ValueError(
            f"No presentation named '{presentation_name}'. "
            f"Available presentations:\n" + "\n".join(available)
        )

    if app.Presentations.Count == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    # With no explicit selector, only the locked session target is valid.
    return ppt._get_pres_impl()


# ---------------------------------------------------------------------------
# Implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _create_presentation_impl(
    template_path: Optional[str],
    slide_width: Optional[float],
    slide_height: Optional[float],
    preset: Optional[str],
    activate: bool,
) -> dict:
    # Creating a new presentation legitimately needs PowerPoint, so launch it
    # if it isn't already running.
    app = ppt._get_app_impl(allow_launch=True)
    # The user just asked to create a deck — make sure it's actually visible
    # to them, even if PowerPoint was running hidden.
    if not app.Visible:
        app.Visible = True

    if template_path:
        # Create from template using Open + Untitled=msoTrue
        abs_path = os.path.abspath(template_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Template not found: {abs_path}")

        # Positional args: FileName, ReadOnly, Untitled, WithWindow
        pres = app.Presentations.Open(abs_path, 0, -1, -1)
    else:
        # Create blank presentation (existing behavior)
        pres = app.Presentations.Add()

        # Apply preset if specified
        if preset:
            preset_key = preset.strip()
            if preset_key not in SLIDE_SIZE_PRESETS:
                raise ValueError(
                    f"Unknown preset '{preset}'. Supported: {list(SLIDE_SIZE_PRESETS.keys())}"
                )
            w, h = SLIDE_SIZE_PRESETS[preset_key]
            pres.PageSetup.SlideWidth = w
            pres.PageSetup.SlideHeight = h
        elif slide_width is not None and slide_height is not None:
            pres.PageSetup.SlideWidth = slide_width
            pres.PageSetup.SlideHeight = slide_height

    template_name = ""
    try:
        template_name = pres.TemplateName
    except Exception:
        pass

    # Find the 1-based index of the newly created presentation
    pres_index = None
    for i in range(1, app.Presentations.Count + 1):
        if app.Presentations(i).Name == pres.Name:
            pres_index = i
            break

    if activate:
        try:
            pres.Windows(1).Activate()
        except Exception as e:
            logger.warning("Could not activate new presentation window: %s", e)
        ppt._target_pres_full_name = pres.FullName

    return {
        "success": True,
        "presentation_index": pres_index,
        "name": pres.Name,
        "slides_count": pres.Slides.Count,
        "slide_width": pres.PageSetup.SlideWidth,
        "slide_height": pres.PageSetup.SlideHeight,
        "template_name": template_name,
        "activated": activate,
    }


def _open_presentation_impl(
    file_path: str,
    read_only: bool,
    with_window: bool,
    activate: bool,
) -> dict:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    # Opening a file legitimately needs PowerPoint, so launch it if not running.
    app = ppt._get_app_impl(allow_launch=True)
    # The user just asked to open a deck — make sure it's actually visible
    # to them, even if PowerPoint was running hidden. Skip this when the
    # caller explicitly requested with_window=False (headless automation).
    if with_window and not app.Visible:
        app.Visible = True
    pres = app.Presentations.Open(
        FileName=file_path,
        ReadOnly=msoTrue if read_only else msoFalse,
        Untitled=msoFalse,
        WithWindow=msoTrue if with_window else msoFalse,
    )

    # Find the 1-based index of the opened presentation
    pres_index = None
    for i in range(1, app.Presentations.Count + 1):
        if app.Presentations(i).Name == pres.Name:
            pres_index = i
            break

    if activate and with_window:
        try:
            pres.Windows(1).Activate()
        except Exception as e:
            logger.warning("Could not activate opened presentation window: %s", e)
        ppt._target_pres_full_name = pres.FullName
    elif activate and not with_window:
        # Headless open — no window to activate but still set the session
        # target so subsequent tool calls land on this deck.
        ppt._target_pres_full_name = pres.FullName

    return {
        "success": True,
        "presentation_index": pres_index,
        "name": pres.Name,
        "full_name": pres.FullName,
        "slides_count": pres.Slides.Count,
        "read_only": int(pres.ReadOnly) == -1,  # msoTrue=-1; bool() misidentifies msoCTrue(1)
        "activated": activate,
    }


def _save_presentation_impl(
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app, presentation_index=presentation_index, presentation_name=presentation_name
    )
    pres.Save()
    return {
        "success": True,
        "name": pres.Name,
        "saved": bool(pres.Saved),
    }


def _save_presentation_as_impl(
    file_path: str,
    format: Optional[str],
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app, presentation_index=presentation_index, presentation_name=presentation_name
    )

    kwargs = {"FileName": file_path}
    if format:
        fmt_key = format.lower().strip()
        if fmt_key not in SAVE_FORMAT_MAP:
            raise ValueError(
                f"Unknown format '{format}'. Supported: {list(SAVE_FORMAT_MAP.keys())}"
            )
        kwargs["FileFormat"] = SAVE_FORMAT_MAP[fmt_key]

    pres.SaveAs(**kwargs)
    return {
        "success": True,
        "name": pres.Name,
        "full_name": pres.FullName,
    }


def _close_presentation_impl(
    save_changes: bool,
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app, presentation_index=presentation_index, presentation_name=presentation_name
    )
    name = pres.Name

    if save_changes:
        pres.Save()
    else:
        # Suppress "save changes?" dialog
        pres.Saved = True

    pres.Close()
    return {"success": True, "closed": name}


def _get_presentation_info_impl(
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app, presentation_index=presentation_index, presentation_name=presentation_name
    )
    page = pres.PageSetup

    template_name = ""
    try:
        template_name = pres.TemplateName
    except Exception:
        pass

    # Fonts: title/body, Latin/East Asian from theme font scheme
    fonts = {
        "title_latin": None,
        "title_east_asian": None,
        "body_latin": None,
        "body_east_asian": None,
    }
    try:
        try:
            font_scheme = pres.Designs(1).SlideMaster.Theme.ThemeFontScheme
        except Exception:
            font_scheme = pres.SlideMaster.Theme.ThemeFontScheme

        def _clean_font(name):
            if name and name.startswith("+"):
                return None
            return name or None

        fonts["body_latin"] = _clean_font(font_scheme.MinorFont(1).Name)
        fonts["body_east_asian"] = _clean_font(font_scheme.MinorFont(3).Name)
        fonts["title_latin"] = _clean_font(font_scheme.MajorFont(1).Name)
        fonts["title_east_asian"] = _clean_font(font_scheme.MajorFont(3).Name)
    except Exception:
        pass

    # Accent colors: accent1–accent6 from theme color scheme (indices 5–10)
    accent_colors = {
        "accent1": None,
        "accent2": None,
        "accent3": None,
        "accent4": None,
        "accent5": None,
        "accent6": None,
    }
    try:
        color_scheme = pres.SlideMaster.Theme.ThemeColorScheme
        accent_keys = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
        for i, key in enumerate(accent_keys, start=5):
            try:
                rgb_int = color_scheme(i).RGB
                accent_colors[key] = int_to_hex(int(rgb_int))
            except Exception:
                pass
    except Exception:
        pass

    full_name = pres.FullName
    local_path = resolve_local_path(full_name)
    local_dir = os.path.dirname(local_path) if local_path else None

    return {
        "name": pres.Name,
        "full_name": full_name,
        "local_path": local_path,
        "local_dir": local_dir,
        "path": pres.Path,
        "slides_count": pres.Slides.Count,
        "read_only": int(pres.ReadOnly) == -1,  # msoTrue=-1; bool() misidentifies msoCTrue(1)
        "saved": bool(pres.Saved),
        "slide_width": page.SlideWidth,
        "slide_height": page.SlideHeight,
        "slide_width_inches": round(page.SlideWidth / 72.0, 3),
        "slide_height_inches": round(page.SlideHeight / 72.0, 3),
        "first_slide_number": page.FirstSlideNumber,
        "template_name": template_name,
        "fonts": fonts,
        "accent_colors": accent_colors,
    }


def _get_default_templates_dir() -> Optional[str]:
    """Return only the explicitly configured template directory."""
    configured = os.getenv("PPT_TEMPLATES_DIR")
    if configured and os.path.isdir(configured):
        return os.path.abspath(configured)
    return None


def _activate_presentation_impl(
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    if presentation_index is None and presentation_name is None:
        raise ValueError(
            "Specify either presentation_index or presentation_name."
        )
    if presentation_index is not None and presentation_name is not None:
        raise ValueError(
            "Specify either presentation_index or presentation_name, not both."
        )
    name_or_index = presentation_index if presentation_index is not None else presentation_name
    return ppt._set_target_pres_impl(name_or_index)


def _list_templates_impl(templates_dir: Optional[str]) -> dict:
    """List reusable PowerPoint source decks and template files.

    This function does NOT require COM access.
    """
    if templates_dir is None:
        templates_dir = _get_default_templates_dir()

    if templates_dir is None:
        return {
            "templates_dir": None,
            "count": 0,
            "templates": [],
            "error": "Could not find a templates directory. Specify templates_dir explicitly.",
        }

    if not os.path.isdir(templates_dir):
        return {
            "templates_dir": templates_dir,
            "count": 0,
            "templates": [],
            "error": f"Directory not found: {templates_dir}",
        }

    templates = []
    for ext in ("*.pptx", "*.pptm", "*.potx", "*.potm", "*.pot"):
        pattern = os.path.join(templates_dir, ext)
        for filepath in glob_mod.glob(pattern):
            templates.append({
                "name": os.path.splitext(os.path.basename(filepath))[0],
                "file_name": os.path.basename(filepath),
                "file_path": os.path.abspath(filepath),
            })

    templates.sort(key=lambda t: t["name"])
    return {
        "templates_dir": templates_dir,
        "count": len(templates),
        "templates": templates,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (return JSON strings)
# ---------------------------------------------------------------------------
def create_presentation(params: CreatePresentationInput) -> str:
    """Create a new presentation, optionally from a template."""
    return execute_json(None,
            _create_presentation_impl,
            params.template_path,
            params.slide_width,
            params.slide_height,
            params.preset,
            params.activate,
        )


def open_presentation(params: OpenPresentationInput) -> str:
    """Open an existing presentation file."""
    return execute_json(None,
            _open_presentation_impl,
            params.file_path,
            params.read_only,
            params.with_window,
            params.activate,
        )


def save_presentation(params: SavePresentationInput) -> str:
    """Save the active or specified presentation."""
    return execute_json(None,
            _save_presentation_impl,
            params.presentation_index,
            params.presentation_name,
        )


def save_presentation_as(params: SavePresentationAsInput) -> str:
    """Save a presentation with a new name and/or format."""
    return execute_json(None,
            _save_presentation_as_impl,
            params.file_path,
            params.format,
            params.presentation_index,
            params.presentation_name,
        )


def close_presentation(params: ClosePresentationInput) -> str:
    """Close a presentation, optionally saving first."""
    return execute_json(None,
            _close_presentation_impl,
            params.save_changes,
            params.presentation_index,
            params.presentation_name,
        )


def get_presentation_info(params: GetPresentationInfoInput) -> str:
    """Get detailed info about a presentation."""
    return execute_json(None,
            _get_presentation_info_impl,
            params.presentation_index,
            params.presentation_name,
        )


def activate_presentation(params: ActivatePresentationInput) -> str:
    """Activate a presentation as the MCP session target."""
    return execute_json(None,
            _activate_presentation_impl,
            params.presentation_index,
            params.presentation_name,
        )


def list_templates(params: ListTemplatesInput) -> str:
    """List available PowerPoint template files."""
    try:
        result = _list_templates_impl(params.templates_dir)
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})

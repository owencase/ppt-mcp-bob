"""SlideShow control tools for PowerPoint COM automation.

Start, stop, navigate, and query slide show state.
"""

import time
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.tool_result import execute_json
from ppt_com.constants import (
    msoTrue,
    msoFalse,
    ppShowTypeSpeaker,
    ppShowTypeWindow,
    ppShowTypeKiosk,
    ppShowAll,
    ppShowSlideRange,
    SLIDESHOW_STATE_NAMES,
    SHOW_TYPE_NAMES,
)


SHOW_TYPE_MAP = {
    "speaker": ppShowTypeSpeaker,
    "window": ppShowTypeWindow,
    "kiosk": ppShowTypeKiosk,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SlideShowStartInput(BaseModel):
    """Input for starting a slide show."""
    model_config = ConfigDict(str_strip_whitespace=True)

    start_slide: Optional[int] = Field(
        default=None,
        description="1-based starting slide index. Default: 1.",
    )
    end_slide: Optional[int] = Field(
        default=None,
        description="1-based ending slide index. Default: last slide.",
    )
    loop: Optional[bool] = Field(
        default=None,
        description="Loop the slide show continuously.",
    )
    show_type: Optional[str] = Field(
        default=None,
        description="Show type: 'speaker' (fullscreen), 'window', or 'kiosk'. Default: 'speaker'.",
    )


class SlideShowGotoInput(BaseModel):
    """Input for navigating to a specific slide in the slide show."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based slide index to navigate to.",
        ge=1,
    )


# ---------------------------------------------------------------------------
# Implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _slideshow_start_impl(
    start_slide: Optional[int],
    end_slide: Optional[int],
    loop: Optional[bool],
    show_type: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    if app.Presentations.Count == 0:
        raise RuntimeError(
            "No presentation is open. "
            "Use ppt_create_presentation or ppt_open_presentation first."
        )
    pres = ppt._get_pres_impl()

    if pres.Slides.Count == 0:
        raise RuntimeError("Presentation has no slides.")

    settings = pres.SlideShowSettings

    # Show type
    if show_type is not None:
        type_key = show_type.lower().strip()
        if type_key not in SHOW_TYPE_MAP:
            raise ValueError(
                f"Unknown show_type '{show_type}'. Supported: {list(SHOW_TYPE_MAP.keys())}"
            )
        settings.ShowType = SHOW_TYPE_MAP[type_key]
    else:
        settings.ShowType = ppShowTypeSpeaker

    # Slide range
    if start_slide is None and end_slide is None:
        # No range specified → show all slides (resets any previous range)
        settings.RangeType = ppShowAll
        actual_start = 1
        actual_end = pres.Slides.Count
    else:
        actual_start = start_slide if start_slide is not None else 1
        actual_end = end_slide if end_slide is not None else pres.Slides.Count

        if actual_start < 1 or actual_start > pres.Slides.Count:
            raise ValueError(
                f"start_slide {actual_start} out of range (1-{pres.Slides.Count})"
            )
        if actual_end < actual_start or actual_end > pres.Slides.Count:
            raise ValueError(
                f"end_slide {actual_end} out of range ({actual_start}-{pres.Slides.Count})"
            )

        settings.RangeType = ppShowSlideRange
        settings.StartingSlide = actual_start
        settings.EndingSlide = actual_end

    # Loop
    if loop is not None:
        settings.LoopUntilStopped = msoTrue if loop else msoFalse

    # Start the show
    ssw = settings.Run()
    time.sleep(0.5)

    view = ssw.View
    return {
        "success": True,
        "show_type": SHOW_TYPE_NAMES.get(settings.ShowType, "unknown"),
        "current_slide": view.CurrentShowPosition,
        "total_slides": pres.Slides.Count,
        "start_slide": actual_start,
        "end_slide": actual_end,
    }


def _slideshow_stop_impl() -> dict:
    app = ppt._get_app_impl()
    if app.SlideShowWindows.Count == 0:
        return {"success": True, "message": "No slide show was running."}

    app.SlideShowWindows(1).View.Exit()
    return {"success": True, "message": "Slide show ended."}


def _slideshow_next_impl() -> dict:
    app = ppt._get_app_impl()
    if app.SlideShowWindows.Count == 0:
        raise RuntimeError("No slide show is running.")

    view = app.SlideShowWindows(1).View
    view.Next()
    return {
        "success": True,
        "current_slide": view.CurrentShowPosition,
        "state": SLIDESHOW_STATE_NAMES.get(view.State, f"unknown({view.State})"),
    }


def _slideshow_previous_impl() -> dict:
    app = ppt._get_app_impl()
    if app.SlideShowWindows.Count == 0:
        raise RuntimeError("No slide show is running.")

    view = app.SlideShowWindows(1).View
    view.Previous()
    return {
        "success": True,
        "current_slide": view.CurrentShowPosition,
        "state": SLIDESHOW_STATE_NAMES.get(view.State, f"unknown({view.State})"),
    }


def _slideshow_goto_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    if app.SlideShowWindows.Count == 0:
        raise RuntimeError("No slide show is running.")

    view = app.SlideShowWindows(1).View
    view.GotoSlide(slide_index)
    return {
        "success": True,
        "current_slide": view.CurrentShowPosition,
        "state": SLIDESHOW_STATE_NAMES.get(view.State, f"unknown({view.State})"),
    }


def _slideshow_get_status_impl() -> dict:
    app = ppt._get_app_impl()

    if app.SlideShowWindows.Count == 0:
        return {"running": False}

    ssw = app.SlideShowWindows(1)
    view = ssw.View

    return {
        "running": True,
        "current_slide": view.CurrentShowPosition,
        "state": view.State,
        "state_name": SLIDESHOW_STATE_NAMES.get(view.State, f"unknown({view.State})"),
        "pointer_type": view.PointerType,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (return JSON strings)
# ---------------------------------------------------------------------------
def slideshow_start(params: SlideShowStartInput) -> str:
    """Start a slide show presentation."""
    return execute_json(None,
            _slideshow_start_impl,
            params.start_slide,
            params.end_slide,
            params.loop,
            params.show_type,
        )


def slideshow_stop() -> str:
    """Stop the running slide show."""
    return execute_json(None, _slideshow_stop_impl)


def slideshow_next() -> str:
    """Navigate to the next slide in the running slide show."""
    return execute_json(None, _slideshow_next_impl)


def slideshow_previous() -> str:
    """Navigate to the previous slide in the running slide show."""
    return execute_json(None, _slideshow_previous_impl)


def slideshow_goto(params: SlideShowGotoInput) -> str:
    """Go to a specific slide in the running slide show."""
    return execute_json(None, _slideshow_goto_impl, params.slide_index)


def slideshow_get_status() -> str:
    """Get the current state of the running slide show."""
    return execute_json(None, _slideshow_get_status_impl)

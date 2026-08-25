"""Hyperlink operations for PowerPoint COM automation.

Handles adding, listing, and removing hyperlinks on shapes.
Supports click and mouseover actions with URL, file, mailto, and slide links.
"""

from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.tool_result import execute_json
from utils.com_objects import get_slide
from utils.com_objects import get_shape as _get_shape
from ppt_com.constants import (
    ppActionNone, ppActionHyperlink,
    ppMouseClick, ppMouseOver,
)


ACTION_ON_MAP: dict[str, int] = {
    "click": ppMouseClick,
    "mouseover": ppMouseOver,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddHyperlinkInput(BaseModel):
    """Input for adding a hyperlink to a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    address: str = Field(
        ..., description="Hyperlink URL, file path, or mailto: address"
    )
    sub_address: Optional[str] = Field(
        default=None,
        description="Sub-address for slide links (e.g. '3,,' to link to slide 3)",
    )
    screen_tip: Optional[str] = Field(
        default=None, description="Tooltip text shown on hover"
    )
    action_on: str = Field(
        default="click",
        description="Trigger action: 'click' or 'mouseover'",
    )


class GetHyperlinksInput(BaseModel):
    """Input for listing hyperlinks on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class RemoveHyperlinkInput(BaseModel):
    """Input for removing a hyperlink from a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    action_on: str = Field(
        default="click",
        description="Which action to remove: 'click' or 'mouseover'",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or 1-based index
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _add_hyperlink_impl(slide_index, shape_name_or_index, address, sub_address, screen_tip, action_on):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    action_key = action_on.strip().lower()
    if action_key not in ACTION_ON_MAP:
        raise ValueError(
            f"Unknown action_on '{action_on}'. Use: {', '.join(ACTION_ON_MAP.keys())}"
        )
    action_idx = ACTION_ON_MAP[action_key]

    # CRITICAL: Must set Action = ppActionHyperlink BEFORE setting Hyperlink.Address
    action_setting = shape.ActionSettings(action_idx)
    action_setting.Action = ppActionHyperlink
    action_setting.Hyperlink.Address = address
    if sub_address is not None:
        action_setting.Hyperlink.SubAddress = sub_address
    if screen_tip is not None:
        action_setting.Hyperlink.ScreenTip = screen_tip

    return {
        "success": True,
        "shape_name": shape.Name,
        "address": address,
        "sub_address": sub_address,
        "action_on": action_key,
    }


def _get_hyperlinks_impl(slide_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    hyperlinks = []
    for i in range(1, slide.Hyperlinks.Count + 1):
        hl = slide.Hyperlinks(i)
        hyperlinks.append({
            "index": i,
            "address": hl.Address,
            "sub_address": hl.SubAddress,
            "type": hl.Type,
        })

    return {
        "success": True,
        "slide_index": slide_index,
        "hyperlinks_count": slide.Hyperlinks.Count,
        "hyperlinks": hyperlinks,
    }


def _remove_hyperlink_impl(slide_index, shape_name_or_index, action_on):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    action_key = action_on.strip().lower()
    if action_key not in ACTION_ON_MAP:
        raise ValueError(
            f"Unknown action_on '{action_on}'. Use: {', '.join(ACTION_ON_MAP.keys())}"
        )
    action_idx = ACTION_ON_MAP[action_key]

    shape.ActionSettings(action_idx).Action = ppActionNone

    return {
        "success": True,
        "shape_name": shape.Name,
        "action_on": action_key,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def add_hyperlink(params: AddHyperlinkInput) -> str:
    """Add a hyperlink to a shape.

    Args:
        params: Hyperlink parameters including shape, address, and action trigger.

    Returns:
        JSON with shape name and hyperlink address.
    """
    return execute_json('Failed to add hyperlink: ',
            _add_hyperlink_impl,
            params.slide_index, params.shape_name_or_index,
            params.address, params.sub_address, params.screen_tip,
            params.action_on,
        )


def get_hyperlinks(params: GetHyperlinksInput) -> str:
    """Get all hyperlinks on a slide.

    Args:
        params: Slide index to list hyperlinks from.

    Returns:
        JSON with hyperlinks count and list of hyperlink details.
    """
    return execute_json('Failed to get hyperlinks: ',
            _get_hyperlinks_impl,
            params.slide_index,
        )


def remove_hyperlink(params: RemoveHyperlinkInput) -> str:
    """Remove a hyperlink from a shape.

    Args:
        params: Shape identifier and which action trigger to remove.

    Returns:
        JSON confirming the hyperlink removal.
    """
    return execute_json('Failed to remove hyperlink: ',
            _remove_hyperlink_impl,
            params.slide_index, params.shape_name_or_index,
            params.action_on,
        )

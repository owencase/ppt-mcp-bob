"""Shape grouping and ungrouping tools for PowerPoint COM automation.

Handles grouping multiple shapes, ungrouping group shapes,
and inspecting the items within a group.
"""

from typing import Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.tool_result import execute_json
from utils.com_objects import get_slide
from utils.com_objects import get_shape as _get_shape
from ppt_com.constants import msoGroup, SHAPE_TYPE_NAMES



# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class GroupShapesInput(BaseModel):
    """Input for grouping shapes on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ..., min_length=2,
        description="List of shape names to group (minimum 2)",
    )


class UngroupShapesInput(BaseModel):
    """Input for ungrouping a group shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Group shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )


class GetGroupItemsInput(BaseModel):
    """Input for getting items within a group shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Group shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _group_shapes_impl(slide_index, shape_names):
    app, pres, slide = get_slide(slide_index)

    # Validate all shape names exist before grouping
    for name in shape_names:
        found = False
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name:
                found = True
                break
        if not found:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")

    shape_range = slide.Shapes.Range(shape_names)
    group = shape_range.Group()

    return {
        "success": True,
        "group_name": group.Name,
        "shape_index": group.ZOrderPosition,
    }


def _ungroup_shapes_impl(slide_index, shape_name_or_index):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if shape.Type != msoGroup:
        raise ValueError(
            f"Shape '{shape.Name}' is not a group (type={shape.Type}). "
            f"Only group shapes (type={msoGroup}) can be ungrouped."
        )

    ungrouped = shape.Ungroup()
    names = []
    for i in range(1, ungrouped.Count + 1):
        names.append(ungrouped(i).Name)

    return {
        "success": True,
        "ungrouped_count": ungrouped.Count,
        "shape_names": names,
    }


def _get_group_items_impl(slide_index, shape_name_or_index):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    if shape.Type != msoGroup:
        raise ValueError(
            f"Shape '{shape.Name}' is not a group (type={shape.Type}). "
            f"Only group shapes (type={msoGroup}) can be inspected."
        )

    items = []
    for i in range(1, shape.GroupItems.Count + 1):
        item = shape.GroupItems(i)
        type_val = item.Type
        items.append({
            "name": item.Name,
            "type": type_val,
            "type_name": SHAPE_TYPE_NAMES.get(type_val, f"Unknown({type_val})"),
            "left": round(item.Left, 2),
            "top": round(item.Top, 2),
            "width": round(item.Width, 2),
            "height": round(item.Height, 2),
        })

    return {
        "success": True,
        "group_name": shape.Name,
        "items": items,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def group_shapes(params: GroupShapesInput) -> str:
    """Group multiple shapes into a single group shape.

    Args:
        params: Slide index and list of shape names to group.

    Returns:
        JSON with group name and shape index.
    """
    return execute_json('Failed to group shapes: ',
            _group_shapes_impl,
            params.slide_index, params.shape_names,
        )


def ungroup_shapes(params: UngroupShapesInput) -> str:
    """Ungroup a group shape into its individual shapes.

    Args:
        params: Slide index and group shape identifier.

    Returns:
        JSON with ungrouped count and shape names.
    """
    return execute_json('Failed to ungroup shapes: ',
            _ungroup_shapes_impl,
            params.slide_index, params.shape_name_or_index,
        )


def get_group_items(params: GetGroupItemsInput) -> str:
    """Get information about all items within a group shape.

    Args:
        params: Slide index and group shape identifier.

    Returns:
        JSON with group name and list of item details.
    """
    return execute_json('Failed to get group items: ',
            _get_group_items_impl,
            params.slide_index, params.shape_name_or_index,
        )

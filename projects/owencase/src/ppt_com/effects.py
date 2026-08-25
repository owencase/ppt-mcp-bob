"""Visual effect tools (glow, reflection, soft edge) for PowerPoint COM automation.

Handles glow, reflection, and soft edge effects on shapes.
"""

from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.tool_result import execute_json
from utils.com_objects import get_slide
from utils.com_objects import get_shape as _get_shape
from utils.color import hex_to_int



# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SetGlowInput(BaseModel):
    """Input for setting glow effect on a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    radius: float = Field(
        ..., ge=0, description="Glow radius in points (0 to remove glow)"
    )
    color: Optional[str] = Field(
        default=None, description="Glow color as '#RRGGBB' hex"
    )
    transparency: Optional[float] = Field(
        default=None, ge=0, le=1,
        description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )


class SetReflectionInput(BaseModel):
    """Input for setting reflection effect on a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    reflection_type: Optional[int] = Field(
        default=None, ge=0, le=9,
        description="MsoReflectionType: 0=none, 1-9=presets"
    )
    blur: Optional[float] = Field(
        default=None, ge=0, description="Reflection blur radius in points"
    )
    offset: Optional[float] = Field(
        default=None, ge=0, description="Reflection offset in points"
    )
    size: Optional[float] = Field(
        default=None, ge=0, le=100,
        description="Reflection size as percentage (0-100)"
    )
    transparency: Optional[float] = Field(
        default=None, ge=0, le=1,
        description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )


class SetSoftEdgeInput(BaseModel):
    """Input for setting soft edge effect on a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    radius: float = Field(
        ..., ge=0, description="Soft edge radius in points (0 to remove)"
    )


# ---------------------------------------------------------------------------
# COM implementation functions
# ---------------------------------------------------------------------------
def _set_glow_impl(slide_index, shape_name_or_index, radius,
                    color, transparency) -> dict:
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    glow = shape.Glow
    glow.Radius = radius

    if color is not None:
        glow.Color.RGB = hex_to_int(color)

    if transparency is not None:
        glow.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
        "glow_radius": radius,
    }


def _set_reflection_impl(slide_index, shape_name_or_index, reflection_type,
                          blur, offset, size, transparency) -> dict:
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    reflection = shape.Reflection

    if reflection_type is not None:
        reflection.Type = reflection_type

    if blur is not None:
        reflection.Blur = blur

    if offset is not None:
        reflection.Offset = offset

    if size is not None:
        reflection.Size = size

    if transparency is not None:
        reflection.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
    }


def _set_soft_edge_impl(slide_index, shape_name_or_index, radius) -> dict:
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shape.SoftEdge.Radius = radius

    return {
        "status": "success",
        "shape_name": shape.Name,
        "soft_edge_radius": radius,
    }


# ---------------------------------------------------------------------------
# MCP tool functions
# ---------------------------------------------------------------------------
def set_glow(params: SetGlowInput) -> str:
    """Set glow effect on a shape."""
    return execute_json(None,
            _set_glow_impl,
            params.slide_index, params.shape_name_or_index, params.radius,
            params.color, params.transparency,
        )


def set_reflection(params: SetReflectionInput) -> str:
    """Set reflection effect on a shape."""
    return execute_json(None,
            _set_reflection_impl,
            params.slide_index, params.shape_name_or_index,
            params.reflection_type, params.blur, params.offset,
            params.size, params.transparency,
        )


def set_soft_edge(params: SetSoftEdgeInput) -> str:
    """Set soft edge effect on a shape."""
    return execute_json(None,
            _set_soft_edge_impl,
            params.slide_index, params.shape_name_or_index, params.radius,
        )

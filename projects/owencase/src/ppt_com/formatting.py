"""Fill, line, and shadow effect tools for PowerPoint COM automation."""

from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.tool_result import execute_json
from utils.com_objects import get_slide
from utils.com_objects import get_shape as _get_shape
from utils.color import hex_to_int
from ppt_com.constants import (
    msoTrue,
    msoFalse,
    msoGradientHorizontal,
    msoGradientVertical,
    msoGradientDiagonalUp,
    msoGradientDiagonalDown,
    msoGradientFromCorner,
    msoGradientFromCenter,
    msoLineSolid,
    msoLineRoundDot,
    msoLineDash,
    msoLineDashDot,
    msoLineLongDash,
)



# ---------------------------------------------------------------------------
# Helper: reuse _get_shape from text module
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# Constant maps
# ---------------------------------------------------------------------------
GRADIENT_STYLE_MAP = {
    "horizontal": msoGradientHorizontal,
    "vertical": msoGradientVertical,
    "diagonal_up": msoGradientDiagonalUp,
    "diagonal_down": msoGradientDiagonalDown,
    "from_corner": msoGradientFromCorner,
    "from_center": msoGradientFromCenter,
}

DASH_STYLE_MAP = {
    "solid": msoLineSolid,
    "round_dot": msoLineRoundDot,
    "dash": msoLineDash,
    "dash_dot": msoLineDashDot,
    "long_dash": msoLineLongDash,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SetFillInput(BaseModel):
    """Input for setting shape fill."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    fill_type: str = Field(
        ..., description="'solid', 'gradient', or 'none'"
    )
    color: Optional[str] = Field(
        default=None, description="Fill color as '#RRGGBB' hex (for solid fill)"
    )
    gradient_color1: Optional[str] = Field(
        default=None, description="Gradient start color as '#RRGGBB'"
    )
    gradient_color2: Optional[str] = Field(
        default=None, description="Gradient end color as '#RRGGBB'"
    )
    gradient_style: Optional[str] = Field(
        default=None,
        description="'horizontal', 'vertical', 'diagonal_up', 'diagonal_down', 'from_corner', or 'from_center'"
    )
    transparency: Optional[float] = Field(
        default=None, description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )


class SetLineInput(BaseModel):
    """Input for setting shape border/line."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    color: Optional[str] = Field(default=None, description="Line color as '#RRGGBB'")
    weight: Optional[float] = Field(default=None, description="Line weight in points")
    dash_style: Optional[str] = Field(
        default=None,
        description="'solid', 'round_dot', 'dash', 'dash_dot', or 'long_dash'"
    )
    visible: Optional[bool] = Field(default=None, description="Line visible on/off")
    transparency: Optional[float] = Field(
        default=None, description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )


class SetShadowInput(BaseModel):
    """Input for setting shadow effect."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    visible: bool = Field(..., description="Shadow visible on/off")
    blur: Optional[float] = Field(default=None, description="Shadow blur radius in points")
    offset_x: Optional[float] = Field(default=None, description="Shadow horizontal offset in points")
    offset_y: Optional[float] = Field(default=None, description="Shadow vertical offset in points")
    color: Optional[str] = Field(default=None, description="Shadow color as '#RRGGBB'")
    transparency: Optional[float] = Field(
        default=None, description="Transparency 0.0 (opaque) to 1.0 (fully transparent)"
    )


# ---------------------------------------------------------------------------
# COM implementation functions
# ---------------------------------------------------------------------------
def _set_fill_impl(slide_index, shape_name_or_index, fill_type,
                    color, gradient_color1, gradient_color2, gradient_style,
                    transparency) -> dict:
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    fill = shape.Fill

    if fill_type == "none":
        fill.Visible = msoFalse
    elif fill_type == "solid":
        fill.Solid()
        if color is not None:
            fill.ForeColor.RGB = hex_to_int(color)
    elif fill_type == "gradient":
        style_val = GRADIENT_STYLE_MAP.get(gradient_style, msoGradientHorizontal)
        fill.TwoColorGradient(Style=style_val, Variant=1)
        if gradient_color1 is not None:
            fill.ForeColor.RGB = hex_to_int(gradient_color1)
        if gradient_color2 is not None:
            fill.BackColor.RGB = hex_to_int(gradient_color2)
    else:
        raise ValueError(
            f"Invalid fill_type '{fill_type}'. Valid values: 'solid', 'gradient', 'none'"
        )

    if transparency is not None and fill_type != "none":
        fill.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
        "fill_type": fill_type,
    }


def _set_line_impl(slide_index, shape_name_or_index,
                    color, weight, dash_style, visible, transparency) -> dict:
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    line = shape.Line

    if visible is not None:
        line.Visible = msoTrue if visible else msoFalse

    if color is not None:
        line.ForeColor.RGB = hex_to_int(color)

    if weight is not None:
        line.Weight = weight

    if dash_style is not None:
        dash_val = DASH_STYLE_MAP.get(dash_style)
        if dash_val is None:
            raise ValueError(
                f"Invalid dash_style '{dash_style}'. "
                f"Valid values: {list(DASH_STYLE_MAP.keys())}"
            )
        line.DashStyle = dash_val

    if transparency is not None:
        line.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
    }


def _set_shadow_impl(slide_index, shape_name_or_index,
                      visible, blur, offset_x, offset_y, color, transparency) -> dict:
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    shadow = shape.Shadow

    shadow.Visible = msoTrue if visible else msoFalse

    if visible:
        if blur is not None:
            shadow.Blur = blur
        if offset_x is not None:
            shadow.OffsetX = offset_x
        if offset_y is not None:
            shadow.OffsetY = offset_y
        if color is not None:
            shadow.ForeColor.RGB = hex_to_int(color)
        if transparency is not None:
            shadow.Transparency = transparency

    return {
        "status": "success",
        "shape_name": shape.Name,
        "shadow_visible": visible,
    }


# ---------------------------------------------------------------------------
# MCP tool functions
# ---------------------------------------------------------------------------
def set_fill(params: SetFillInput) -> str:
    """Set shape fill (solid, gradient, or none)."""
    return execute_json(None,
            _set_fill_impl,
            params.slide_index, params.shape_name_or_index, params.fill_type,
            params.color, params.gradient_color1, params.gradient_color2,
            params.gradient_style, params.transparency,
        )


def set_line(params: SetLineInput) -> str:
    """Set shape border/line properties."""
    return execute_json(None,
            _set_line_impl,
            params.slide_index, params.shape_name_or_index,
            params.color, params.weight, params.dash_style,
            params.visible, params.transparency,
        )


def set_shadow(params: SetShadowInput) -> str:
    """Set shadow effect on a shape."""
    return execute_json(None,
            _set_shadow_impl,
            params.slide_index, params.shape_name_or_index,
            params.visible, params.blur, params.offset_x, params.offset_y,
            params.color, params.transparency,
        )

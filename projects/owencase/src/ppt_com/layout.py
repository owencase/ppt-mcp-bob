"""Layout tools for PowerPoint COM automation.

Handles shape alignment, distribution, flipping, merging,
slide size, and slide background configuration.
"""

import os
from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict, model_validator

from utils.com_wrapper import ppt
from utils.tool_result import execute_json
from utils.com_objects import get_slide
from utils.com_objects import get_shape as _get_shape
from utils.navigation import goto_slide
from utils.color import hex_to_int
from ppt_com.constants import (
    msoTrue,
    msoFalse,
    ALIGN_CMD_MAP,
    DISTRIBUTE_CMD_MAP,
    FLIP_CMD_MAP,
    MERGE_CMD_MAP,
    SLIDE_SIZE_MAP,
    GRADIENT_STYLE_MAP,
)



# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AlignShapesInput(BaseModel):
    """Input for aligning shapes on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ..., min_length=2,
        description="List of shape names to align (minimum 2)",
    )
    align_to: str = Field(
        ...,
        description="Alignment direction: 'left', 'center', 'right', 'top', 'middle', or 'bottom'",
    )
    relative_to_slide: bool = Field(
        default=False,
        description="If true, align relative to the slide; otherwise align relative to each other",
    )


class DistributeShapesInput(BaseModel):
    """Input for distributing shapes evenly on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ..., min_length=3,
        description="List of shape names to distribute (minimum 3)",
    )
    direction: str = Field(
        ...,
        description="Distribution direction: 'horizontal' or 'vertical'",
    )
    relative_to_slide: bool = Field(
        default=False,
        description="If true, distribute relative to the slide edges; otherwise between outermost shapes",
    )


class GetSlideSizeInput(BaseModel):
    """Input for getting slide size (no parameters needed)."""
    model_config = ConfigDict(str_strip_whitespace=True)


class SetSlideSizeInput(BaseModel):
    """Input for setting slide size."""
    model_config = ConfigDict(str_strip_whitespace=True)

    width: Optional[float] = Field(
        default=None,
        description="Slide width in points (72 points = 1 inch)",
    )
    height: Optional[float] = Field(
        default=None,
        description="Slide height in points (72 points = 1 inch)",
    )
    preset: Optional[str] = Field(
        default=None,
        description="Preset size: '16:9', '4:3', 'a4', 'letter', 'widescreen', '16:10', '35mm', 'overhead', 'banner'",
    )
    orientation: Optional[str] = Field(
        default=None,
        description="Orientation: 'landscape' or 'portrait'",
    )


class SetSlideBackgroundInput(BaseModel):
    """Input for setting a slide's background."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: Optional[int] = Field(
        default=None, ge=1,
        description="1-based slide index. Required if slide_indices is not provided.",
    )
    slide_indices: Optional[list[int]] = Field(
        default=None,
        description="List of 1-based slide indices to apply the background to. "
        "Overrides slide_index when provided.",
    )
    fill_type: str = Field(
        ...,
        description="Fill type: 'solid', 'gradient', 'picture', 'none', or 'master'",
    )
    color: Optional[str] = Field(
        default=None,
        description="Solid fill color as '#RRGGBB'",
    )
    gradient_color1: Optional[str] = Field(
        default=None,
        description="First gradient color as '#RRGGBB'",
    )
    gradient_color2: Optional[str] = Field(
        default=None,
        description="Second gradient color as '#RRGGBB'",
    )
    gradient_style: Optional[str] = Field(
        default=None,
        description="Gradient style: 'horizontal', 'vertical', 'diagonal_up', 'diagonal_down', 'from_corner', 'from_title', 'from_center'",
    )
    image_path: Optional[str] = Field(
        default=None,
        description="Absolute path to image file for picture fill",
    )
    transparency: Optional[float] = Field(
        default=None, ge=0.0, le=1.0,
        description="Fill transparency (0 = opaque, 1 = fully transparent)",
    )

    @model_validator(mode="after")
    def validate_slide_target(self):
        """Ensure at least one slide target is provided and indices are valid."""
        if self.slide_indices is not None and len(self.slide_indices) == 0:
            raise ValueError("slide_indices must not be empty")
        if self.slide_index is None and not self.slide_indices:
            raise ValueError(
                "Either slide_index or slide_indices must be provided"
            )
        if self.slide_indices is not None:
            for idx in self.slide_indices:
                if idx < 1:
                    raise ValueError(
                        f"All slide indices must be >= 1, got {idx}"
                    )
        return self


class FlipShapeInput(BaseModel):
    """Input for flipping a shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_name_or_index: Union[str, int] = Field(
        ..., description="Shape name (str) or 1-based index (int). Prefer name — indices shift when shapes are added/removed"
    )
    direction: str = Field(
        ...,
        description="Flip direction: 'horizontal' or 'vertical'",
    )


class MergeShapesInput(BaseModel):
    """Input for merging shapes using Boolean operations."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    shape_names: list[str] = Field(
        ..., min_length=2,
        description="List of shape names to merge (minimum 2)",
    )
    merge_type: str = Field(
        ...,
        description="Merge type: 'union', 'combine', 'intersect', 'subtract', or 'fragment'",
    )
    primary_shape: Optional[str] = Field(
        default=None,
        description="Name of the primary shape (determines formatting of result). If omitted, the first shape in the range is used.",
    )


# ---------------------------------------------------------------------------
# Helper: find a shape by name or index
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _align_shapes_impl(slide_index, shape_names, align_to, relative_to_slide):
    app, pres, slide = get_slide(slide_index)

    # Validate align_to
    align_key = align_to.strip().lower()
    align_cmd = ALIGN_CMD_MAP.get(align_key)
    if align_cmd is None:
        raise ValueError(
            f"Unknown align_to '{align_to}'. "
            f"Valid values: {list(ALIGN_CMD_MAP.keys())}"
        )

    # Validate all shape names exist
    for name in shape_names:
        found = False
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name:
                found = True
                break
        if not found:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")

    relative = msoTrue if relative_to_slide else msoFalse
    shape_range = slide.Shapes.Range(tuple(shape_names))
    shape_range.Align(align_cmd, relative)

    return {
        "success": True,
        "aligned_count": len(shape_names),
        "align_to": align_key,
        "relative_to_slide": relative_to_slide,
    }


def _distribute_shapes_impl(slide_index, shape_names, direction, relative_to_slide):
    app, pres, slide = get_slide(slide_index)

    # Validate direction
    dir_key = direction.strip().lower()
    dist_cmd = DISTRIBUTE_CMD_MAP.get(dir_key)
    if dist_cmd is None:
        raise ValueError(
            f"Unknown direction '{direction}'. "
            f"Valid values: {list(DISTRIBUTE_CMD_MAP.keys())}"
        )

    # Validate all shape names exist
    for name in shape_names:
        found = False
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name:
                found = True
                break
        if not found:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")

    relative = msoTrue if relative_to_slide else msoFalse
    shape_range = slide.Shapes.Range(tuple(shape_names))
    shape_range.Distribute(dist_cmd, relative)

    return {
        "success": True,
        "distributed_count": len(shape_names),
        "direction": dir_key,
        "relative_to_slide": relative_to_slide,
    }


def _get_slide_size_impl():
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    ps = pres.PageSetup

    width_pt = ps.SlideWidth
    height_pt = ps.SlideHeight
    slide_size = ps.SlideSize
    orientation = ps.SlideOrientation

    # Reverse-lookup the preset name
    preset_name = None
    for name, val in SLIDE_SIZE_MAP.items():
        if val == slide_size:
            preset_name = name
            break

    return {
        "success": True,
        "width_points": round(width_pt, 2),
        "height_points": round(height_pt, 2),
        "width_inches": round(width_pt / 72.0, 4),
        "height_inches": round(height_pt / 72.0, 4),
        "slide_size_type": slide_size,
        "slide_size_name": preset_name,
        "orientation": "landscape" if orientation == 1 else "portrait",
    }


def _set_slide_size_impl(width, height, preset, orientation):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    ps = pres.PageSetup

    # Preset dimensions in points (width, height)
    PRESET_DIMENSIONS = {
        "16:9": (960, 540),
        "widescreen": (960, 540),
        "4:3": (960, 720),
        "16:10": (960, 600),
        "a4": (842, 595),
        "a3": (1191, 842),
        "letter": (792, 612),
        "35mm": (792, 528),
        "overhead": (720, 540),
        "banner": (576, 72),
    }

    # Set preset FIRST (may change width/height)
    if preset is not None:
        preset_key = preset.strip().lower()
        dims = PRESET_DIMENSIONS.get(preset_key)
        if dims is None:
            raise ValueError(
                f"Unknown preset '{preset}'. "
                f"Valid values: {list(PRESET_DIMENSIONS.keys())}"
            )
        ps.SlideWidth = dims[0]
        ps.SlideHeight = dims[1]

    # Then set explicit width/height (overrides preset dimensions)
    if width is not None:
        ps.SlideWidth = width
    if height is not None:
        ps.SlideHeight = height

    # Then set orientation
    if orientation is not None:
        orient_key = orientation.strip().lower()
        if orient_key == "landscape":
            ps.SlideOrientation = 1
        elif orient_key == "portrait":
            ps.SlideOrientation = 2
        else:
            raise ValueError(
                f"Unknown orientation '{orientation}'. "
                f"Valid values: 'landscape', 'portrait'"
            )

    # Read back final values
    return {
        "success": True,
        "width_points": round(ps.SlideWidth, 2),
        "height_points": round(ps.SlideHeight, 2),
        "width_inches": round(ps.SlideWidth / 72.0, 4),
        "height_inches": round(ps.SlideHeight / 72.0, 4),
    }


def _set_slide_background_impl(slide_index, fill_type, color,
                                gradient_color1, gradient_color2,
                                gradient_style, image_path, transparency,
                                slide_indices=None):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()

    # Determine target slides
    targets = slide_indices if slide_indices else [slide_index]

    # Validate parameters once before the loop
    fill_key = fill_type.strip().lower()

    if fill_key == "solid":
        if color is None:
            raise ValueError("color is required for solid fill")
        color_int = hex_to_int(color)
    elif fill_key == "gradient":
        if gradient_color1 is None or gradient_color2 is None:
            raise ValueError(
                "gradient_color1 and gradient_color2 are required for gradient fill"
            )
        style_key = (gradient_style or "horizontal").strip().lower()
        style_val = GRADIENT_STYLE_MAP.get(style_key)
        if style_val is None:
            raise ValueError(
                f"Unknown gradient_style '{gradient_style}'. "
                f"Valid values: {list(GRADIENT_STYLE_MAP.keys())}"
            )
        color1_int = hex_to_int(gradient_color1)
        color2_int = hex_to_int(gradient_color2)
    elif fill_key == "picture":
        if image_path is None:
            raise ValueError("image_path is required for picture fill")
        abs_path = os.path.abspath(image_path)
        if not os.path.isfile(abs_path):
            raise ValueError(f"Image file not found: {abs_path}")
    elif fill_key not in ("none", "master"):
        raise ValueError(
            f"Unknown fill_type '{fill_type}'. "
            f"Valid values: 'solid', 'gradient', 'picture', 'none', 'master'"
        )

    applied = []
    for idx in targets:
        goto_slide(app, idx)
        slide = pres.Slides(idx)

        if fill_key == "master":
            slide.FollowMasterBackground = msoTrue
        else:
            # Detach from master background
            slide.FollowMasterBackground = msoFalse
            fill = slide.Background.Fill

            if fill_key == "solid":
                fill.Solid()
                fill.ForeColor.RGB = color_int

            elif fill_key == "gradient":
                fill.TwoColorGradient(style_val, 1)
                fill.ForeColor.RGB = color1_int
                fill.BackColor.RGB = color2_int

            elif fill_key == "picture":
                fill.UserPicture(abs_path)

            elif fill_key == "none":
                fill.Background()

            # Apply transparency if specified
            if transparency is not None and fill_key not in ("none", "master"):
                fill.Transparency = transparency

        applied.append(idx)

    result = {
        "success": True,
        "slide_indices": applied,
        "fill_type": fill_key,
    }
    # Backward compatibility: include slide_index when called with single target
    if slide_indices is None:
        result["slide_index"] = applied[0]
    return result


def _flip_shape_impl(slide_index, shape_name_or_index, direction):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, shape_name_or_index)

    dir_key = direction.strip().lower()
    flip_cmd = FLIP_CMD_MAP.get(dir_key)
    if flip_cmd is None:
        raise ValueError(
            f"Unknown direction '{direction}'. "
            f"Valid values: {list(FLIP_CMD_MAP.keys())}"
        )

    shape.Flip(flip_cmd)

    # Read back flip state
    h_flip = shape.HorizontalFlip
    v_flip = shape.VerticalFlip

    return {
        "success": True,
        "shape_name": shape.Name,
        "horizontal_flip": bool(h_flip),
        "vertical_flip": bool(v_flip),
    }


def _merge_shapes_impl(slide_index, shape_names, merge_type, primary_shape):
    app, pres, slide = get_slide(slide_index)

    # Validate merge type
    merge_key = merge_type.strip().lower()
    merge_cmd = MERGE_CMD_MAP.get(merge_key)
    if merge_cmd is None:
        raise ValueError(
            f"Unknown merge_type '{merge_type}'. "
            f"Valid values: {list(MERGE_CMD_MAP.keys())}"
        )

    # Validate all shape names exist
    for name in shape_names:
        found = False
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name:
                found = True
                break
        if not found:
            raise ValueError(f"Shape '{name}' not found on slide {slide_index}")

    shape_range = slide.Shapes.Range(tuple(shape_names))

    # MergeShapes always requires a primary shape reference.
    # If not specified, use the first shape in the list.
    if primary_shape is not None:
        primary = _get_shape(slide, primary_shape)
    else:
        primary = _get_shape(slide, shape_names[0])
    shape_range.MergeShapes(merge_cmd, primary)

    return {
        "success": True,
        "merge_type": merge_key,
        "merged_count": len(shape_names),
    }


# ---------------------------------------------------------------------------
# MCP tool functions (sync wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def align_shapes(params: AlignShapesInput) -> str:
    """Align multiple shapes on a slide.

    Args:
        params: Slide index, shape names, alignment direction, and relative flag.

    Returns:
        JSON confirming the alignment operation.
    """
    return execute_json('Failed to align shapes: ',
            _align_shapes_impl,
            params.slide_index, params.shape_names,
            params.align_to, params.relative_to_slide,
        )


def distribute_shapes(params: DistributeShapesInput) -> str:
    """Distribute shapes evenly on a slide.

    Args:
        params: Slide index, shape names, direction, and relative flag.

    Returns:
        JSON confirming the distribution operation.
    """
    return execute_json('Failed to distribute shapes: ',
            _distribute_shapes_impl,
            params.slide_index, params.shape_names,
            params.direction, params.relative_to_slide,
        )


def get_slide_size(params: GetSlideSizeInput) -> str:
    """Get the current slide size of the active presentation.

    Args:
        params: No parameters needed.

    Returns:
        JSON with slide dimensions in points and inches.
    """
    return execute_json('Failed to get slide size: ', _get_slide_size_impl)


def set_slide_size(params: SetSlideSizeInput) -> str:
    """Set the slide size of the active presentation.

    Args:
        params: Optional preset, width, height, and orientation.

    Returns:
        JSON with the resulting slide dimensions.
    """
    return execute_json('Failed to set slide size: ',
            _set_slide_size_impl,
            params.width, params.height,
            params.preset, params.orientation,
        )


def set_slide_background(params: SetSlideBackgroundInput) -> str:
    """Set the background of a specific slide.

    Args:
        params: Slide index, fill type, and fill-specific options.

    Returns:
        JSON confirming the background change.
    """
    return execute_json('Failed to set slide background: ',
            _set_slide_background_impl,
            params.slide_index, params.fill_type,
            params.color, params.gradient_color1,
            params.gradient_color2, params.gradient_style,
            params.image_path, params.transparency,
            params.slide_indices,
        )


def flip_shape(params: FlipShapeInput) -> str:
    """Flip a shape horizontally or vertically.

    Args:
        params: Slide index, shape identifier, and flip direction.

    Returns:
        JSON with the resulting flip state.
    """
    return execute_json('Failed to flip shape: ',
            _flip_shape_impl,
            params.slide_index, params.shape_name_or_index,
            params.direction,
        )


def merge_shapes(params: MergeShapesInput) -> str:
    """Merge shapes using a Boolean operation.

    Args:
        params: Slide index, shape names, merge type, and optional primary shape.

    Returns:
        JSON confirming the merge operation.
    """
    return execute_json('Failed to merge shapes: ',
            _merge_shapes_impl,
            params.slide_index, params.shape_names,
            params.merge_type, params.primary_shape,
        )

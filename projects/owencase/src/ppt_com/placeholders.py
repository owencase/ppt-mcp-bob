"""Placeholder operations for PowerPoint COM automation."""

from typing import Optional, Union

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.tool_result import execute_json
from utils.com_objects import get_slide
from utils.color import int_to_hex
from ppt_com.constants import (
    ppPlaceholderTitle,
    ppPlaceholderBody,
    ppPlaceholderCenterTitle,
    ppPlaceholderSubtitle,
    PLACEHOLDER_TYPE_NAMES,
)



# ---------------------------------------------------------------------------
# Placeholder type name/index maps
# ---------------------------------------------------------------------------
PLACEHOLDER_TYPE_MAP = {
    "title": ppPlaceholderTitle,
    "body": ppPlaceholderBody,
    "center_title": ppPlaceholderCenterTitle,
    "subtitle": ppPlaceholderSubtitle,
    "object": 7,
    "slide_number": 13,
    "footer": 15,
    "date": 16,
    "picture": 18,
}

CONTAINED_TYPE_NAMES = {
    1: "AutoShape",
    3: "Chart",
    11: "LinkedPicture",
    13: "Picture",
    14: "Placeholder",
    16: "Media",
    19: "Table",
    24: "SmartArt",
}


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
def _find_placeholder_by_type(slide, placeholder_type: int):
    """Find the first placeholder of a given type on a slide.

    Returns Shape COM object, or None if not found.
    """
    phs = slide.Shapes.Placeholders
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        if ph.PlaceholderFormat.Type == placeholder_type:
            return ph
    return None


def _resolve_placeholder(slide, placeholder_index=None, placeholder_type=None):
    """Resolve a placeholder by index (int) or type name (str) or type int.

    Args:
        slide: Slide COM object
        placeholder_index: 1-based placeholder index
        placeholder_type: PpPlaceholderType int or type name string

    Returns:
        Shape COM object

    Raises:
        ValueError: If placeholder not found or neither parameter specified
    """
    if placeholder_index is not None:
        return slide.Shapes.Placeholders(placeholder_index)
    elif placeholder_type is not None:
        # Convert string type to int if needed
        if isinstance(placeholder_type, str):
            type_int = PLACEHOLDER_TYPE_MAP.get(placeholder_type.lower())
            if type_int is None:
                raise ValueError(
                    f"Unknown placeholder type '{placeholder_type}'. "
                    f"Valid types: {list(PLACEHOLDER_TYPE_MAP.keys())}"
                )
        else:
            type_int = placeholder_type

        ph = _find_placeholder_by_type(slide, type_int)

        if ph is None:
            type_name = PLACEHOLDER_TYPE_NAMES.get(type_int, str(type_int))
            raise ValueError(
                f"No placeholder of type '{type_name}' found on slide"
            )
        return ph
    else:
        raise ValueError("Must specify either placeholder_index or placeholder_type")


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class ListPlaceholdersInput(BaseModel):
    """Input for listing placeholders on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")


class GetPlaceholderInput(BaseModel):
    """Input for getting placeholder content."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    placeholder_index: Optional[int] = Field(
        default=None, description="1-based placeholder index"
    )
    placeholder_type: Optional[Union[str, int]] = Field(
        default=None,
        description="Placeholder type: name ('title', 'body', 'subtitle') or PpPlaceholderType int"
    )


class SetPlaceholderTextInput(BaseModel):
    """Input for setting placeholder text."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., description="1-based slide index")
    placeholder_index: Optional[int] = Field(
        default=None, description="1-based placeholder index"
    )
    placeholder_type: Optional[Union[str, int]] = Field(
        default=None,
        description="Placeholder type: name ('title', 'body', 'subtitle') or PpPlaceholderType int"
    )
    text: str = Field(..., description="Text content. Use \\n for paragraph breaks.")


class ListDesignsInput(BaseModel):
    """Input for listing all designs (slide masters) in the presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    include_layouts: bool = Field(
        default=False,
        description="If true, include layout names for each design. Default: false (compact).",
    )


class ListLayoutsInput(BaseModel):
    """Input for listing slide layouts."""
    model_config = ConfigDict(str_strip_whitespace=True)

    design_index: Optional[int] = Field(
        default=1, description="1-based design (master) index. Default: 1"
    )


class GetSlideMasterInfoInput(BaseModel):
    """Input for getting slide master information."""
    model_config = ConfigDict(str_strip_whitespace=True)

    design_index: Optional[int] = Field(
        default=1, description="1-based design index. Default: 1"
    )


# ---------------------------------------------------------------------------
# COM implementation functions
# ---------------------------------------------------------------------------
def _list_placeholders_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)
    phs = slide.Shapes.Placeholders

    placeholders = []
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        pf = ph.PlaceholderFormat

        info = {
            "index": i,
            "name": ph.Name,
            "type": pf.Type,
            "type_name": PLACEHOLDER_TYPE_NAMES.get(pf.Type, f"Unknown({pf.Type})"),
            "has_text": False,
            "text_preview": None,
            "left": round(ph.Left, 1),
            "top": round(ph.Top, 1),
            "width": round(ph.Width, 1),
            "height": round(ph.Height, 1),
        }

        if ph.HasTextFrame:
            tf = ph.TextFrame
            has_text = bool(tf.HasText)
            info["has_text"] = has_text
            if has_text:
                text = tf.TextRange.Text
                info["text_preview"] = text[:100] + ("..." if len(text) > 100 else "")

        placeholders.append(info)

    return {
        "status": "success",
        "slide_index": slide_index,
        "placeholder_count": phs.Count,
        "placeholders": placeholders,
    }


def _get_placeholder_impl(slide_index, placeholder_index, placeholder_type) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    ph = _resolve_placeholder(slide, placeholder_index, placeholder_type)
    pf = ph.PlaceholderFormat
    resolved_index = next(
        i
        for i in range(1, slide.Shapes.Placeholders.Count + 1)
        if slide.Shapes.Placeholders(i).Id == ph.Id
    )

    info = {
        "index": resolved_index,
        "name": ph.Name,
        "type": pf.Type,
        "type_name": PLACEHOLDER_TYPE_NAMES.get(pf.Type, f"Unknown({pf.Type})"),
        "contained_type": pf.ContainedType,
        "contained_type_name": CONTAINED_TYPE_NAMES.get(
            pf.ContainedType, f"Unknown({pf.ContainedType})"
        ),
        "left": round(ph.Left, 1),
        "top": round(ph.Top, 1),
        "width": round(ph.Width, 1),
        "height": round(ph.Height, 1),
        "rotation": ph.Rotation,
        "has_text_frame": bool(ph.HasTextFrame),
    }

    if ph.HasTextFrame:
        tf = ph.TextFrame
        tr = tf.TextRange
        info["text"] = tr.Text if tf.HasText else None
        info["paragraph_count"] = tr.Paragraphs().Count
        if tf.HasText:
            info["font_name"] = tr.Font.Name
            info["font_name_far_east"] = tr.Font.NameFarEast
            try:
                info["font_size"] = tr.Font.Size
            except Exception:
                info["font_size"] = None
            info["alignment"] = tr.ParagraphFormat.Alignment

    return {
        "status": "success",
        "placeholder": info,
    }


def _set_placeholder_text_impl(slide_index, placeholder_index, placeholder_type, text) -> dict:
    app, pres, slide = get_slide(slide_index)

    ph = _resolve_placeholder(slide, placeholder_index, placeholder_type)

    if not ph.HasTextFrame:
        raise ValueError(
            f"Placeholder '{ph.Name}' does not have a text frame "
            f"(type={ph.PlaceholderFormat.Type})"
        )

    text = text.replace('\n', '\r')
    ph.TextFrame.TextRange.Text = text

    return {
        "status": "success",
        "slide_index": slide_index,
        "placeholder_name": ph.Name,
        "placeholder_type": ph.PlaceholderFormat.Type,
        "text_length": ph.TextFrame.TextRange.Length,
    }


def _list_layouts_impl(design_index) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    design = pres.Designs(design_index)
    master = design.SlideMaster
    layouts_col = master.CustomLayouts

    layouts = []
    for i in range(1, layouts_col.Count + 1):
        layout = layouts_col(i)
        phs = layout.Shapes.Placeholders

        ph_list = []
        for j in range(1, phs.Count + 1):
            ph = phs(j)
            pf = ph.PlaceholderFormat
            ph_list.append({
                "type": pf.Type,
                "type_name": PLACEHOLDER_TYPE_NAMES.get(pf.Type, f"Unknown({pf.Type})"),
                "name": ph.Name,
            })

        layouts.append({
            "index": i,
            "name": layout.Name,
            "placeholder_count": phs.Count,
            "placeholders": ph_list,
        })

    return {
        "status": "success",
        "design_name": design.Name,
        "layout_count": layouts_col.Count,
        "layouts": layouts,
    }


def _list_designs_impl(include_layouts: bool = False) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    designs = pres.Designs
    result = []
    for d in range(1, designs.Count + 1):
        design = designs(d)
        master = design.SlideMaster
        entry = {
            "design_index": d,
            "design_name": design.Name,
            "layout_count": master.CustomLayouts.Count,
        }
        if include_layouts:
            layout_names = []
            for i in range(1, master.CustomLayouts.Count + 1):
                layout_names.append(master.CustomLayouts(i).Name)
            entry["layout_names"] = layout_names
        result.append(entry)
    return {
        "status": "success",
        "design_count": designs.Count,
        "designs": result,
    }


def _get_slide_master_info_impl(design_index) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    design = pres.Designs(design_index)
    master = design.SlideMaster

    # Get theme colors
    colors = []
    try:
        tcs = master.Theme.ThemeColorScheme
        theme_color_names = [
            "Dark1", "Light1", "Dark2", "Light2",
            "Accent1", "Accent2", "Accent3", "Accent4",
            "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink",
        ]
        for i in range(1, min(tcs.Count + 1, 13)):
            color_val = tcs(i).RGB
            hex_color = int_to_hex(color_val)
            colors.append({
                "index": i,
                "name": theme_color_names[i - 1] if i <= 12 else f"Color{i}",
                "color_hex": hex_color,
            })
    except Exception:
        pass

    return {
        "status": "success",
        "design_name": design.Name,
        "master_name": master.Name,
        "layout_count": master.CustomLayouts.Count,
        "has_title_master": bool(design.HasTitleMaster),
        "theme_colors": colors,
    }


# ---------------------------------------------------------------------------
# MCP tool functions
# ---------------------------------------------------------------------------
def list_placeholders(params: ListPlaceholdersInput) -> str:
    """List all placeholders on a slide."""
    return execute_json(None, _list_placeholders_impl, params.slide_index)


def get_placeholder(params: GetPlaceholderInput) -> str:
    """Get placeholder content and formatting details."""
    return execute_json(None,
            _get_placeholder_impl,
            params.slide_index, params.placeholder_index, params.placeholder_type,
        )


def set_placeholder_text(params: SetPlaceholderTextInput) -> str:
    """Set text in a placeholder."""
    return execute_json(None,
            _set_placeholder_text_impl,
            params.slide_index, params.placeholder_index,
            params.placeholder_type, params.text,
        )


def list_designs(params: ListDesignsInput) -> str:
    """List all designs (slide masters) in the presentation."""
    return execute_json(None, _list_designs_impl, params.include_layouts)


def list_layouts(params: ListLayoutsInput) -> str:
    """List available slide layouts."""
    return execute_json(None, _list_layouts_impl, params.design_index)


def get_slide_master_info(params: GetSlideMasterInfoInput) -> str:
    """Get slide master information including theme colors."""
    return execute_json(None, _get_slide_master_info_impl, params.design_index)

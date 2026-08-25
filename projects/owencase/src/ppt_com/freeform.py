"""Freeform shape creation and node editing for PowerPoint COM automation.

Provides tools for building freeform (path) shapes using FreeformBuilder
and editing their nodes via the ShapeNodes COM API.
"""

import json
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict, model_validator

from utils.com_wrapper import ppt
from utils.com_objects import get_slide
from ppt_com.constants import (
    msoFreeform,
    msoSegmentLine, msoEditingAuto,
    EDITING_TYPE_MAP, EDITING_TYPE_NAMES,
    SEGMENT_TYPE_MAP, SEGMENT_TYPE_NAMES,
)
from utils.com_objects import get_shape as _get_shape



# ---------------------------------------------------------------------------
# Input models
# ---------------------------------------------------------------------------

class NodeSpec(BaseModel):
    """A single path segment node for freeform building."""
    model_config = ConfigDict(str_strip_whitespace=True)

    segment_type: str = Field(
        ...,
        description="Segment type: 'line' or 'curve'.",
    )
    editing_type: str = Field(
        default="auto",
        description=(
            "Node editing type: 'auto', 'corner', 'smooth', or 'symmetric'. "
            "For line segments, always forced to 'auto'. "
            "For curve with 'auto': provide x1, y1 (endpoint). "
            "For curve with 'corner': provide x1, y1 (control point 1), "
            "x2, y2 (control point 2), x3, y3 (endpoint)."
        ),
    )
    x1: float = Field(..., description="X coordinate in points (endpoint for line/auto, first control point for corner curve).")
    y1: float = Field(..., description="Y coordinate in points.")
    x2: Optional[float] = Field(default=None, description="Second control point X (corner curve only).")
    y2: Optional[float] = Field(default=None, description="Second control point Y (corner curve only).")
    x3: Optional[float] = Field(default=None, description="Endpoint X (corner curve only).")
    y3: Optional[float] = Field(default=None, description="Endpoint Y (corner curve only).")

    @model_validator(mode="after")
    def validate_node(self):
        seg = self.segment_type.lower()
        et = self.editing_type.lower()
        if seg not in SEGMENT_TYPE_MAP:
            raise ValueError(f"segment_type must be 'line' or 'curve', got '{seg}'")
        # Line segments must use auto
        if seg == "line":
            self.editing_type = "auto"
            return self
        # Curve: FreeformBuilder.AddNodes only supports 'auto' and 'corner'
        # ('smooth'/'symmetric' are for editing existing nodes via SetEditingType)
        if et not in ("auto", "corner"):
            raise ValueError(
                f"curve editing_type for new freeforms must be 'auto' or 'corner', got '{et}'. "
                "Use ppt_set_node_editing_type to apply 'smooth' or 'symmetric' after creation."
            )
        # Corner curve requires all 6 extra coordinates
        if et == "corner":
            missing = [n for n, v in [("x2", self.x2), ("y2", self.y2), ("x3", self.x3), ("y3", self.y3)] if v is None]
            if missing:
                raise ValueError(
                    f"curve with editing_type='corner' requires x2, y2, x3, y3. Missing: {missing}"
                )
        return self


class BuildFreeformInput(BaseModel):
    """Input for creating a new freeform shape from path segments."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index.")
    start_x: float = Field(..., description="X position of the first node in points.")
    start_y: float = Field(..., description="Y position of the first node in points.")
    start_editing_type: str = Field(
        default="corner",
        description=(
            "Editing type of the first node: 'auto' or 'corner'. "
            "('smooth'/'symmetric' are not supported by BuildFreeform; use "
            "ppt_set_node_editing_type after creation.)"
        ),
    )

    @model_validator(mode="after")
    def validate_start_editing_type(self):
        et = self.start_editing_type.lower()
        if et not in ("auto", "corner"):
            raise ValueError(
                f"start_editing_type must be 'auto' or 'corner', got '{et}'. "
                "Use ppt_set_node_editing_type to apply 'smooth' or 'symmetric' after creation."
            )
        self.start_editing_type = et
        return self
    nodes: list[NodeSpec] = Field(
        ...,
        min_length=1,
        description="List of path segments. At least one node is required.",
    )
    close_path: bool = Field(
        default=False,
        description="If true, automatically adds a straight line back to the start point to close the path.",
    )
    shape_name: Optional[str] = Field(default=None, description="Optional name for the created shape.")


class ShapeNodeInput(BaseModel):
    """Common input for tools that identify a shape and a node index."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index.")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred).")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index.")
    node_index: int = Field(..., ge=1, description="1-based node index.")


class GetShapeNodesInput(BaseModel):
    """Input for reading nodes of an existing freeform shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index.")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred).")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index.")


class SetNodePositionInput(BaseModel):
    """Input for moving a freeform node to new coordinates."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index.")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred).")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index.")
    node_index: int = Field(..., ge=1, description="1-based node index.")
    x: float = Field(..., description="New X position in points (relative to slide upper-left).")
    y: float = Field(..., description="New Y position in points (relative to slide upper-left).")


class InsertNodeInput(BaseModel):
    """Input for inserting a new node after an existing node."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index.")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred).")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index.")
    after_index: int = Field(..., ge=1, description="Insert new node after this 1-based index.")
    segment_type: str = Field(default="line", description="Segment type: 'line' or 'curve'.")
    editing_type: str = Field(
        default="auto",
        description=(
            "Node editing type: 'auto' or 'corner'. "
            "For line segments, always forced to 'auto'."
        ),
    )
    x1: float = Field(..., description="X coordinate in points (endpoint for line/auto, first control point for corner).")
    y1: float = Field(..., description="Y coordinate in points.")
    x2: Optional[float] = Field(default=None, description="Second control point X (corner curve only).")
    y2: Optional[float] = Field(default=None, description="Second control point Y (corner curve only).")
    x3: Optional[float] = Field(default=None, description="Endpoint X (corner curve only).")
    y3: Optional[float] = Field(default=None, description="Endpoint Y (corner curve only).")

    @model_validator(mode="after")
    def validate_insert_node(self):
        seg = self.segment_type.lower()
        et = self.editing_type.lower()
        if seg not in SEGMENT_TYPE_MAP:
            raise ValueError(f"segment_type must be 'line' or 'curve', got '{seg}'")
        if seg == "line":
            self.editing_type = "auto"
            return self
        # Curve: ShapeNodes.Insert only supports 'auto' and 'corner'
        if et not in ("auto", "corner"):
            raise ValueError(
                f"curve editing_type must be 'auto' or 'corner', got '{et}'. "
                "Use ppt_set_node_editing_type to apply 'smooth' or 'symmetric' after insertion."
            )
        if et == "corner":
            missing = [n for n, v in [("x2", self.x2), ("y2", self.y2), ("x3", self.x3), ("y3", self.y3)] if v is None]
            if missing:
                raise ValueError(
                    f"curve with editing_type='corner' requires x2, y2, x3, y3. Missing: {missing}"
                )
        return self


class DeleteNodeInput(BaseModel):
    """Input for deleting a node from a freeform shape."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index.")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred).")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index.")
    node_index: int = Field(..., ge=1, description="1-based node index to delete.")


class SetNodeEditingTypeInput(BaseModel):
    """Input for changing a node's editing type."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index.")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred).")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index.")
    node_index: int = Field(..., ge=1, description="1-based node index.")
    editing_type: str = Field(
        ...,
        description="New editing type: 'auto', 'corner', 'smooth', or 'symmetric'.",
    )


class SetSegmentTypeInput(BaseModel):
    """Input for changing the segment type after a node."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index.")
    shape_name: Optional[str] = Field(default=None, description="Shape name (preferred).")
    shape_index: Optional[int] = Field(default=None, ge=1, description="1-based shape index.")
    node_index: int = Field(..., ge=1, description="1-based node index. The segment AFTER this node is changed.")
    segment_type: str = Field(
        ...,
        description="New segment type: 'line' or 'curve'.",
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _check_freeform(shape):
    """Raise ValueError if shape is not a freeform."""
    if shape.Type != msoFreeform:
        raise ValueError(
            f"Shape '{shape.Name}' is not a freeform (type={shape.Type}). "
            "Only freeform shapes (type=5) support node operations."
        )


def _read_nodes(shape) -> list[dict]:
    """Read all nodes from a freeform shape into a list of dicts.

    Uses shape.Vertices for XY positions (node.Points is unreliable in pywin32)
    and shape.Nodes.Item(i) for editing/segment type metadata.

    Closing nodes (added when close_path=True) and Bézier control points
    created with msoEditingCorner have position data but inaccessible COM
    metadata; they are returned with segment_type/editing_type = "inaccessible".
    """
    nodes_com = shape.Nodes
    # shape.Vertices: 2D Variant array, 0-based in pywin32: verts[i][0]=X, verts[i][1]=Y
    vertices = shape.Vertices

    result = []
    for i in range(1, nodes_com.Count + 1):
        vx = round(float(vertices[i - 1][0]), 2)
        vy = round(float(vertices[i - 1][1]), 2)
        try:
            node = nodes_com.Item(i)
            et_int = node.EditingType
            seg_int = node.SegmentType
            result.append({
                "index": i,
                "x": vx,
                "y": vy,
                "editing_type": EDITING_TYPE_NAMES.get(et_int, str(et_int)),
                "segment_type": SEGMENT_TYPE_NAMES.get(seg_int, str(seg_int)),
            })
        except Exception:
            # Some nodes (closing nodes, Bézier control points created with
            # msoEditingCorner) have position data but inaccessible COM metadata.
            result.append({
                "index": i,
                "x": vx,
                "y": vy,
                "editing_type": "inaccessible",
                "segment_type": "inaccessible",
                "note": "Metadata not accessible via COM (control point or closing node).",
            })
    return result


# ---------------------------------------------------------------------------
# COM implementation functions (run on STA thread via ppt.execute)
# ---------------------------------------------------------------------------

def _build_freeform_impl(slide_index, start_et_int, start_x, start_y, nodes_data, close_path, shape_name):
    app, pres, slide = get_slide(slide_index)

    builder = slide.Shapes.BuildFreeform(start_et_int, start_x, start_y)

    for nd in nodes_data:
        seg_int = nd["seg_int"]
        et_int = nd["et_int"]
        x1, y1 = nd["x1"], nd["y1"]

        if seg_int == msoSegmentLine:
            # Line segment: EditingType must be msoEditingAuto, 4-arg form
            builder.AddNodes(msoSegmentLine, msoEditingAuto, x1, y1)
        elif et_int == msoEditingAuto:
            # Curve + auto: 4-arg form
            builder.AddNodes(seg_int, msoEditingAuto, x1, y1)
        else:
            # Curve + corner: 8-arg form with 2 control points + endpoint
            builder.AddNodes(seg_int, et_int, x1, y1, nd["x2"], nd["y2"], nd["x3"], nd["y3"])

    if close_path:
        # Close the path with a straight line back to the start
        builder.AddNodes(msoSegmentLine, msoEditingAuto, start_x, start_y)

    shape = builder.ConvertToShape()
    if shape_name:
        shape.Name = shape_name

    return json.dumps({
        "success": True,
        "shape_name": shape.Name,
        "shape_index": shape.ZOrderPosition,
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
    })


def _get_shape_nodes_impl(slide_index, shape_name, shape_index):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    _check_freeform(shape)

    nodes = _read_nodes(shape)

    return json.dumps({
        "shape_name": shape.Name,
        "node_count": len(nodes),
        "nodes": nodes,
    })


def _set_node_position_impl(slide_index, shape_name, shape_index, node_index, x, y):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    _check_freeform(shape)

    nodes_com = shape.Nodes
    if node_index > nodes_com.Count:
        raise ValueError(f"node_index {node_index} out of range (shape has {nodes_com.Count} nodes).")

    nodes_com.SetPosition(node_index, x, y)

    # Re-read actual position via Vertices (node.Points is unreliable in pywin32)
    verts = shape.Vertices
    return json.dumps({
        "success": True,
        "shape_name": shape.Name,
        "node_index": node_index,
        "x": round(float(verts[node_index - 1][0]), 2),
        "y": round(float(verts[node_index - 1][1]), 2),
    })


def _insert_node_impl(slide_index, shape_name, shape_index, after_index, seg_int, et_int, x1, y1, x2, y2, x3, y3):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    _check_freeform(shape)

    nodes_com = shape.Nodes
    if after_index > nodes_com.Count:
        raise ValueError(f"after_index {after_index} out of range (shape has {nodes_com.Count} nodes).")

    if seg_int == msoSegmentLine:
        # Line: force msoEditingAuto, 5-arg form
        nodes_com.Insert(after_index, msoSegmentLine, msoEditingAuto, x1, y1)
    elif et_int == msoEditingAuto:
        # Curve + auto: 5-arg form
        nodes_com.Insert(after_index, seg_int, msoEditingAuto, x1, y1)
    else:
        # Curve + corner: 8-arg form
        nodes_com.Insert(after_index, seg_int, et_int, x1, y1, x2, y2, x3, y3)

    return json.dumps({
        "success": True,
        "shape_name": shape.Name,
        "new_node_count": shape.Nodes.Count,
    })


def _delete_node_impl(slide_index, shape_name, shape_index, node_index):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    _check_freeform(shape)

    nodes_com = shape.Nodes
    if node_index > nodes_com.Count:
        raise ValueError(f"node_index {node_index} out of range (shape has {nodes_com.Count} nodes).")

    nodes_com.Delete(node_index)
    return json.dumps({
        "success": True,
        "shape_name": shape.Name,
        "remaining_node_count": shape.Nodes.Count,
    })


def _set_node_editing_type_impl(slide_index, shape_name, shape_index, node_index, et_int):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    _check_freeform(shape)

    nodes_com = shape.Nodes
    if node_index > nodes_com.Count:
        raise ValueError(f"node_index {node_index} out of range (shape has {nodes_com.Count} nodes).")

    nodes_com.SetEditingType(node_index, et_int)
    try:
        new_et = nodes_com.Item(node_index).EditingType
        editing_type_str = EDITING_TYPE_NAMES.get(new_et, str(new_et))
        result = {
            "success": True,
            "shape_name": shape.Name,
            "node_index": node_index,
            "editing_type": editing_type_str,
        }
    except Exception:
        # Control-point and closing nodes have inaccessible COM metadata,
        # consistent with _read_nodes. The SetEditingType call itself succeeded.
        result = {
            "success": True,
            "shape_name": shape.Name,
            "node_index": node_index,
            "editing_type": "inaccessible",
            "note": "Metadata not accessible via COM (control point or closing node).",
        }
    return json.dumps(result)


def _set_segment_type_impl(slide_index, shape_name, shape_index, node_index, seg_int):
    app, pres, slide = get_slide(slide_index)
    shape = _get_shape(slide, None, shape_name=shape_name, shape_index=shape_index)
    _check_freeform(shape)

    nodes_com = shape.Nodes
    if node_index > nodes_com.Count:
        raise ValueError(f"node_index {node_index} out of range (shape has {nodes_com.Count} nodes).")

    old_count = nodes_com.Count
    nodes_com.SetSegmentType(node_index, seg_int)
    new_count = shape.Nodes.Count

    result = {
        "success": True,
        "shape_name": shape.Name,
        "node_index": node_index,
        "segment_type": SEGMENT_TYPE_NAMES.get(seg_int, str(seg_int)),
        "old_node_count": old_count,
        "new_node_count": new_count,
    }
    if old_count != new_count:
        result["note"] = (
            "Node count changed — switching line↔curve adds or removes control-point nodes. "
            "Re-call ppt_get_shape_nodes to see updated indices."
        )
    return json.dumps(result)


def build_freeform(params: BuildFreeformInput) -> str:
    """Create a freeform path from line and Bezier segments."""
    nodes_data = [
        {
            "seg_int": SEGMENT_TYPE_MAP[node.segment_type.lower()],
            "et_int": EDITING_TYPE_MAP[node.editing_type.lower()],
            "x1": node.x1,
            "y1": node.y1,
            "x2": node.x2,
            "y2": node.y2,
            "x3": node.x3,
            "y3": node.y3,
        }
        for node in params.nodes
    ]
    return ppt.execute(
        _build_freeform_impl,
        params.slide_index,
        EDITING_TYPE_MAP[params.start_editing_type.lower()],
        params.start_x,
        params.start_y,
        nodes_data,
        params.close_path,
        params.shape_name,
    )


def get_shape_nodes(params: GetShapeNodesInput) -> str:
    """Read the nodes of a freeform shape."""
    return ppt.execute(
        _get_shape_nodes_impl,
        params.slide_index,
        params.shape_name,
        params.shape_index,
    )


def set_node_position(params: SetNodePositionInput) -> str:
    """Move a freeform node to an absolute slide position."""
    return ppt.execute(
        _set_node_position_impl,
        params.slide_index,
        params.shape_name,
        params.shape_index,
        params.node_index,
        params.x,
        params.y,
    )


def insert_node(params: InsertNodeInput) -> str:
    """Insert a node after an existing freeform node."""
    return ppt.execute(
        _insert_node_impl,
        params.slide_index,
        params.shape_name,
        params.shape_index,
        params.after_index,
        SEGMENT_TYPE_MAP[params.segment_type.lower()],
        EDITING_TYPE_MAP[params.editing_type.lower()],
        params.x1,
        params.y1,
        params.x2,
        params.y2,
        params.x3,
        params.y3,
    )


def delete_node(params: DeleteNodeInput) -> str:
    """Delete a freeform node and its following segment."""
    return ppt.execute(
        _delete_node_impl,
        params.slide_index,
        params.shape_name,
        params.shape_index,
        params.node_index,
    )


def set_node_editing_type(params: SetNodeEditingTypeInput) -> str:
    """Set a freeform node's editing type."""
    return ppt.execute(
        _set_node_editing_type_impl,
        params.slide_index,
        params.shape_name,
        params.shape_index,
        params.node_index,
        EDITING_TYPE_MAP[params.editing_type.lower()],
    )


def set_segment_type(params: SetSegmentTypeInput) -> str:
    """Set the segment type after a freeform node."""
    return ppt.execute(
        _set_segment_type_impl,
        params.slide_index,
        params.shape_name,
        params.shape_index,
        params.node_index,
        SEGMENT_TYPE_MAP[params.segment_type.lower()],
    )

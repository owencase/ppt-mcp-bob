"""Shared PowerPoint COM object lookup helpers."""

from __future__ import annotations

from typing import Any

from utils.com_wrapper import ppt
from utils.navigation import goto_slide


def get_slide(slide_index: int, *, navigate: bool = True) -> tuple[Any, Any, Any]:
    """Return the app, active presentation, and a 1-based slide."""
    app = ppt._get_app_impl()
    presentation = ppt._get_pres_impl()
    if not 1 <= slide_index <= presentation.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{presentation.Slides.Count})"
        )
    if navigate:
        goto_slide(app, slide_index)
    return app, presentation, presentation.Slides(slide_index)


def get_shape(
    slide: Any,
    name_or_index: str | int | None = None,
    shape_name: str | None = None,
    shape_index: int | None = None,
    shape_id: int | None = None,
) -> Any:
    """Find a shape by stable ID, name, or 1-based index."""
    if shape_id is not None:
        for index in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(index)
            if int(shape.Id) == shape_id:
                return shape
        raise ValueError(f"Shape ID {shape_id} not found on this slide.")

    identifier = (
        shape_name
        if shape_name is not None
        else shape_index
        if shape_index is not None
        else name_or_index
    )
    if identifier is None:
        raise ValueError("One of shape_id, shape_name, or shape_index must be provided.")

    if isinstance(identifier, int):
        if not 1 <= identifier <= slide.Shapes.Count:
            raise ValueError(
                f"Shape index {identifier} out of range (1-{slide.Shapes.Count})"
            )
        return slide.Shapes(identifier)

    for index in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(index)
        if shape.Name == identifier:
            return shape
    raise ValueError(f"Shape '{identifier}' not found on slide")


def get_typed_shape(slide: Any, identifier: str | int, property_name: str, kind: str) -> Any:
    """Find a shape and require a truthy COM capability property."""
    shape = get_shape(slide, identifier)
    if not getattr(shape, property_name):
        raise ValueError(f"Shape '{shape.Name}' is not a {kind}")
    return shape

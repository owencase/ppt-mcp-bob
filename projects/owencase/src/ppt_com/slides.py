"""Slide-level operations for PowerPoint COM automation.

Add, delete, duplicate, move, list, and query slides. Manage speaker notes.
"""

from typing import NamedTuple, Optional

from pydantic import BaseModel, Field, ConfigDict, model_validator

from utils.color import hex_to_int
from utils.com_wrapper import ppt
from utils.tool_result import execute_json
from utils.navigation import goto_slide as nav_goto_slide
from ppt_com.constants import ppLayoutBlank, msoTrue, msoFalse


# Friendly layout name -> PpSlideLayout constant
LAYOUT_NAME_MAP = {
    "title": 1,
    "text": 2,
    "two_column_text": 3,
    "table": 4,
    "title_only": 11,
    "blank": 12,
    "section_header": 33,
    "comparison": 34,
    "content_with_caption": 35,
    "picture_with_caption": 36,
}


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddSlideInput(BaseModel):
    """Input for adding a new slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    position: Optional[int] = Field(
        default=None,
        description=(
            "1-based position for the new slide. "
            "If omitted, the slide is added at the end."
        ),
    )
    layout: Optional[int] = Field(
        default=None,
        description=(
            "PpSlideLayout constant: 1=Title, 2=Text, 11=TitleOnly, 12=Blank, "
            "33=SectionHeader, 34=Comparison, 35=ContentWithCaption, "
            "36=PictureWithCaption. Ignored if layout_name is provided."
        ),
    )
    layout_name: Optional[str] = Field(
        default=None,
        description=(
            "Layout name to look up from the slide master's custom layouts. "
            "Alternatively use a friendly name: 'title', 'text', 'blank', "
            "'title_only', 'section_header', 'comparison', "
            "'content_with_caption', 'picture_with_caption'. "
            "Overrides the layout parameter."
        ),
    )
    design_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based design (slide master) index to search for the layout_name. "
            "If omitted, searches the active presentation's default master first, "
            "then all designs."
        ),
    )
    like_slide_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of an existing slide to copy the design + layout from. "
            "Inherits that slide's exact CustomLayout (design/master AND layout), "
            "so the new slide looks like it without any layout-name ambiguity. "
            "This is the most reliable way to 'add a slide like this one'. "
            "Takes precedence over layout, layout_name, and design_index. "
            "Note: this inherits the look only — it does NOT copy the slide's "
            "content/shapes (use ppt_duplicate_slide for a full copy)."
        ),
    )
    count: int = Field(
        default=1,
        ge=1,
        description="Number of slides to add. All slides use the same layout. "
        "When count > 1, returns a list of created slide indices.",
    )


class DeleteSlideInput(BaseModel):
    """Input for deleting one or more slides.

    Provide exactly ONE of:
      - slide_index (single slide), or
      - slide_indices (an explicit list), or
      - from_index + to_index (an inclusive range).
    """
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: Optional[int] = Field(
        default=None,
        description="1-based index of a single slide to delete.",
    )
    slide_indices: Optional[list[int]] = Field(
        default=None,
        description=(
            "List of 1-based slide indices to delete in one call. "
            "Deleted from highest index first internally, so the indices you "
            "pass all refer to the CURRENT numbering (no manual bookkeeping)."
        ),
    )
    from_index: Optional[int] = Field(
        default=None,
        description=(
            "Start of an inclusive 1-based range to delete (used with to_index)."
        ),
    )
    to_index: Optional[int] = Field(
        default=None,
        description=(
            "End of an inclusive 1-based range to delete (used with from_index)."
        ),
    )

    @model_validator(mode="after")
    def _check_exactly_one_form(self):
        single = self.slide_index is not None
        listed = self.slide_indices is not None
        ranged = self.from_index is not None or self.to_index is not None

        forms = [single, listed, ranged]
        if sum(forms) == 0:
            raise ValueError(
                "Provide one of: slide_index, slide_indices, or "
                "from_index+to_index"
            )
        if sum(forms) > 1:
            raise ValueError(
                "Provide only one of: slide_index, slide_indices, or "
                "from_index+to_index (not a combination)"
            )
        if listed:
            if len(self.slide_indices) == 0:
                raise ValueError("slide_indices must not be empty")
            if any(i < 1 for i in self.slide_indices):
                raise ValueError("slide_indices must all be >= 1")
        if single and self.slide_index < 1:
            raise ValueError("slide_index must be >= 1")
        if ranged:
            if self.from_index is None or self.to_index is None:
                raise ValueError(
                    "Both from_index and to_index are required for a range"
                )
            if self.from_index < 1:
                raise ValueError("from_index must be >= 1")
            if self.from_index > self.to_index:
                raise ValueError(
                    f"from_index ({self.from_index}) must be <= to_index "
                    f"({self.to_index})"
                )
        return self


class DuplicateSlideInput(BaseModel):
    """Input for duplicating a slide (optionally to a target position, N times)."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide to duplicate.",
    )
    insert_at: Optional[int] = Field(
        default=None,
        description=(
            "1-based target position for the copies. "
            "If omitted, copies are inserted immediately after the source "
            "(the historical behavior). Use -1 to append at the end. "
            "Collapses the old duplicate+move dance into one call."
        ),
    )
    count: int = Field(
        default=1,
        ge=1,
        description=(
            "Number of copies to make. When > 1, copies are placed "
            "consecutively starting at the target position."
        ),
    )

    @model_validator(mode="after")
    def _check_insert_at(self):
        if self.insert_at is not None and self.insert_at != -1 and self.insert_at < 1:
            raise ValueError(
                "insert_at must be a positive 1-based position or -1 (end)"
            )
        return self


class MoveSlideInput(BaseModel):
    """Input for moving one or more slides.

    Provide exactly ONE of slide_index (single) or slide_indices (a block).
    """
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: Optional[int] = Field(
        default=None,
        description="Current 1-based index of a single slide to move.",
    )
    slide_indices: Optional[list[int]] = Field(
        default=None,
        description=(
            "List of current 1-based slide indices to move as a block. "
            "Their relative order is preserved; they end up contiguous "
            "starting at new_position. Indices refer to the CURRENT numbering."
        ),
    )
    new_position: int = Field(
        ...,
        description=(
            "Target 1-based position. For a single slide it is the destination "
            "index; for a block it is the index of the first moved slide."
        ),
    )

    @model_validator(mode="after")
    def _check_exactly_one_form(self):
        single = self.slide_index is not None
        listed = self.slide_indices is not None
        if single == listed:  # both set or both unset
            raise ValueError(
                "Provide exactly one of slide_index or slide_indices"
            )
        if listed:
            if len(self.slide_indices) == 0:
                raise ValueError("slide_indices must not be empty")
            if any(i < 1 for i in self.slide_indices):
                raise ValueError("slide_indices must all be >= 1")
        if single and self.slide_index < 1:
            raise ValueError("slide_index must be >= 1")
        if self.new_position < 1:
            raise ValueError("new_position must be >= 1")
        return self


class CopySlideInput(BaseModel):
    """Input for copying slides, optionally to another open presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: Optional[int] = Field(
        default=None,
        description="1-based index of a single source slide to copy.",
    )
    slide_indices: Optional[list[int]] = Field(
        default=None,
        description=(
            "List of 1-based source slide indices to copy in one call. "
            "Copied in the given order."
        ),
    )
    source_presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the source presentation. "
            "If omitted, uses the locked session target."
        ),
    )
    source_presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "File name of the source presentation (e.g. 'Template.pptx'). "
            "Alternative to source_presentation_index."
        ),
    )
    to_presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the destination presentation. "
            "If omitted, copies into the active presentation."
        ),
    )
    to_presentation_name: Optional[str] = Field(
        default=None,
        description=(
            "File name of the destination presentation (e.g. 'Deck.pptx'). "
            "Alternative to to_presentation_index."
        ),
    )
    insert_at: Optional[int] = Field(
        default=None,
        description=(
            "1-based position in the destination where the copies are inserted. "
            "If omitted or -1, copies are appended at the end."
        ),
    )

    @model_validator(mode="after")
    def _check_source(self):
        single = self.slide_index is not None
        listed = self.slide_indices is not None
        if single == listed:  # both set or both unset
            raise ValueError(
                "Provide exactly one of slide_index or slide_indices"
            )
        if listed:
            if len(self.slide_indices) == 0:
                raise ValueError("slide_indices must not be empty")
            if any(i < 1 for i in self.slide_indices):
                raise ValueError("slide_indices must all be >= 1")
        if single and self.slide_index < 1:
            raise ValueError("slide_index must be >= 1")
        if self.insert_at is not None and self.insert_at != -1 and self.insert_at < 1:
            raise ValueError(
                "insert_at must be a positive 1-based position or -1 (end)"
            )
        if (
            self.source_presentation_index is not None
            and self.source_presentation_name is not None
        ):
            raise ValueError(
                "Specify either source_presentation_index or "
                "source_presentation_name, not both"
            )
        if (
            self.to_presentation_index is not None
            and self.to_presentation_name is not None
        ):
            raise ValueError(
                "Specify either to_presentation_index or "
                "to_presentation_name, not both"
            )
        return self


class ListSlidesInput(BaseModel):
    """Input for listing slides (no required params)."""
    model_config = ConfigDict(str_strip_whitespace=True)

    presentation_index: Optional[int] = Field(
        default=None,
        description=(
            "1-based index of the presentation to list slides from. "
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


class GetSlideInfoInput(BaseModel):
    """Input for getting detailed slide info."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide to query.",
    )


class SetSlideNotesInput(BaseModel):
    """Input for setting speaker notes."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide.",
    )
    notes_text: Optional[str] = Field(
        default=None,
        description="The speaker notes text to set. If omitted, only formatting is applied to existing notes.",
    )
    font_name: Optional[str] = Field(
        default=None,
        description="Latin font name (e.g. 'Arial'). Also sets the East Asian font unless font_name_fareast is provided.",
    )
    font_name_fareast: Optional[str] = Field(
        default=None,
        description="East Asian (CJK) font name (e.g. 'BIZ UDPゴシック'). Overrides the Far East font independently of font_name.",
    )
    font_size: Optional[float] = Field(
        default=None,
        description="Font size in points. Note: formatting affects printed notes and PDF export only — the Notes pane and Presenter View ignore font size (Presenter View has its own A+/A- zoom).",
    )
    bold: Optional[bool] = Field(
        default=None,
        description="Bold on/off.",
    )
    italic: Optional[bool] = Field(
        default=None,
        description="Italic on/off.",
    )
    color: Optional[str] = Field(
        default=None,
        description="Font color as '#RRGGBB' hex string.",
    )


class GetSlideNotesInput(BaseModel):
    """Input for getting speaker notes."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide.",
    )


# ---------------------------------------------------------------------------
# Helper to resolve a presentation
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
class _LayoutMatch(NamedTuple):
    """One design that contains a custom layout with the requested name."""
    design_index: int
    design_name: str
    custom_layout: object


def _find_layout_matches(pres, layout_name: str, design_index: Optional[int]):
    """Find custom layouts matching ``layout_name`` across designs.

    Returns a list of _LayoutMatch (design_index, design_name, custom_layout) —
    one per design that contains a layout with that exact name. When
    design_index is given, only that design is searched.
    """
    if design_index is not None:
        search = [design_index]
    else:
        search = range(1, pres.Designs.Count + 1)

    matches = []
    for d in search:
        design = pres.Designs(d)
        master = design.SlideMaster
        for i in range(1, master.CustomLayouts.Count + 1):
            lay = master.CustomLayouts(i)
            if lay.Name == layout_name:
                matches.append(_LayoutMatch(d, design.Name, lay))
                break  # at most one layout of a given name per design
    return matches


def _add_slide_impl(
    position: Optional[int],
    layout: Optional[int],
    layout_name: Optional[str],
    design_index: Optional[int] = None,
    count: int = 1,
    like_slide_index: Optional[int] = None,
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if position is None:
        position = pres.Slides.Count + 1

    if position < 1 or position > pres.Slides.Count + 1:
        raise ValueError(
            f"Position {position} out of range (1-{pres.Slides.Count + 1})"
        )

    # Resolve layout once before the loop
    use_custom_layout = False
    custom_layout = None
    friendly_layout = None
    layout_val = None
    resolved_layout_name = None
    resolved_design_name = None
    layout_ambiguous = False
    ambiguous_designs = None

    if like_slide_index is not None:
        # Highest precedence: inherit the exact CustomLayout (design + layout)
        # of an existing slide. Using the object reference directly means there
        # is no layout-name ambiguity across designs.
        if like_slide_index < 1 or like_slide_index > pres.Slides.Count:
            raise ValueError(
                f"like_slide_index {like_slide_index} out of range "
                f"(1-{pres.Slides.Count})"
            )
        src_slide = pres.Slides(like_slide_index)
        custom_layout = src_slide.CustomLayout
        use_custom_layout = True
        resolved_layout_name = custom_layout.Name
        try:
            resolved_design_name = src_slide.Design.Name
        except Exception:
            resolved_design_name = None
    elif layout_name:
        # Check friendly name map first
        friendly_key = layout_name.lower().strip().replace(" ", "_")
        if friendly_key in LAYOUT_NAME_MAP:
            friendly_layout = LAYOUT_NAME_MAP[friendly_key]
        else:
            if design_index is not None and (
                design_index < 1 or design_index > pres.Designs.Count
            ):
                raise ValueError(
                    f"Design index {design_index} out of range "
                    f"(1-{pres.Designs.Count})"
                )

            matches = _find_layout_matches(pres, layout_name, design_index)

            if not matches:
                # Collect available layouts for the error message, scoped to
                # the same designs that were actually searched — otherwise a
                # design_index lookup would list layouts from unrelated designs
                # (including the one where the name does exist), contradicting
                # the "not found" message.
                error_search = (
                    [design_index] if design_index is not None
                    else range(1, pres.Designs.Count + 1)
                )
                available = {}
                for d in error_search:
                    design = pres.Designs(d)
                    m = design.SlideMaster
                    names = []
                    for i in range(1, m.CustomLayouts.Count + 1):
                        names.append(m.CustomLayouts(i).Name)
                    available[design.Name] = names
                raise ValueError(
                    f"Layout '{layout_name}' not found. "
                    f"Available custom layouts by design: {available}"
                )

            # Prefer the default master's design when the name is ambiguous,
            # matching the historical selection order (default master first).
            # Tiebreak is by master name; in the rare case two designs share a
            # master name the pick may differ, but the ambiguity warning below
            # lists every candidate so the caller can override with
            # design_index or like_slide_index.
            chosen = matches[0]
            if len(matches) > 1:
                try:
                    default_master_name = pres.SlideMaster.Name
                    for mt in matches:
                        if pres.Designs(mt.design_index).SlideMaster.Name == default_master_name:
                            chosen = mt
                            break
                except Exception:
                    pass
                layout_ambiguous = True
                ambiguous_designs = [m.design_name for m in matches]

            custom_layout = chosen.custom_layout
            resolved_design_name = chosen.design_name
            use_custom_layout = True
            resolved_layout_name = custom_layout.Name
    else:
        layout_val = layout if layout is not None else ppLayoutBlank

    # Add slides in a loop
    created_slides = []
    for i in range(count):
        insert_pos = position + i
        if use_custom_layout:
            slide = pres.Slides.AddSlide(Index=insert_pos, pCustomLayout=custom_layout)
        elif friendly_layout is not None:
            slide = pres.Slides.Add(Index=insert_pos, Layout=friendly_layout)
        else:
            slide = pres.Slides.Add(Index=insert_pos, Layout=layout_val)
        created_slides.append({
            "slide_index": slide.SlideIndex,
            "slide_id": slide.SlideID,
        })

    # Navigate to the last created slide
    nav_goto_slide(app, created_slides[-1]["slide_index"])

    # Read actual layout from the first created slide
    first_slide = pres.Slides(created_slides[0]["slide_index"])

    # Resolve the actually-applied layout/design names for caller verification.
    final_layout_name = resolved_layout_name
    final_design_name = resolved_design_name
    try:
        if final_layout_name is None:
            final_layout_name = first_slide.CustomLayout.Name
        if final_design_name is None:
            final_design_name = first_slide.Design.Name
    except Exception:
        pass

    result = {
        "success": True,
        "slides_created": len(created_slides),
        "slides": created_slides,
        # "layout" is always the PpSlideLayout integer constant; the
        # human-readable custom layout name is in "layout_name".
        "layout": first_slide.Layout,
        "layout_name": final_layout_name,
        "design_name": final_design_name,
    }

    # Surface ambiguous layout-name matches so callers can detect a possible
    # wrong-design selection (the issue's core failure mode).
    if layout_ambiguous:
        result["layout_ambiguous"] = True
        result["warning"] = (
            f"Layout '{layout_name}' exists in multiple designs "
            f"({', '.join(ambiguous_designs)}); used design "
            f"'{final_design_name}'. Pass design_index or like_slide_index "
            f"to select a specific one."
        )

    # Backward compatibility: include top-level slide_index/slide_id for count=1
    if count == 1:
        result["slide_index"] = created_slides[0]["slide_index"]
        result["slide_id"] = created_slides[0]["slide_id"]

    return result


def _delete_slide_impl(
    slide_index: Optional[int] = None,
    slide_indices: Optional[list] = None,
    from_index: Optional[int] = None,
    to_index: Optional[int] = None,
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)
    total = pres.Slides.Count

    # Normalize the three input forms into a sorted, de-duplicated index set.
    if slide_index is not None:
        targets = [slide_index]
    elif slide_indices is not None:
        targets = list(slide_indices)
    else:
        targets = list(range(from_index, to_index + 1))

    targets = sorted(set(targets))
    if not targets:
        raise ValueError("No slides to delete")

    out_of_range = [i for i in targets if i < 1 or i > total]
    if out_of_range:
        raise ValueError(
            f"Slide index(es) {out_of_range} out of range (1-{total})"
        )
    if len(targets) >= total:
        raise ValueError(
            "Cannot delete every slide — a presentation must keep at least "
            f"one slide (requested {len(targets)} of {total})"
        )

    # Navigate to the lowest target before deleting so the editor follows.
    nav_goto_slide(app, targets[0])

    # Delete from the highest index first so earlier indices stay valid.
    for idx in sorted(targets, reverse=True):
        pres.Slides(idx).Delete()

    result = {
        "success": True,
        "deleted_indices": targets,
        "deleted_count": len(targets),
        "remaining_count": pres.Slides.Count,
    }
    # Backward compatibility for single-slide callers.
    if slide_index is not None:
        result["deleted_index"] = slide_index
    return result


def _duplicate_slide_impl(
    slide_index: int,
    insert_at: Optional[int] = None,
    count: int = 1,
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )

    # Validate an explicit insert_at up front (raise rather than silently clamp,
    # consistent with ppt_copy_slide). Up to one past the end is allowed.
    if insert_at is not None and insert_at != -1 and insert_at > pres.Slides.Count + 1:
        raise ValueError(
            f"insert_at {insert_at} out of range "
            f"(1-{pres.Slides.Count + 1})"
        )

    nav_goto_slide(app, slide_index)

    # Track the source by SlideID — once we start moving copies around, its
    # index can shift (e.g. inserting a copy before it), so a fixed index is
    # unsafe across iterations.
    src_id = pres.Slides(slide_index).SlideID

    new_indices = []
    new_ids = []
    for i in range(count):
        cur_src_idx = pres.Slides.FindBySlideID(src_id).SlideIndex
        new_slide = pres.Slides(cur_src_idx).Duplicate()(1)
        new_id = new_slide.SlideID

        # Resolve the desired FINAL 1-based position of this copy. insert_at was
        # validated above, so each target lands in a valid range as the deck
        # grows; min() is a defensive cap, not a silent correctness clamp.
        if insert_at is None:
            target = cur_src_idx + 1 + i          # right after the source, in order
        elif insert_at == -1:
            target = pres.Slides.Count             # append (count includes the copy)
        else:
            target = min(insert_at + i, pres.Slides.Count)

        if new_slide.SlideIndex != target:
            new_slide.MoveTo(target)

        new_indices.append(new_slide.SlideIndex)
        new_ids.append(new_id)

    nav_goto_slide(app, new_indices[-1])

    result = {
        "success": True,
        "count": len(new_indices),
        "new_slide_indices": new_indices,
        "new_slide_ids": new_ids,
    }
    # Backward compatibility for single-copy callers.
    if count == 1:
        result["new_slide_index"] = new_indices[0]
        result["new_slide_id"] = new_ids[0]
    return result


def _compute_final_order(all_ids: list, src_ids: list, new_position: int) -> list:
    """Compute the desired final order of slide IDs for a bulk move.

    ``all_ids`` is every slide ID in current order; ``src_ids`` is the selected
    IDs in their original relative order. The selected slides occupy positions
    ``new_position .. new_position + len(src_ids) - 1`` and the unselected slides
    keep their relative order around them. Pure function (no COM) so the
    order-building logic is unit-testable without PowerPoint.
    """
    selected = set(src_ids)
    others = [sid for sid in all_ids if sid not in selected]
    return others[: new_position - 1] + list(src_ids) + others[new_position - 1:]


def _move_slide_impl(
    new_position: int,
    slide_index: Optional[int] = None,
    slide_indices: Optional[list] = None,
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)
    total = pres.Slides.Count

    # Normalize to a sorted, de-duplicated list of source indices.
    sources = [slide_index] if slide_index is not None else list(slide_indices)
    sources = sorted(set(sources))

    out_of_range = [i for i in sources if i < 1 or i > total]
    if out_of_range:
        raise ValueError(
            f"Slide index(es) {out_of_range} out of range (1-{total})"
        )

    k = len(sources)
    max_start = total - k + 1
    if new_position < 1 or new_position > max_start:
        raise ValueError(
            f"new_position {new_position} out of range (1-{max_start}) "
            f"for a block of {k} slide(s)"
        )

    nav_goto_slide(app, sources[0])

    # Build the FULL desired final order of slide IDs, then realize it. A
    # direction heuristic based on sources[0] is not enough: a non-contiguous
    # selection that straddles the target can leave a member outside the block
    # (e.g. [1,4,5] -> position 2). Instead, lay out the final order explicitly
    # via _compute_final_order. Read every ID once (total COM calls) and derive
    # src_ids from that list rather than re-crossing the COM boundary per slide.
    all_ids = [pres.Slides(i).SlideID for i in range(1, total + 1)]
    src_ids = [all_ids[i - 1] for i in sources]
    final_ids = _compute_final_order(all_ids, src_ids, new_position)

    # Realize the target order left-to-right. Invariant: once we reach position
    # f, positions 1..f-1 already hold their final slide, so the slide wanted at
    # f is necessarily at some index >= f. MoveTo(f) then shifts only positions
    # >= f, never disturbing what is already placed.
    for f in range(1, total + 1):
        cur_idx = pres.Slides.FindBySlideID(final_ids[f - 1]).SlideIndex
        if cur_idx != f:
            pres.Slides(cur_idx).MoveTo(toPos=f)

    result = {
        "success": True,
        "moved_slide_ids": src_ids,
        "moved_count": k,
        "new_start_position": new_position,
    }
    # Backward compatibility for single-slide callers.
    if slide_index is not None:
        result["moved_from"] = slide_index
        result["moved_to"] = new_position
    return result


def _copy_slide_impl(
    slide_index: Optional[int],
    slide_indices: Optional[list],
    source_presentation_index: Optional[int],
    source_presentation_name: Optional[str],
    to_presentation_index: Optional[int],
    to_presentation_name: Optional[str],
    insert_at: Optional[int],
) -> dict:
    app = ppt._get_app_impl()

    src_pres = _resolve_presentation(
        app,
        presentation_index=source_presentation_index,
        presentation_name=source_presentation_name,
    )
    dst_pres = _resolve_presentation(
        app,
        presentation_index=to_presentation_index,
        presentation_name=to_presentation_name,
    )

    sources = [slide_index] if slide_index is not None else list(slide_indices)

    src_total = src_pres.Slides.Count
    out_of_range = [i for i in sources if i < 1 or i > src_total]
    if out_of_range:
        raise ValueError(
            f"Source slide index(es) {out_of_range} out of range "
            f"(1-{src_total})"
        )

    append = insert_at is None or insert_at == -1
    if not append:
        # Allow inserting anywhere from the front up to one past the end.
        if insert_at < 1 or insert_at > dst_pres.Slides.Count + 1:
            raise ValueError(
                f"insert_at {insert_at} out of range "
                f"(1-{dst_pres.Slides.Count + 1})"
            )

    # Capture source SlideIDs up front. For a same-presentation copy each paste
    # shifts the numbering, so a fixed src_i would select the wrong slide on
    # later iterations — resolve by SlideID each time instead.
    src_ids = [src_pres.Slides(i).SlideID for i in sources]

    # Copy each source slide via the clipboard, which preserves formatting and
    # carries the source design across presentations.
    new_indices = []
    for j, sid in enumerate(src_ids):
        src_pres.Slides.FindBySlideID(sid).Copy()
        # Paste(Index) inserts before slide Index; there is no slide one past
        # the end to paste before, so fall back to append in that case.
        if append or (insert_at + j) > dst_pres.Slides.Count:
            rng = dst_pres.Slides.Paste()
        else:
            rng = dst_pres.Slides.Paste(insert_at + j)
        new_indices.append(rng(1).SlideIndex)

    # Navigate the destination's window to the last pasted slide.
    try:
        nav_goto_slide(app, new_indices[-1])
    except Exception:
        pass

    return {
        "success": True,
        "copied_count": len(new_indices),
        "new_slide_indices": new_indices,
        "target_presentation": dst_pres.Name,
        "source_presentation": src_pres.Name,
    }


def _list_slides_impl(
    presentation_index: Optional[int],
    presentation_name: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(
        app, presentation_index=presentation_index, presentation_name=presentation_name
    )

    slides = []
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)

        layout_name = ""
        try:
            layout_name = slide.CustomLayout.Name
        except Exception:
            pass

        has_notes = False
        try:
            notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            has_notes = len(notes_text.strip()) > 0
        except Exception:
            pass

        slides.append({
            "index": slide.SlideIndex,
            "slide_id": slide.SlideID,
            "name": slide.Name,
            "layout": slide.Layout,
            "layout_name": layout_name,
            "hidden": bool(slide.SlideShowTransition.Hidden),
            "shapes_count": slide.Shapes.Count,
            "has_notes": has_notes,
        })

    return {"slides_count": pres.Slides.Count, "slides": slides}


def _get_slide_info_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )

    slide = pres.Slides(slide_index)
    trans = slide.SlideShowTransition

    layout_name = ""
    try:
        layout_name = slide.CustomLayout.Name
    except Exception:
        pass

    title_text = ""
    has_title = bool(slide.Shapes.HasTitle)
    if has_title:
        try:
            title_text = slide.Shapes.Title.TextFrame.TextRange.Text
        except Exception:
            pass

    notes_text = ""
    try:
        notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
    except Exception:
        pass

    design_name = ""
    try:
        design_name = slide.Design.Name
    except Exception:
        pass

    return {
        "index": slide.SlideIndex,
        "slide_id": slide.SlideID,
        "slide_number": slide.SlideNumber,
        "name": slide.Name,
        "layout": slide.Layout,
        "layout_name": layout_name,
        "hidden": bool(trans.Hidden),
        "shapes_count": slide.Shapes.Count,
        "has_title": has_title,
        "title_text": title_text,
        "notes_text": notes_text,
        "follow_master_background": bool(slide.FollowMasterBackground),
        "transition_effect": trans.EntryEffect,
        "advance_on_click": bool(trans.AdvanceOnClick),
        "advance_on_time": bool(trans.AdvanceOnTime),
        "advance_time": trans.AdvanceTime,
        "design_name": design_name,
    }


def _set_slide_notes_impl(
    slide_index: int,
    notes_text: Optional[str],
    font_name: Optional[str],
    font_name_fareast: Optional[str],
    font_size: Optional[float],
    bold: Optional[bool],
    italic: Optional[bool],
    color: Optional[str],
) -> dict:
    app = ppt._get_app_impl()
    nav_goto_slide(app, slide_index)
    pres = _resolve_presentation(app)

    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )

    slide = pres.Slides(slide_index)
    text_range = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange

    if notes_text is not None:
        text_range.Text = notes_text

    # Apply formatting to entire notes text range
    font = text_range.Font
    if font_name is not None:
        font.Name = font_name
        if font_name_fareast is None:
            font.NameFarEast = font_name
    if font_name_fareast is not None:
        font.NameFarEast = font_name_fareast
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = msoTrue if bold else msoFalse
    if italic is not None:
        font.Italic = msoTrue if italic else msoFalse
    if color is not None:
        font.Color.RGB = hex_to_int(color)

    return {"success": True}


def _get_slide_notes_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = _resolve_presentation(app)

    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )

    slide = pres.Slides(slide_index)
    try:
        notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
    except Exception:
        notes_text = ""

    return {"slide_index": slide_index, "notes_text": notes_text}


# ---------------------------------------------------------------------------
# MCP tool functions (return JSON strings)
# ---------------------------------------------------------------------------
def add_slide(params: AddSlideInput) -> str:
    """Add a new slide to the active presentation."""
    return execute_json(None,
            _add_slide_impl, params.position, params.layout,
            params.layout_name, params.design_index, params.count,
            params.like_slide_index,
        )


def delete_slide(params: DeleteSlideInput) -> str:
    """Delete one or more slides."""
    return execute_json(None,
            _delete_slide_impl,
            params.slide_index,
            params.slide_indices,
            params.from_index,
            params.to_index,
        )


def duplicate_slide(params: DuplicateSlideInput) -> str:
    """Duplicate a slide."""
    return execute_json(None,
            _duplicate_slide_impl,
            params.slide_index,
            params.insert_at,
            params.count,
        )


def move_slide(params: MoveSlideInput) -> str:
    """Move one or more slides to a new position."""
    return execute_json(None,
            _move_slide_impl,
            params.new_position,
            params.slide_index,
            params.slide_indices,
        )


def copy_slide(params: CopySlideInput) -> str:
    """Copy slides, optionally into another open presentation."""
    return execute_json(None,
            _copy_slide_impl,
            params.slide_index,
            params.slide_indices,
            params.source_presentation_index,
            params.source_presentation_name,
            params.to_presentation_index,
            params.to_presentation_name,
            params.insert_at,
        )


def list_slides(params: ListSlidesInput) -> str:
    """List all slides in the active presentation."""
    return execute_json(None,
            _list_slides_impl,
            params.presentation_index,
            params.presentation_name,
        )


def get_slide_info(params: GetSlideInfoInput) -> str:
    """Get detailed info about a specific slide."""
    return execute_json(None, _get_slide_info_impl, params.slide_index)


def set_slide_notes(params: SetSlideNotesInput) -> str:
    """Set speaker notes for a slide."""
    return execute_json(None,
            _set_slide_notes_impl,
            params.slide_index,
            params.notes_text,
            params.font_name,
            params.font_name_fareast,
            params.font_size,
            params.bold,
            params.italic,
            params.color,
        )


def get_slide_notes(params: GetSlideNotesInput) -> str:
    """Get speaker notes for a slide."""
    return execute_json(None, _get_slide_notes_impl, params.slide_index)


class GotoSlideInput(BaseModel):
    """Input for navigating to a specific slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(
        ...,
        description="1-based index of the slide to navigate to.",
        ge=1,
    )


def _goto_slide_impl(slide_index: int) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(
            f"Slide index {slide_index} out of range (1-{pres.Slides.Count})"
        )
    app.ActiveWindow.View.GotoSlide(slide_index)
    return {
        "success": True,
        "active_slide_index": slide_index,
    }


def goto_slide(params: GotoSlideInput) -> str:
    """Navigate the active window to display a specific slide."""
    return execute_json(None, _goto_slide_impl, params.slide_index)

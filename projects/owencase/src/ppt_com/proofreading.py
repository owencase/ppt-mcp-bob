"""Read-only text proofreading for PowerPoint presentations.

PowerPoint exposes proofing language tags but does not expose a programmable
collection of spelling errors.  This module therefore combines deterministic
checks with a complete, location-aware text payload that the MCP client can
review contextually before the presentation is saved.
"""

from __future__ import annotations

import json
import re
from collections import Counter
from typing import Any, Optional

from pydantic import BaseModel, ConfigDict, Field, field_validator

from ppt_com.constants import (
    msoChart,
    msoGroup,
    msoSmartArt,
    msoTable,
    xlCategory,
    xlValue,
)
from utils.com_wrapper import ppt


_PROOFREAD_BATCH_SIZE = 15

# Deliberately small, high-confidence seed list.  It is not presented as a
# complete Korean dictionary: callers can extend it with custom_replacements
# and the MCP client must still review every returned text unit contextually.
COMMON_KOREAN_TYPOS: dict[str, str] = {
    "프리젠테이션": "프레젠테이션",
    "프리젠테이숀": "프레젠테이션",
    "프레젠테이숀": "프레젠테이션",
    "메세지": "메시지",
    "데이타": "데이터",
    "컨텐츠": "콘텐츠",
    "플렛폼": "플랫폼",
    "파트너쉽": "파트너십",
    "라이센스": "라이선스",
}

COMMON_ENGLISH_TYPOS: dict[str, str] = {
    "teh": "the",
    "recieve": "receive",
    "seperate": "separate",
    "occured": "occurred",
    "succesful": "successful",
    "enviroment": "environment",
}


class ProofreadTextInput(BaseModel):
    """Options for deterministic checks and contextual text review."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    slide_indices: Optional[list[int]] = Field(
        default=None,
        description="1-based slide indices. Omit to inspect every slide.",
    )
    include_notes: bool = Field(
        default=False,
        description="Include speaker notes in proofreading text units.",
    )
    include_text_units: bool = Field(
        default=True,
        description=(
            "Return every location-aware text unit so the MCP client can perform "
            "contextual spelling, spacing, and terminology review."
        ),
    )
    check_common_typos: bool = True
    check_repeated_words: bool = True
    check_punctuation: bool = True
    check_brackets: bool = True
    check_suspicious_characters: bool = True
    custom_replacements: dict[str, str] = Field(
        default_factory=dict,
        description="Additional exact typo-to-correction mappings.",
    )
    allowed_terms: list[str] = Field(
        default_factory=list,
        description="Terms that must not be reported by the built-in typo list.",
    )
    max_findings: int = Field(default=200, ge=1, le=1000)

    @field_validator("slide_indices")
    @classmethod
    def validate_slide_indices(cls, value):
        if value is None:
            return value
        if not value:
            raise ValueError("slide_indices must not be empty")
        if any(index < 1 for index in value):
            raise ValueError("slide_indices values must all be >= 1")
        if len(value) != len(set(value)):
            raise ValueError("slide_indices must not contain duplicates")
        return value

    @field_validator("custom_replacements")
    @classmethod
    def validate_custom_replacements(cls, value):
        for typo, correction in value.items():
            if not typo.strip() or not correction.strip():
                raise ValueError("custom_replacements keys and values must not be empty")
            if typo == correction:
                raise ValueError("custom replacement must change the original text")
        return value

    @field_validator("allowed_terms")
    @classmethod
    def validate_allowed_terms(cls, value):
        if any(not term.strip() for term in value):
            raise ValueError("allowed_terms values must not be empty")
        if len({term.casefold() for term in value}) != len(value):
            raise ValueError("allowed_terms must not contain duplicates")
        return value


def _normalise_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")


def _shape_identity(shape) -> dict[str, Any]:
    identity: dict[str, Any] = {}
    try:
        identity["shape_id"] = int(shape.Id)
    except Exception:
        pass
    try:
        identity["shape_name"] = str(shape.Name)
    except Exception:
        pass
    return identity


def _append_text_unit(
    units: list[dict[str, Any]],
    *,
    slide_index: int,
    container_type: str,
    text: Any,
    **location: Any,
) -> None:
    normalised = _normalise_text(text)
    if not normalised.strip():
        return
    units.append(
        {
            "unit_id": f"slide-{slide_index}:unit-{len(units) + 1}",
            "slide_index": slide_index,
            "container_type": container_type,
            **location,
            "text": normalised,
        }
    )


def _collect_chart_text(shape, slide_index: int, units: list[dict[str, Any]]) -> None:
    identity = _shape_identity(shape)
    try:
        chart = shape.Chart
    except Exception:
        return

    try:
        if chart.HasTitle:
            _append_text_unit(
                units,
                slide_index=slide_index,
                container_type="chart_title",
                text=chart.ChartTitle.Text,
                **identity,
            )
    except Exception:
        pass

    for axis_type, axis_name in ((xlCategory, "category"), (xlValue, "value")):
        try:
            axis = chart.Axes(axis_type)
            if axis.HasTitle:
                _append_text_unit(
                    units,
                    slide_index=slide_index,
                    container_type="chart_axis_title",
                    text=axis.AxisTitle.Text,
                    axis=axis_name,
                    **identity,
                )
        except Exception:
            pass

    try:
        series_count = int(chart.SeriesCollection().Count)
    except Exception:
        series_count = 0
    categories_added = False
    for series_index in range(1, series_count + 1):
        try:
            series = chart.SeriesCollection(series_index)
        except Exception:
            continue
        try:
            _append_text_unit(
                units,
                slide_index=slide_index,
                container_type="chart_series_name",
                text=series.Name,
                series_index=series_index,
                **identity,
            )
        except Exception:
            pass
        if categories_added:
            continue
        try:
            categories = list(series.XValues)
        except Exception:
            categories = []
        for category_index, category in enumerate(categories, 1):
            if isinstance(category, str) and category.strip():
                _append_text_unit(
                    units,
                    slide_index=slide_index,
                    container_type="chart_category",
                    text=category,
                    category_index=category_index,
                    **identity,
                )
        categories_added = bool(categories)


def _collect_shape_text(
    shape,
    slide_index: int,
    units: list[dict[str, Any]],
    group_path: Optional[list[str]] = None,
) -> None:
    identity = _shape_identity(shape)
    path_fields: dict[str, Any] = {}
    if group_path:
        path_fields["group_path"] = group_path

    try:
        shape_type = int(shape.Type)
    except Exception:
        shape_type = None

    if shape_type == msoGroup:
        name = identity.get("shape_name", "group")
        next_path = [*(group_path or []), name]
        try:
            for item_index in range(1, int(shape.GroupItems.Count) + 1):
                _collect_shape_text(
                    shape.GroupItems(item_index), slide_index, units, next_path
                )
        except Exception:
            pass
        return

    if shape_type == msoTable:
        try:
            table = shape.Table
            for row in range(1, int(table.Rows.Count) + 1):
                for column in range(1, int(table.Columns.Count) + 1):
                    text = table.Cell(row, column).Shape.TextFrame.TextRange.Text
                    _append_text_unit(
                        units,
                        slide_index=slide_index,
                        container_type="table_cell",
                        text=text,
                        row=row,
                        column=column,
                        **identity,
                        **path_fields,
                    )
        except Exception:
            pass
        return

    if shape_type == msoSmartArt:
        try:
            smart_art = shape.SmartArt
            for node_index in range(1, int(smart_art.AllNodes.Count) + 1):
                node = smart_art.AllNodes(node_index)
                _append_text_unit(
                    units,
                    slide_index=slide_index,
                    container_type="smartart_node",
                    text=node.TextFrame2.TextRange.Text,
                    node_index=node_index,
                    **identity,
                    **path_fields,
                )
        except Exception:
            pass
        return

    if shape_type == msoChart:
        _collect_chart_text(shape, slide_index, units)
        return

    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            _append_text_unit(
                units,
                slide_index=slide_index,
                container_type="shape_text",
                text=shape.TextFrame.TextRange.Text,
                **identity,
                **path_fields,
            )
    except Exception:
        pass


def _collect_text_units_impl(
    slide_indices: list[int], include_notes: bool
) -> list[dict[str, Any]]:
    pres = ppt._get_pres_impl()
    units: list[dict[str, Any]] = []
    for slide_index in slide_indices:
        slide = pres.Slides(slide_index)
        for shape_index in range(1, int(slide.Shapes.Count) + 1):
            _collect_shape_text(slide.Shapes(shape_index), slide_index, units)

        if include_notes:
            try:
                notes_text = (
                    slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
                )
                _append_text_unit(
                    units,
                    slide_index=slide_index,
                    container_type="speaker_notes",
                    text=notes_text,
                )
            except Exception:
                pass
    return units


def _line_column(text: str, start: int) -> tuple[int, int]:
    line = text.count("\n", 0, start) + 1
    last_break = text.rfind("\n", 0, start)
    column = start + 1 if last_break < 0 else start - last_break
    return line, column


def _context(text: str, start: int, length: int, radius: int = 30) -> str:
    before = text[max(0, start - radius):start]
    match = text[start:start + length]
    after = text[start + length:start + length + radius]
    return f"{before}[{match}]{after}".replace("\n", " ")


def _finding(
    unit: dict[str, Any],
    *,
    code: str,
    severity: str,
    start: int,
    length: int,
    message: str,
    suggestion: Optional[str] = None,
    source: Optional[str] = None,
) -> dict[str, Any]:
    line, column = _line_column(unit["text"], start)
    result = {
        "code": code,
        "severity": severity,
        "unit_id": unit["unit_id"],
        "slide_index": unit["slide_index"],
        "container_type": unit["container_type"],
        "start": start,
        "length": length,
        "line": line,
        "column": column,
        "original": unit["text"][start:start + length],
        "context": _context(unit["text"], start, length),
        "message": message,
    }
    for key in (
        "shape_id",
        "shape_name",
        "group_path",
        "row",
        "column",
        "node_index",
        "axis",
        "series_index",
        "category_index",
    ):
        if key in unit:
            # Text position column is kept as text_column when a table column
            # is also present.
            if key == "column" and unit["container_type"] == "table_cell":
                result["text_column"] = result.pop("column")
                result["column"] = unit[key]
            else:
                result[key] = unit[key]
    if suggestion is not None:
        result["suggestion"] = suggestion
    if source is not None:
        result["source"] = source
    return result


def _term_pattern(term: str, *, custom: bool) -> re.Pattern[str]:
    flags = 0
    if term.isascii():
        flags = re.IGNORECASE
        return re.compile(
            rf"(?<![0-9A-Za-z_]){re.escape(term)}(?![0-9A-Za-z_])",
            flags,
        )
    # Custom replacements intentionally use exact substring matching. Korean
    # particles and compounds often attach directly to the target term.
    return re.compile(re.escape(term), flags)


def _check_brackets(unit: dict[str, Any]) -> list[dict[str, Any]]:
    text = unit["text"]
    pairs = {"(": ")", "[": "]", "{": "}", "“": "”", "‘": "’", "「": "」", "『": "』", "《": "》"}
    closers = {value: key for key, value in pairs.items()}
    stack: list[tuple[str, int]] = []
    findings: list[dict[str, Any]] = []
    for index, character in enumerate(text):
        if character in pairs:
            stack.append((character, index))
        elif character in closers:
            if not stack or stack[-1][0] != closers[character]:
                findings.append(
                    _finding(
                        unit,
                        code="UNMATCHED_BRACKET",
                        severity="warning",
                        start=index,
                        length=1,
                        message=f"Closing bracket {character!r} has no matching opener.",
                    )
                )
            else:
                stack.pop()
    for character, index in stack:
        findings.append(
            _finding(
                unit,
                code="UNMATCHED_BRACKET",
                severity="warning",
                start=index,
                length=1,
                message=f"Opening bracket {character!r} has no matching closer.",
                suggestion=pairs[character],
            )
        )
    return findings


def _analyse_text_units(
    units: list[dict[str, Any]], params: ProofreadTextInput
) -> tuple[list[dict[str, Any]], int]:
    findings: list[dict[str, Any]] = []
    allowed = {term.casefold() for term in params.allowed_terms}

    replacements: list[tuple[str, str, str, bool]] = []
    if params.check_common_typos:
        for typo, correction in {**COMMON_KOREAN_TYPOS, **COMMON_ENGLISH_TYPOS}.items():
            replacements.append((typo, correction, "built_in", False))
    for typo, correction in params.custom_replacements.items():
        replacements = [item for item in replacements if item[0].casefold() != typo.casefold()]
        replacements.append((typo, correction, "custom", True))
    replacements.sort(key=lambda item: len(item[0]), reverse=True)

    for unit in units:
        text = unit["text"]

        for typo, correction, source, custom in replacements:
            if typo.casefold() in allowed:
                continue
            for match in _term_pattern(typo, custom=custom).finditer(text):
                findings.append(
                    _finding(
                        unit,
                        code="SUSPECTED_TYPO",
                        severity="error",
                        start=match.start(),
                        length=match.end() - match.start(),
                        message=f"Replace {match.group(0)!r} with {correction!r}.",
                        suggestion=correction,
                        source=source,
                    )
                )

        if params.check_repeated_words:
            repeated_word = re.compile(
                r"(?<![0-9A-Za-z가-힣_])([0-9A-Za-z가-힣_]+)\s+\1(?![0-9A-Za-z가-힣_])",
                re.IGNORECASE,
            )
            for match in repeated_word.finditer(text):
                findings.append(
                    _finding(
                        unit,
                        code="REPEATED_WORD",
                        severity="error",
                        start=match.start(),
                        length=match.end() - match.start(),
                        message="The same word appears twice in succession.",
                        suggestion=match.group(1),
                    )
                )

        if params.check_punctuation:
            punctuation_patterns = (
                (re.compile(r"[!?;,:]{2,}"), "Repeated punctuation."),
                (re.compile(r"\.{4,}"), "More than three consecutive periods."),
                (re.compile(r"[ \t]+(?=[,.;:!?])"), "Whitespace before punctuation."),
                (re.compile(r"(?<=\S)[ \t]{2,}(?=\S)"), "Repeated internal whitespace."),
            )
            for pattern, message in punctuation_patterns:
                for match in pattern.finditer(text):
                    suggestion = ""
                    if "punctuation" in message.lower() and not match.group(0).isspace():
                        suggestion = match.group(0)[0]
                    elif "whitespace" in message.lower():
                        suggestion = " " if "internal" in message.lower() else ""
                    findings.append(
                        _finding(
                            unit,
                            code="SUSPICIOUS_PUNCTUATION",
                            severity="warning",
                            start=match.start(),
                            length=match.end() - match.start(),
                            message=message,
                            suggestion=suggestion,
                        )
                    )

        if params.check_brackets:
            findings.extend(_check_brackets(unit))

        if params.check_suspicious_characters:
            suspicious_patterns = (
                (re.compile(r"[\uFFFD\u200B\u200C\u200D\uFEFF]"), "Suspicious invisible or replacement character."),
                (re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]"), "Unexpected control character."),
                (re.compile(r"(?:ï¿½|â€™|â€œ|â€|â€“|â€”)", re.IGNORECASE), "Possible mojibake text."),
            )
            for pattern, message in suspicious_patterns:
                for match in pattern.finditer(text):
                    findings.append(
                        _finding(
                            unit,
                            code="SUSPICIOUS_CHARACTER",
                            severity="error",
                            start=match.start(),
                            length=match.end() - match.start(),
                            message=message,
                        )
                    )

    # A single character can participate in more than one rule. Keep distinct
    # codes, but remove exact duplicates from overlapping replacement sources.
    unique: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()
    for item in findings:
        key = (
            item["unit_id"], item["code"], item["start"], item["length"],
            item.get("suggestion"),
        )
        if key not in seen:
            seen.add(key)
            unique.append(item)
    unique.sort(key=lambda item: (item["slide_index"], item["unit_id"], item["start"], item["code"]))
    return unique[:params.max_findings], len(unique)


def _resolve_slide_indices(total_slides: int, requested: Optional[list[int]]) -> list[int]:
    if requested is None:
        return list(range(1, total_slides + 1))
    invalid = [index for index in requested if index > total_slides]
    if invalid:
        raise ValueError(
            f"Slide indices out of range: {invalid}; valid range is 1-{total_slides}"
        )
    return requested


def proofread_text(params: ProofreadTextInput) -> str:
    """Collect presentation text and run deterministic typo checks."""
    try:
        identity = ppt.execute(
            lambda: {
                "presentation_full_name": str(ppt._get_pres_impl().FullName),
                "slide_count": int(ppt._get_pres_impl().Slides.Count),
            }
        )
        indices = _resolve_slide_indices(identity["slide_count"], params.slide_indices)
        units: list[dict[str, Any]] = []
        for offset in range(0, len(indices), _PROOFREAD_BATCH_SIZE):
            batch = indices[offset:offset + _PROOFREAD_BATCH_SIZE]
            units.extend(ppt.execute(_collect_text_units_impl, batch, params.include_notes))

        findings, total_findings = _analyse_text_units(units, params)
        by_code = Counter(item["code"] for item in findings)
        by_container = Counter(unit["container_type"] for unit in units)
        result: dict[str, Any] = {
            "valid": total_findings == 0,
            "semantic_review_required": True,
            "presentation_full_name": identity["presentation_full_name"],
            "checked_slides": indices,
            "checked_text_units": len(units),
            "checked_characters": sum(len(unit["text"]) for unit in units),
            "findings": findings,
            "summary": {
                "finding_count": total_findings,
                "returned_finding_count": len(findings),
                "findings_truncated": total_findings > len(findings),
                "by_code": dict(sorted(by_code.items())),
                "by_container_type": dict(sorted(by_container.items())),
            },
            "review_instructions": (
                "Review every text unit contextually for spelling, Korean spacing, "
                "grammar, terminology, and proper nouns. Deterministic valid=true "
                "does not replace this contextual review. Correct confirmed issues "
                "and rerun ppt_proofread_text before saving."
            ),
        }
        if params.include_text_units:
            result["text_units"] = units
        return json.dumps(result, ensure_ascii=False)
    except Exception as exc:
        return json.dumps({"error": f"Text proofreading failed: {exc}"}, ensure_ascii=False)


def register_tools(mcp):
    """Register the location-aware text proofreading tool."""

    @mcp.tool(
        name="ppt_proofread_text",
        annotations={
            "title": "Proofread Presentation Text",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_ppt_proofread_text(params: ProofreadTextInput) -> str:
        """Check presentation text for likely typos without changing the deck.

        Collects location-aware text from regular shapes, placeholders, grouped
        shapes, table cells, SmartArt nodes, chart titles/axes/series/categories,
        and optionally speaker notes. Runs deterministic checks for common Korean
        and English typos, caller-supplied replacements, repeated words,
        punctuation, unmatched brackets, control characters, and mojibake.

        The result also includes every text unit by default. Review those units
        contextually for Korean spelling and spacing because no deterministic
        dictionary can safely resolve domain terms and proper nouns. A result of
        valid=true means deterministic rules passed, not that contextual review
        may be skipped. Correct confirmed issues and rerun this tool before save.
        """
        return proofread_text(params)

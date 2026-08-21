"""Import smoke tests for the MCP server entry point.

Every other test module imports individual tool modules but never `server.py`
itself, so a dependency that renames or drops a symbol `server.py` imports goes
undetected — mcp 2.0 removing `mcp.server.fastmcp` (issue #176) broke the server
at module load for every user before any test noticed.

Importing the server does not launch PowerPoint: the COM connection is
established lazily on the first tool call (issue #148).
"""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="module")
def server():
    pytest.importorskip("win32com", reason="pywin32 required to import the server")
    import src.server as server_module

    return server_module


def test_server_module_imports(server):
    """server.py loads and exposes a server instance and an entry point."""
    assert server.mcp is not None
    assert callable(server.main)


def test_server_advertises_quiet_three_line_handoff(server):
    """Bob receives the quiet-work and exact three-line handoff policy."""
    instructions = server.mcp.instructions

    assert instructions == server.SERVER_INSTRUCTIONS
    assert "Do not send user-visible plans" in instructions
    assert "Do not send a user-visible response while autonomous work remains" in instructions
    assert "exactly\nthree non-empty lines" in instructions
    assert "run the required diff and\nvalidation first" in instructions


def test_server_isolates_full_resolution_visual_context(server):
    """Inline previews must be kept out of Bob's long-lived main task."""
    instructions = server.mcp.instructions

    assert 'default `delivery="file"`' in instructions
    assert "`spawn_subagent`" in instructions
    assert "`fork_context=false`" in instructions
    assert "exactly one slide index" in instructions
    assert "Run visual-inspection subagents sequentially" in instructions
    assert "Never load multiple inline previews" in instructions


def test_tools_are_registered(server):
    """All tools register successfully under whichever mcp major is installed."""
    tools = asyncio.run(server.mcp.list_tools())
    assert len(tools) > 100
    assert all(t.name.startswith("ppt_") for t in tools)
    assert any(t.name == "ppt_proofread_text" for t in tools)


def test_slide_preview_defaults_to_context_safe_file_delivery(server):
    """Calling the preview tool without an override must not inject an image."""
    tools = asyncio.run(server.mcp.list_tools())
    tool = next(item for item in tools if item.name == "ppt_get_slide_preview")
    schema = tool.inputSchema
    preview_input = schema["$defs"]["GetSlidePreviewInput"]
    delivery = preview_input["properties"]["delivery"]

    assert delivery["default"] == "file"
    assert delivery["enum"] == ["file", "inline"]
    assert "full-resolution PNG" in delivery["description"]


def test_default_preview_result_contains_metadata_not_image_bytes(server, monkeypatch):
    """File delivery must keep ImageContent out of Bob's main transcript."""
    from utils.com_wrapper import ppt

    monkeypatch.setattr(
        ppt,
        "execute",
        lambda *_args: {
            "success": True,
            "slide_index": 2,
            "delivery": "file",
            "path": r"C:\Temp\ppt-mcp-bob\previews\deck\slide-2.png",
            "mime_type": "image/png",
            "bytes": 123456,
            "presentation": r"C:\Decks\deck.pptx",
        },
    )

    result = asyncio.run(
        server.tool_ppt_get_slide_preview(
            server.GetSlidePreviewInput(slide_index=2)
        )
    )
    payload = json.loads(result)

    assert not isinstance(result, server.Image)
    assert payload["delivery"] == "file"
    assert payload["context_safe"] is True
    assert payload["bytes"] == 123456


def test_inline_preview_remains_full_resolution_opt_in(server, monkeypatch):
    """A context-isolated subagent can still request the original PNG bytes."""
    from utils.com_wrapper import ppt

    image_data = b"full-resolution-png-bytes"
    monkeypatch.setattr(
        ppt,
        "execute",
        lambda *_args: {
            "success": True,
            "slide_index": 2,
            "delivery": "inline",
            "path": r"C:\Temp\ppt-mcp-bob\previews\deck\slide-2.png",
            "mime_type": "image/png",
            "bytes": len(image_data),
            "presentation": r"C:\Decks\deck.pptx",
            "image_data": image_data,
        },
    )

    result = asyncio.run(
        server.tool_ppt_get_slide_preview(
            server.GetSlidePreviewInput(slide_index=2, delivery="inline")
        )
    )

    assert isinstance(result, server.Image)
    assert result.data == image_data


def test_proofreading_tool_is_read_only(server):
    """Proofreading must never mutate a presentation."""
    tools = asyncio.run(server.mcp.list_tools())
    tool = next(item for item in tools if item.name == "ppt_proofread_text")
    annotations = tool.annotations.model_dump(by_alias=True, exclude_none=True)

    assert annotations["readOnlyHint"] is True
    assert annotations["destructiveHint"] is False
    assert annotations["idempotentHint"] is True


def test_annotations_serialize_with_protocol_field_names(server):
    """Annotations must reach the wire in camelCase, as MCP specifies.

    Tools declare annotations as plain dicts (`{"readOnlyHint": ...}`). Both mcp
    majors coerce those into a `ToolAnnotations` model whose fields are
    snake_case, so the protocol names survive only via serialization aliases.
    """
    tools = asyncio.run(server.mcp.list_tools())
    annotated = [t for t in tools if t.annotations is not None]
    assert annotated, "expected tools to carry annotations"

    dumped = annotated[0].annotations.model_dump(by_alias=True, exclude_none=True)
    assert "readOnlyHint" in dumped
    assert "read_only_hint" not in dumped

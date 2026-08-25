"""IBM Bob PowerPoint MCP server.

Safe, real-time PowerPoint editing through Windows COM automation.
"""

import logging
import sys
from pathlib import Path
from typing import Any

# When installed via PyPI (entry point: src.server:main), ensure the src/
# directory is in sys.path so that internal imports like
# `from utils.com_wrapper import ppt` resolve correctly.
_src_dir = str(Path(__file__).parent)
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from contextlib import asynccontextmanager

# mcp 2.0 renamed mcp.server.fastmcp to mcp.server.mcpserver, and FastMCP to
# MCPServer, without leaving a compatibility shim behind. Import from whichever
# module the installed major provides so the server runs on both 1.x and 2.x.
#
# Catch ModuleNotFoundError rather than ImportError: the fallback should only
# trigger when the 2.x module is absent. A broken 2.x install that raises
# ImportError from one of its own imports must surface that traceback instead
# of being masked by the 1.x fallback failing afterwards.
try:
    from mcp.server.mcpserver import MCPServer as _MCPServerBase  # mcp >= 2.0
except ModuleNotFoundError:
    from mcp.server.fastmcp import FastMCP as _MCPServerBase

from utils.bob_safety import guarded_tool_decorator

# Configure logging to stderr (stdout is used for MCP protocol)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("ppt-mcp-bob")


SERVER_INSTRUCTIONS = """Work silently until the requested presentation is saved and validated.
For edits, activate and inspect the exact deck, set work mode with its full path
and slide count, then capture a shape snapshot. Keep allow_create=false unless
the user explicitly requests new content. Target existing objects by shape_id;
use ppt_transform_shapes, ppt_delete_shapes, and ppt_replace_visual for guarded
changes. Preserve the theme, master, layout, background, typography, and brand
objects unless explicitly targeted.

For newly designed slides, use ppt_get_design_reference when a React Bits-inspired
direction is useful. Treat it as visual guidance only: rebuild the composition
with editable PowerPoint-native text and shapes, and do not import or execute
React, CSS, canvas, or WebGL code inside the presentation.

After editing, run ppt_proofread_text with include_text_units=true, review the
returned text in context, correct confirmed issues, and rerun it. Then run
ppt_diff_shape_snapshot and ppt_validate_presentation before saving.

When complete, reply in the user's language with exactly three non-empty lines:
outcome, saved/exported path and scope, then validation result or one caveat.
"""


class MCPServer(_MCPServerBase):
    """FastMCP server that enforces IBM Bob safety on every tool."""

    def tool(self, *args: Any, **kwargs: Any):
        register = super().tool(*args, **kwargs)
        tool_name = kwargs.get("name")
        if not tool_name:
            raise ValueError("Every PowerPoint tool must declare an explicit name.")
        return guarded_tool_decorator(
            register,
            tool_name=tool_name,
            annotations=kwargs.get("annotations"),
        )


@asynccontextmanager
async def app_lifespan(server: MCPServer):
    """Manage COM lifecycle for the MCP server."""
    from utils.com_wrapper import ppt

    from utils.com_wrapper import AUTO_DISMISS_DIALOG
    logger.info("AUTO_DISMISS_DIALOG=%s (set PPT_AUTO_DISMISS_DIALOG=true to enable)", AUTO_DISMISS_DIALOG)
    logger.info("Starting PowerPoint COM worker thread...")
    ppt.start()
    # Do NOT connect to PowerPoint here. Connecting at startup would launch
    # PowerPoint.exe the moment the MCP client boots, even when the user never
    # invokes a ppt_* tool. The COM connection is established lazily on the
    # first tool call instead (see PowerPointCOMWrapper._get_app_impl).
    try:
        yield {}
    finally:
        logger.info("Shutting down PowerPoint COM worker thread...")
        ppt.stop()


mcp = MCPServer(
    "powerpoint_mcp",
    lifespan=app_lifespan,
    instructions=SERVER_INSTRUCTIONS,
)

from tools import register_tools

register_tools(mcp)


def main():
    """Entry point for the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()

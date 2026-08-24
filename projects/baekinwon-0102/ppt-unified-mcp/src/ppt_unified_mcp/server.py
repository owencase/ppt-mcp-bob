"""One MCP server for PPT Master generation and PowerPoint COM editing."""

from __future__ import annotations

import importlib.util
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


PACKAGE_LOCATION = Path(__file__).resolve().parent


def _find_integration_root() -> Path:
    seeds = (Path(__file__).resolve(), Path(sys.executable).resolve(), Path.cwd().resolve())
    for seed in seeds:
        for candidate in (seed, *seed.parents):
            if (
                candidate.name == "ppt-unified-mcp"
                and (candidate / "pyproject.toml").is_file()
            ):
                return candidate
    return Path.cwd().resolve()


INTEGRATION_ROOT = _find_integration_root()
CHECKOUT_ROOT = INTEGRATION_ROOT.parent


def _configured_path(env_name: str, default: Path) -> Path:
    return Path(os.environ.get(env_name, str(default))).expanduser().resolve()


def _configured_path_with_legacy(
    env_name: str, legacy_env_name: str, default: Path
) -> Path:
    value = os.environ.get(env_name) or os.environ.get(legacy_env_name) or str(default)
    return Path(value).expanduser().resolve()


PPT_TO_PYTHON_ROOT = _configured_path_with_legacy(
    "PPT_TO_PYTHON_ROOT", "PPT_MASTER_ROOT", CHECKOUT_ROOT / "ppt-to-python"
)
PPT_TO_COM_ROOT = _configured_path_with_legacy(
    "PPT_TO_COM_ROOT", "PPT_MCP_ROOT", CHECKOUT_ROOT / "ppt-to-com"
)
GENERATION_ROOT = _configured_path(
    "PPT_UNIFIED_WORKSPACE", INTEGRATION_ROOT / "workspace"
)
PPT_MASTER_SKILL = PPT_TO_PYTHON_ROOT / "skills" / "ppt-master"
PPT_MASTER_SCRIPTS = PPT_MASTER_SKILL / "scripts"
SERVER_MODE = os.environ.get("PPT_UNIFIED_MODE", "generate").strip().lower()
if SERVER_MODE not in {"all", "generate", "edit"}:
    raise RuntimeError("PPT_UNIFIED_MODE must be one of: all, generate, edit")
EDIT_PROFILE = os.environ.get("PPT_EDIT_PROFILE", "compact").strip().lower()
if EDIT_PROFILE not in {"compact", "full"}:
    raise RuntimeError("PPT_EDIT_PROFILE must be one of: compact, full")


def _load_com_server():
    server_path = PPT_TO_COM_ROOT / "src" / "server.py"
    if not server_path.is_file():
        raise RuntimeError(
            f"ppt-to-com was not found at {PPT_TO_COM_ROOT}. "
            "Set PPT_TO_COM_ROOT to its directory."
        )

    src_dir = str(server_path.parent)
    if src_dir not in sys.path:
        sys.path.insert(0, src_dir)

    spec = importlib.util.spec_from_file_location(
        "_ppt_unified_upstream_com_server", server_path
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load ppt-mcp server from {server_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


_com_server = _load_com_server()
mcp = _com_server.mcp

# New decks must go through PPT Master's python-pptx pipeline. The remaining
# upstream tools still open, inspect, edit, and save existing presentations.
mcp.remove_tool("ppt_create_presentation")

_upstream_instructions = mcp.instructions or ""
_mode_instructions = {
    "generate": """
This server is in NEW-PRESENTATION mode. Only use the `ppt_generate_*` tools.
The route uses PPT Master and python-pptx and does not require PowerPoint.
Read `ppt-guide://overview` and the relevant guide resource before creating
the first project. `ppt_generate_read_guide` remains available for clients
that do not support MCP resources.
""",
    "edit": """
This server is in EXISTING-PRESENTATION mode. Open an existing file with
`ppt_open_presentation`, activate it, inspect it, apply COM-backed edits, and
save it. Microsoft PowerPoint is required. Prefer Save As unless the user
explicitly authorizes overwriting the source file.
""",
    "all": """
- For a NEW presentation, use only `ppt_generate_*` tools. Start with
  `ppt_generate_read_guide`, create a project, write SVG slides, validate, and
  export. This route uses PPT Master and python-pptx; it does not launch
  PowerPoint.
- To MODIFY an EXISTING presentation, call `ppt_open_presentation`, then use
  the COM-backed `ppt_*` editing tools and save the file. Windows and Microsoft
  PowerPoint are required only for this route.
- Never use COM to create a new deck. The upstream COM creation tool is not
  registered in this unified server.
""",
}[SERVER_MODE]
if SERVER_MODE == "edit" and EDIT_PROFILE == "compact":
    _mode_instructions += """
This server uses the compact edit profile. Common inspection and file tools
are exposed directly. For every other COM capability, discover the operation
with `ppt_edit_capabilities`, then call `ppt_edit_execute` or combine related
operations with `ppt_edit_batch`. Hidden operations retain their original
input validation and behavior.
Capability schemas are also readable through the
`ppt-edit://capabilities/{operation}` resource template.
"""
_edit_instructions = _upstream_instructions if SERVER_MODE in {"all", "edit"} else ""
_unified_instructions = f"""
## Unified creation/editing router

{_mode_instructions.strip()}

{_edit_instructions}
""".strip()
if hasattr(mcp, "_mcp_server"):
    # FastMCP 1.x exposes instructions as a read-only facade over the
    # low-level protocol server. MCP 2.x retains this backing server.
    mcp._mcp_server.instructions = _unified_instructions
elif hasattr(mcp, "_instructions"):
    mcp._instructions = _unified_instructions
else:  # pragma: no cover - guards a future incompatible MCP release
    raise RuntimeError("Installed MCP version does not expose server instructions")


class GenerateGuideInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    document: Literal[
        "overview",
        "quick-generate",
        "canvas-formats",
        "shared-standards",
        "svg-effects",
    ] = Field(default="overview", description="PPT Master guide to read")


class CreateGenerationProjectInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    project_name: str = Field(
        min_length=1,
        max_length=80,
        description="Single-component project name without slashes",
    )
    canvas_format: Literal[
        "ppt169", "ppt43", "wechat", "xiaohongshu", "moments", "story", "banner", "a4"
    ] = "ppt169"


class ProjectInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    project_path: str = Field(description="Path returned by ppt_generate_create_project")


class ImportSourcesInput(ProjectInput):
    sources: list[str] = Field(
        min_length=1,
        description="Local source paths or URLs to archive in the generation project",
    )


class WriteSlideSvgInput(ProjectInput):
    slide_name: str = Field(
        min_length=1,
        max_length=100,
        description="Slide file name such as P01 or 01_cover.svg",
    )
    svg: str = Field(
        min_length=1,
        description="Complete SVG markup matching the project's selected canvas",
    )
    overwrite: bool = Field(default=False)


class ExportPresentationInput(ProjectInput):
    output_path: str | None = Field(
        default=None,
        description="Optional .pptx path. Relative paths resolve inside project/exports.",
    )
    with_notes: bool = Field(default=False)
    transition: str = Field(default="fade")
    native_charts_and_tables: bool = Field(default=False)
    conversion_trace: bool = Field(default=False)


GUIDE_FILES = {
    "quick-generate": PPT_MASTER_SKILL / "workflows" / "profiles" / "quick-generate.md",
    "canvas-formats": PPT_MASTER_SKILL / "references" / "canvas-formats.md",
    "shared-standards": PPT_MASTER_SKILL / "references" / "shared-standards-core.md",
    "svg-effects": PPT_MASTER_SKILL / "references" / "svg-effects.md",
}


def _result(success: bool, **values) -> str:
    return json.dumps({"success": success, **values}, ensure_ascii=False)


def _run_master_script(
    script_name: str, args: list[str], *, timeout: int = 600
) -> subprocess.CompletedProcess[str]:
    script = PPT_MASTER_SCRIPTS / script_name
    if not script.is_file():
        raise RuntimeError(
            f"PPT Master script not found: {script}. "
            "Set PPT_TO_PYTHON_ROOT to the ppt-to-python directory."
        )
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return subprocess.run(
        [sys.executable, str(script), *args],
        cwd=PPT_TO_PYTHON_ROOT,
        env=env,
        stdin=subprocess.DEVNULL,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        check=False,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )


def _create_master_project(
    project_name: str, canvas_format: str
) -> tuple[Path, dict]:
    """Create a Quick Generate project in-process without touching MCP stdio."""
    scripts_dir = str(PPT_MASTER_SCRIPTS)
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)

    from attribution_guard import require_skill_integrity
    from project_management.cli import ProjectManager
    from workflow_log import append_note

    require_skill_integrity()
    stdout = io.StringIO()
    stderr = io.StringIO()
    with redirect_stdout(stdout), redirect_stderr(stderr):
        project_path = Path(
            ProjectManager().init_project(
                project_name,
                canvas_format,
                base_dir=str(GENERATION_ROOT),
                quick_generate=True,
            )
        ).resolve()
        try:
            append_note(
                project_path,
                f"Project initialized: profile=quick; canvas={canvas_format}; "
                f"path={project_path}",
            )
        except OSError as exc:
            print(f"[WARN] Workflow audit unavailable: {exc}", file=sys.stderr)
    return project_path, {
        "mode": "in_process",
        "exit_code": 0,
        "stdout": stdout.getvalue().strip(),
        "stderr": stderr.getvalue().strip(),
    }


def _contained_path(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _resolve_project(project_path: str, *, must_exist: bool = True) -> Path:
    candidate = Path(project_path).expanduser()
    if not candidate.is_absolute():
        candidate = GENERATION_ROOT / candidate
    candidate = candidate.resolve()
    if not _contained_path(candidate, GENERATION_ROOT):
        raise ValueError(
            f"Project must be inside the configured generation workspace: {GENERATION_ROOT}"
        )
    if must_exist and not candidate.is_dir():
        raise ValueError(f"Generation project does not exist: {candidate}")
    return candidate


def _command_payload(completed: subprocess.CompletedProcess[str]) -> dict:
    return {
        "exit_code": completed.returncode,
        "stdout": completed.stdout.strip(),
        "stderr": completed.stderr.strip(),
    }


def _read_json_output(output: str) -> dict | list | None:
    stripped = output.strip()
    if not stripped:
        return None
    try:
        return json.loads(stripped)
    except json.JSONDecodeError:
        start = min(
            (index for index in (stripped.find("{"), stripped.find("[")) if index >= 0),
            default=-1,
        )
        if start >= 0:
            try:
                return json.loads(stripped[start:])
            except json.JSONDecodeError:
                pass
    return None


def _project_status(project: Path) -> dict:
    svg_dir = project / "svg_output"
    exports_dir = project / "exports"
    validation_dir = project / "validation"
    return {
        "project_path": str(project),
        "slides": [str(path) for path in sorted(svg_dir.glob("*.svg"))],
        "exports": [str(path) for path in sorted(exports_dir.glob("*.pptx"))]
        if exports_dir.is_dir()
        else [],
        "validation_reports": [
            str(path) for path in sorted(validation_dir.glob("*.json"))
        ]
        if validation_dir.is_dir()
        else [],
    }


@mcp.tool(
    name="ppt_generate_read_guide",
    annotations={
        "title": "Read PPT Master Generation Guide",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_generate_read_guide(params: GenerateGuideInput) -> str:
    """Read the upstream PPT Master guidance required to author a new deck."""
    if params.document == "overview":
        return _result(
            True,
            route="new-presentation",
            engine="ppt-master/python-pptx",
            steps=[
                "Read quick-generate plus the relevant visual references.",
                "Call ppt_generate_create_project.",
                "Optionally call ppt_generate_import_sources.",
                "Author complete SVG slides and call ppt_generate_write_slide_svg for each.",
                "Call ppt_generate_validate until blocking_errors is zero.",
                "Call ppt_generate_export_presentation.",
            ],
            editing_route=(
                "For an existing PPTX, use ppt_open_presentation and the COM-backed "
                "ppt_* editing tools instead."
            ),
            available_documents=list(GUIDE_FILES),
        )
    guide = GUIDE_FILES[params.document]
    if not guide.is_file():
        return _result(False, error=f"Guide not found: {guide}")
    return _result(True, document=params.document, content=guide.read_text(encoding="utf-8"))


@mcp.tool(
    name="ppt_generate_create_project",
    annotations={
        "title": "Create PPT Master Project",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def tool_generate_create_project(params: CreateGenerationProjectInput) -> str:
    """Create a lockless PPT Master Quick Generate workspace for a new deck."""
    try:
        GENERATION_ROOT.mkdir(parents=True, exist_ok=True)
        project, command = _create_master_project(
            params.project_name, params.canvas_format
        )
        project = _resolve_project(str(project))
        return _result(True, **_project_status(project), command=command)
    except Exception as exc:
        return _result(False, error=str(exc))


@mcp.tool(
    name="ppt_generate_import_sources",
    annotations={
        "title": "Import PPT Master Sources",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def tool_generate_import_sources(params: ImportSourcesInput) -> str:
    """Copy source files or URLs into a new-deck project using PPT Master."""
    try:
        project = _resolve_project(params.project_path)
        completed = _run_master_script(
            "project_manager.py",
            ["import-sources", str(project), *params.sources, "--copy"],
        )
        return _result(
            completed.returncode == 0,
            project_path=str(project),
            error=None if completed.returncode == 0 else "Source import failed",
            command=_command_payload(completed),
        )
    except Exception as exc:
        return _result(False, error=str(exc))


@mcp.tool(
    name="ppt_generate_write_slide_svg",
    annotations={
        "title": "Write PPT Master Slide SVG",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_generate_write_slide_svg(params: WriteSlideSvgInput) -> str:
    """Validate and atomically write one authored SVG slide into svg_output."""
    try:
        project = _resolve_project(params.project_path)
        file_name = params.slide_name
        if not file_name.lower().endswith(".svg"):
            file_name += ".svg"
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*\.svg", file_name):
            raise ValueError("slide_name must be a plain SVG file name without path separators")
        if "<!DOCTYPE" in params.svg.upper() or "<!ENTITY" in params.svg.upper():
            raise ValueError("DOCTYPE and ENTITY declarations are not allowed in slide SVG")
        root = ET.fromstring(params.svg)
        if root.tag.rsplit("}", 1)[-1] != "svg":
            raise ValueError("The document root must be an <svg> element")

        target = project / "svg_output" / file_name
        if target.exists() and not params.overwrite:
            raise FileExistsError(f"Slide already exists: {target}; set overwrite=true to replace it")
        target.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=target.parent,
            prefix=f".{target.stem}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            handle.write(params.svg)
            temporary = Path(handle.name)
        try:
            os.replace(temporary, target)
        finally:
            temporary.unlink(missing_ok=True)
        return _result(True, slide_path=str(target), **_project_status(project))
    except Exception as exc:
        return _result(False, error=str(exc))


@mcp.tool(
    name="ppt_generate_get_status",
    annotations={
        "title": "Get PPT Master Project Status",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_generate_get_status(params: ProjectInput) -> str:
    """List authored slides, quality reports, and exported PPTX files."""
    try:
        return _result(True, **_project_status(_resolve_project(params.project_path)))
    except Exception as exc:
        return _result(False, error=str(exc))


@mcp.tool(
    name="ppt_generate_validate",
    annotations={
        "title": "Validate PPT Master Slides",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_generate_validate(params: ProjectInput) -> str:
    """Run PPT Master's required lockless final SVG quality gate."""
    try:
        project = _resolve_project(params.project_path)
        completed = _run_master_script(
            "svg_quality_checker.py",
            [str(project), "--quick-generate", "--stage", "final", "--json"],
        )
        report = _read_json_output(completed.stdout)
        return _result(
            completed.returncode == 0,
            project_path=str(project),
            report=report,
            error=None if completed.returncode == 0 else "SVG quality gate failed",
            command=_command_payload(completed),
        )
    except Exception as exc:
        return _result(False, error=str(exc))


def _resolve_output(project: Path, output_path: str | None) -> Path | None:
    if output_path is None:
        return None
    candidate = Path(output_path).expanduser()
    if not candidate.is_absolute():
        candidate = project / "exports" / candidate
    candidate = candidate.resolve()
    if candidate.suffix.lower() != ".pptx":
        raise ValueError("output_path must end with .pptx")
    if not _contained_path(candidate, project):
        raise ValueError("output_path must stay inside the generation project")
    candidate.parent.mkdir(parents=True, exist_ok=True)
    return candidate


@mcp.tool(
    name="ppt_generate_export_presentation",
    annotations={
        "title": "Export New Presentation",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def tool_generate_export_presentation(params: ExportPresentationInput) -> str:
    """Validate SVGs, then export a new editable PPTX through python-pptx."""
    try:
        project = _resolve_project(params.project_path)
        output = _resolve_output(project, params.output_path)

        quality = _run_master_script(
            "svg_quality_checker.py",
            [str(project), "--quick-generate", "--stage", "final", "--json"],
        )
        if quality.returncode != 0:
            return _result(
                False,
                error="SVG quality gate failed; repair the reported slides before export",
                project_path=str(project),
                report=_read_json_output(quality.stdout),
                quality_command=_command_payload(quality),
            )

        args = [str(project), "--quick-generate"]
        args.append("--with-notes" if params.with_notes else "--no-notes")
        args.extend(["--transition", params.transition])
        if params.native_charts_and_tables:
            args.append("--native-charts-and-tables")
        if params.conversion_trace:
            args.append("--conversion-trace")
        if output is not None:
            args.extend(["--output", str(output)])

        started_at = time.time()
        exported = _run_master_script("svg_to_pptx.py", args, timeout=1200)
        if exported.returncode != 0:
            return _result(
                False,
                error="PPTX export failed",
                project_path=str(project),
                quality_report=_read_json_output(quality.stdout),
                export_command=_command_payload(exported),
            )

        if output is None:
            candidates = [
                path
                for path in (project / "exports").glob("*.pptx")
                if path.stat().st_mtime >= started_at - 2
            ]
            output = max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None
        return _result(
            True,
            project_path=str(project),
            output_path=str(output) if output else None,
            quality_report=_read_json_output(quality.stdout),
            export_command=_command_payload(exported),
        )
    except Exception as exc:
        return _result(False, error=str(exc))


class EditCapabilitiesInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    operation: str | None = Field(
        default=None,
        description="Exact original COM tool name to inspect, such as ppt_set_fill",
    )
    category: str | None = Field(
        default=None,
        description=(
            "Optional category: session, slides, content, design, data, export, "
            "motion, document, media, or advanced"
        ),
    )
    query: str | None = Field(
        default=None,
        max_length=100,
        description="Optional case-insensitive search over tool names and descriptions",
    )
    include_schema: bool = Field(
        default=False,
        description="Include input schemas in search results; exact operation lookup always includes it",
    )
    limit: int = Field(default=15, ge=1, le=30)


class EditOperationInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    operation: str = Field(
        min_length=1,
        max_length=100,
        description="Original COM tool name returned by ppt_edit_capabilities",
    )
    arguments: dict[str, Any] = Field(
        default_factory=dict,
        description="Arguments shown in the operation's input_schema; omit the outer params wrapper",
    )


class EditBatchInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    operations: list[EditOperationInput] = Field(min_length=1, max_length=30)
    stop_on_error: bool = Field(
        default=True,
        description="Stop after the first failed operation; completed COM edits are not rolled back",
    )


COMPACT_EDIT_CORE_TOOLS = frozenset(
    {
        "ppt_open_presentation",
        "ppt_activate_presentation",
        "ppt_get_presentation_info",
        "ppt_list_slides",
        "ppt_get_all_text",
        "ppt_list_shapes",
        "ppt_get_slide_preview",
        "ppt_save_presentation",
        "ppt_save_presentation_as",
        "ppt_close_presentation",
    }
)
EDIT_GATEWAY_TOOLS = frozenset(
    {"ppt_edit_capabilities", "ppt_edit_execute", "ppt_edit_batch"}
)

# Keep the validated upstream Tool objects available after compact mode removes
# their public schemas from the MCP tool manager.
EDIT_TOOL_REGISTRY = {
    tool.name: tool
    for tool in list(mcp._tool_manager.list_tools())
    if not tool.name.startswith("ppt_generate_")
}

EDIT_MODULE_CATEGORIES = {
    "app": "session",
    "presentation": "session",
    "slides": "slides",
    "shapes": "content",
    "text": "content",
    "placeholders": "content",
    "formatting": "design",
    "groups": "design",
    "connectors": "design",
    "layout": "design",
    "effects": "design",
    "themes": "design",
    "batch_apply": "design",
    "tables": "data",
    "charts": "data",
    "export": "export",
    "slideshow": "motion",
    "animation": "motion",
    "hyperlinks": "document",
    "sections": "document",
    "properties": "document",
    "comments": "document",
    "media": "media",
    "smartart": "media",
    "edit_ops": "advanced",
    "advanced_ops": "advanced",
    "freeform": "advanced",
}


def _edit_tool_category(tool: Any) -> str:
    module_name = getattr(tool.fn, "__module__", "").rsplit(".", 1)[-1]
    if module_name == "_ppt_unified_upstream_com_server":
        return "session" if tool.name != "ppt_get_slide_preview" else "content"
    return EDIT_MODULE_CATEGORIES.get(module_name, "advanced")


def _edit_tool_record(tool: Any, include_schema: bool = False) -> dict[str, Any]:
    description = " ".join((tool.description or "").split())
    record: dict[str, Any] = {
        "operation": tool.name,
        "category": _edit_tool_category(tool),
        "summary": description.split("\n", 1)[0][:240],
    }
    if include_schema:
        schema = tool.parameters
        properties = schema.get("properties", {})
        if list(properties) == ["params"]:
            parameter_schema = properties["params"]
            if "$ref" in parameter_schema:
                definition = parameter_schema["$ref"].rsplit("/", 1)[-1]
                parameter_schema = schema.get("$defs", {}).get(definition, parameter_schema)
            schema = parameter_schema
        record["input_schema"] = schema
    return record


def _normalize_edit_arguments(tool: Any, arguments: dict[str, Any]) -> dict[str, Any]:
    properties = tool.parameters.get("properties", {})
    if list(properties) == ["params"] and "params" not in arguments:
        return {"params": arguments}
    return arguments


async def _execute_edit_operation(operation: str, arguments: dict[str, Any]) -> Any:
    tool = EDIT_TOOL_REGISTRY.get(operation)
    if tool is None:
        raise ValueError(
            f"Unknown COM operation: {operation}. Use ppt_edit_capabilities first."
        )
    return await tool.run(_normalize_edit_arguments(tool, arguments))


@mcp.tool(
    name="ppt_edit_capabilities",
    annotations={
        "title": "Discover PowerPoint Edit Capabilities",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def tool_edit_capabilities(params: EditCapabilitiesInput) -> str:
    """Find hidden COM operations and retrieve their validated input schemas."""
    if params.operation:
        tool = EDIT_TOOL_REGISTRY.get(params.operation)
        if tool is None:
            return _result(False, error=f"Unknown COM operation: {params.operation}")
        return _result(True, capability=_edit_tool_record(tool, include_schema=True))

    category = params.category.lower() if params.category else None
    valid_categories = sorted(set(EDIT_MODULE_CATEGORIES.values()))
    if category and category not in valid_categories:
        return _result(
            False,
            error=f"Unknown category: {category}",
            categories=valid_categories,
        )
    query = params.query.lower() if params.query else None
    matches = []
    for tool in EDIT_TOOL_REGISTRY.values():
        if category and _edit_tool_category(tool) != category:
            continue
        searchable = f"{tool.name} {tool.description or ''}".lower()
        if query and query not in searchable:
            continue
        matches.append(_edit_tool_record(tool, params.include_schema))
    matches.sort(key=lambda item: item["operation"])
    return _result(
        True,
        categories=valid_categories,
        total_matches=len(matches),
        capabilities=matches[: params.limit],
        truncated=len(matches) > params.limit,
    )


@mcp.tool(
    name="ppt_edit_execute",
    annotations={
        "title": "Execute PowerPoint Edit Operation",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def tool_edit_execute(params: EditOperationInput) -> str:
    """Execute one original COM operation with its original input validation."""
    try:
        result = await _execute_edit_operation(params.operation, params.arguments)
        return result if isinstance(result, str) else _result(
            True, operation=params.operation, result=result
        )
    except Exception as exc:
        return _result(False, operation=params.operation, error=str(exc))


@mcp.tool(
    name="ppt_edit_batch",
    annotations={
        "title": "Execute PowerPoint Edit Batch",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def tool_edit_batch(params: EditBatchInput) -> str:
    """Execute up to 30 validated COM operations sequentially in one MCP call."""
    results = []
    for index, item in enumerate(params.operations):
        try:
            value = await _execute_edit_operation(item.operation, item.arguments)
            parsed = _read_json_output(value) if isinstance(value, str) else value
            results.append(
                {
                    "index": index,
                    "operation": item.operation,
                    "success": True,
                    "result": parsed if parsed is not None else value,
                }
            )
        except Exception as exc:
            results.append(
                {
                    "index": index,
                    "operation": item.operation,
                    "success": False,
                    "error": str(exc),
                }
            )
            if params.stop_on_error:
                break
    return _result(
        all(item["success"] for item in results),
        completed=len(results),
        requested=len(params.operations),
        results=results,
    )


def _register_mode_resources_and_prompts() -> None:
    """Register read-only context and user-selectable starts for active engines."""
    if SERVER_MODE in {"generate", "all"}:

        @mcp.resource(
            "ppt-guide://overview",
            name="ppt_generation_overview",
            title="PowerPoint Generation Overview",
            description="Read the new-presentation workflow before generating a deck.",
            mime_type="application/json",
        )
        async def generation_overview_resource() -> str:
            return await tool_generate_read_guide(GenerateGuideInput(document="overview"))

        @mcp.resource(
            "ppt-guide://documents/{document}",
            name="ppt_generation_guide",
            title="PowerPoint Generation Guide",
            description=(
                "Read one generation guide: quick-generate, canvas-formats, "
                "shared-standards, or svg-effects."
            ),
            mime_type="text/markdown",
        )
        def generation_guide_resource(document: str) -> str:
            guide = GUIDE_FILES.get(document)
            if guide is None:
                available = ", ".join(sorted(GUIDE_FILES))
                raise ValueError(
                    f"Unknown generation guide '{document}'. Available: {available}"
                )
            if not guide.is_file():
                raise FileNotFoundError(
                    f"Generation guide not found: {guide}. Check PPT_TO_PYTHON_ROOT."
                )
            return guide.read_text(encoding="utf-8")

        @mcp.resource(
            "ppt-project://{project_name}/status",
            name="ppt_generation_project_status",
            title="PowerPoint Generation Project Status",
            description=(
                "Read slides, validation reports, and exports for a project below "
                "PPT_UNIFIED_WORKSPACE. Use the project directory name, not a path."
            ),
            mime_type="application/json",
        )
        def generation_project_status_resource(project_name: str) -> str:
            project = _resolve_project(project_name)
            return _result(True, **_project_status(project))

        @mcp.prompt(
            name="create_powerpoint",
            title="Create a New PowerPoint",
            description=(
                "Start a new presentation through the validated python-pptx generation workflow."
            ),
        )
        def create_powerpoint_prompt(
            topic: str,
            audience: str = "general audience",
            requirements: str = "",
        ) -> str:
            return (
                "Create a new PowerPoint presentation.\n"
                f"Topic: {topic}\n"
                f"Audience: {audience}\n"
                f"Requirements: {requirements or 'Use a clear narrative and readable visual hierarchy.'}\n\n"
                "Use only the new-presentation route. Read ppt-guide://overview and "
                "ppt-guide://documents/quick-generate, create a project, author and "
                "validate every SVG slide, then export the editable PPTX."
            )

    if SERVER_MODE in {"edit", "all"}:

        @mcp.resource(
            "ppt-edit://categories",
            name="ppt_edit_categories",
            title="PowerPoint Edit Categories",
            description="Read compact COM capability categories and operation counts.",
            mime_type="application/json",
        )
        def edit_categories_resource() -> str:
            counts: dict[str, int] = {}
            for tool in EDIT_TOOL_REGISTRY.values():
                category = _edit_tool_category(tool)
                counts[category] = counts.get(category, 0) + 1
            return _result(
                True,
                total_operations=len(EDIT_TOOL_REGISTRY),
                categories=dict(sorted(counts.items())),
            )

        @mcp.resource(
            "ppt-edit://capabilities/{operation}",
            name="ppt_edit_capability",
            title="PowerPoint Edit Capability",
            description="Read the description and validated input schema for one COM operation.",
            mime_type="application/json",
        )
        def edit_capability_resource(operation: str) -> str:
            tool = EDIT_TOOL_REGISTRY.get(operation)
            if tool is None:
                raise ValueError(
                    f"Unknown COM operation '{operation}'. "
                    "Read ppt-edit://categories or use ppt_edit_capabilities to search."
                )
            return _result(True, capability=_edit_tool_record(tool, include_schema=True))

        @mcp.prompt(
            name="edit_powerpoint",
            title="Edit an Existing PowerPoint",
            description=(
                "Start a safe existing-presentation edit through Microsoft PowerPoint COM."
            ),
        )
        def edit_powerpoint_prompt(file_path: str, request: str) -> str:
            return (
                "Edit an existing PowerPoint presentation.\n"
                f"File: {file_path}\n"
                f"Requested changes: {request}\n\n"
                "Open and activate the exact file, inspect it before editing, read each "
                "needed operation schema through ppt-edit resources or "
                "ppt_edit_capabilities, apply only the requested changes, and save as a "
                "new file unless overwriting was explicitly requested."
            )


_register_mode_resources_and_prompts()


def _apply_server_mode() -> None:
    """Expose only the selected engine and edit profile to the MCP client."""
    tools = list(mcp._tool_manager.list_tools())
    for tool in tools:
        is_generation_tool = tool.name.startswith("ppt_generate_")
        is_gateway_tool = tool.name in EDIT_GATEWAY_TOOLS
        remove = False
        if SERVER_MODE == "generate":
            remove = not is_generation_tool
        elif SERVER_MODE == "edit":
            remove = is_generation_tool
            if EDIT_PROFILE == "compact" and not is_generation_tool:
                remove = tool.name not in COMPACT_EDIT_CORE_TOOLS | EDIT_GATEWAY_TOOLS
            elif EDIT_PROFILE == "full":
                remove = is_generation_tool or is_gateway_tool
        elif SERVER_MODE == "all":
            remove = is_gateway_tool
        if remove:
            mcp.remove_tool(tool.name)


_apply_server_mode()


def main() -> None:
    """Run the unified MCP server over stdio."""
    mcp.run()


if __name__ == "__main__":
    main()

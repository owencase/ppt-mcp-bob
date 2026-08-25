<p align="center">
  <img src="assets/ppt-mcp-logo-letter.png" alt="IBM Bob PowerPoint MCP" width="480">
</p>

<p align="center">
  <a href="README_ja.md">日本語</a>
</p>

<p align="center">
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue.svg" alt="Python 3.10+"></a>
  <img src="https://img.shields.io/badge/Platform-Windows-0078d4.svg" alt="Windows">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="MIT License">
  <img src="https://img.shields.io/badge/IBM%20Bob-ready-052FAD.svg" alt="IBM Bob ready">
</p>

<h1 align="center">IBM Bob PowerPoint MCP</h1>

<p align="center">
  <strong>Safe, real-time PowerPoint editing for IBM Bob through Windows COM automation.</strong>
</p>

This project is a Model Context Protocol (MCP) server built to let IBM Bob inspect and edit a live Microsoft PowerPoint presentation. It talks directly to the desktop PowerPoint application through Windows COM, so edits are visible immediately and preserve PowerPoint-native objects, layouts, themes, animations, and media.

Unlike file-only libraries such as `python-pptx`, this server operates on the presentation that is open in PowerPoint. It adds target locking, shape-state preconditions, validation, and retry-aware errors to make agent-driven editing safer.

## Highlights

- **Built for IBM Bob** — tool instructions and guardrails guide Bob through an inspect, edit, diff, and validate workflow.
- **Live COM automation** — controls the running PowerPoint application instead of reconstructing the deck from a file model.
- **165 tools in 28 categories** — covers presentations, slides, shapes, text, tables, charts, themes, animations, media, design references, export, and more.
- **Fail-closed targeting** — a missing or closed target never silently falls back to a different open deck.
- **Preconditioned editing** — high-level transform, delete, and visual-replacement tools verify the observed shape state before mutation.
- **Retry-aware failures** — MCP errors include a `retryable` value and a corrective `hint`, reducing unsafe blind retries.
- **Constrained file output** — save-as and export operations stay inside an explicit trusted output directory.
- **Native validation** — Bob can compare shape snapshots and validate the result before saving.

## Requirements

- Windows 10 or Windows 11
- Desktop Microsoft PowerPoint
- Python 3.10 or newer
- [`uv`](https://docs.astral.sh/uv/getting-started/installation/)

PowerPoint COM automation is Windows-only. The Python test suite can exercise most validation and policy logic without opening PowerPoint, but live editing requires PowerPoint on Windows.

## Install from this repository

Run the following from `projects/owencase`:

```powershell
uv sync --frozen
uv run ppt-mcp-bob
```

The server uses stdio for MCP communication. Running it directly leaves it waiting for an MCP client; normally IBM Bob starts it from its MCP configuration.

## Register with IBM Bob

Configure Bob to run this local source tree. Replace the example paths with absolute paths on the Windows machine:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "C:\\Users\\YOUR_NAME\\.local\\bin\\uv.exe",
      "args": [
        "--directory",
        "C:\\ABSOLUTE\\PATH\\TO\\ppt-mcp-bob\\projects\\owencase",
        "run",
        "ppt-mcp-bob"
      ],
      "env": {
        "PPT_TEMPLATES_DIR": "C:\\ABSOLUTE\\PATH\\TO\\templates",
        "PPT_MCP_OUTPUT_DIR": "C:\\ABSOLUTE\\PATH\\TO\\output"
      }
    }
  }
}
```

Use the local `--directory` configuration while developing or deploying this integration. `uvx ppt-mcp` resolves the separately published upstream package and does not run the code in this directory.

The legacy `ppt-mcp` command remains available as a compatibility alias, but `ppt-mcp-bob` is the canonical command for this project.

## Safe Bob workflow

For an existing presentation, Bob should follow this sequence:

1. Call `ppt_activate_presentation` and inspect the exact deck with `ppt_get_presentation_info` and `ppt_list_shapes`.
2. Call `ppt_set_work_mode` with the observed full path and slide count.
3. Keep `allow_create=false` for move, resize, edit, deletion, and replacement requests.
4. Capture a shape snapshot and retain the stable `shape_id` values.
5. Use `ppt_transform_shapes`, `ppt_delete_shapes`, or `ppt_replace_visual` with observed-state preconditions.
6. Call `ppt_proofread_text`, inspect every returned text unit contextually, correct confirmed spelling or spacing issues, and rerun it.
7. Call `ppt_diff_shape_snapshot` and `ppt_validate_presentation` before saving.
8. Save only after proofreading findings are empty and the diff and validation match the requested change.

`ppt_proofread_text` is read-only. It checks high-confidence Korean and English
typos, custom replacements, repeated words, punctuation, brackets, control
characters, and mojibake. It also returns location-aware text from shapes,
groups, tables, SmartArt, and charts so Bob can perform the required contextual
review. Use `allowed_terms` for product names and `custom_replacements` for
organization-specific terminology. Speaker notes are optional.

Bob receives a quiet-handoff policy from the MCP server. During presentation work,
it should not expose plans, tool calls or results, progress, or other
intermediate context. After saving and validation are complete, it should respond
in the user's language with exactly three lines covering the outcome, output path
and scope, and validation result. A client UI may still render its own tool-call
cards because their visibility is controlled by the client, not by MCP.

The default work mode is intentionally strict:

| Setting | Default | Effect |
|---|---:|---|
| `allow_create` | `false` | Blocks add, copy, and duplicate tools until creation is explicitly authorized. |
| `require_preconditions` | `true` | Requires the expected presentation path and shape-state checks for mutations. |

## Environment variables

| Variable | Required | Description |
|---|---:|---|
| `PPT_TEMPLATES_DIR` | No | Directory scanned by `ppt_list_templates` when no directory argument is provided. |
| `PPT_MCP_OUTPUT_DIR` | Recommended | Trust boundary for save-as and export paths. If omitted, the locked presentation's local directory is used when available. |
| `PPT_DOWNLOAD_TIMEOUT_SECONDS` | No | Timeout for remote images and icon metadata. Default: `15`. |
| `PPT_MAX_DOWNLOAD_BYTES` | No | Maximum remote download size in bytes. Default: `20971520` (20 MiB). |
| `PPT_AUTO_DISMISS_DIALOG` | No | Sends Escape when PowerPoint rejects a COM call as busy. Disabled by default. |

`PPT_AUTO_DISMISS_DIALOG=true` is useful for unattended runs, but it can cancel a dialog the user currently has open. Leave it disabled for interactive sessions unless automatic dismissal is intended.

## React Bits design references

`ppt_get_design_reference` provides a small set of curated React Bits links and
translates the selected visual direction into a fixed 16:9, PowerPoint-native
design contract. The result includes palette, typography, layout, native shape
ideas, restrained animation guidance, and the existing `ppt_*` tools to use.

This is intentionally a reference-only connection. It does not install React,
fetch component source, or attempt to reproduce CSS, canvas, WebGL, hover,
cursor, or scroll behavior. Core presentation content remains editable native
PowerPoint objects.

## Tool categories

| Category | Tools | Coverage |
|---|---:|---|
| App | 5 | Connection, application state, active window, and open presentations |
| Presentation | 8 | Create, open, save, close, activate, inspect, and list templates |
| Slides | 10 | Add, delete, duplicate, move, copy, inspect, notes, and navigation |
| Shapes | 10 | Add and inspect shapes, text boxes, pictures, and lines using stable IDs |
| Safe editing | 8 | Work mode, guarded transform/delete/replace, snapshots, diff, and validation |
| Text | 11 | Text content, ranges, paragraphs, bullets, search, extraction, typography, and proofreading |
| Placeholders | 6 | Inspect and update placeholder content |
| Formatting | 3 | Fill, line, and shadow |
| Tables | 13 | Data, cells, rows, columns, merge/split, style, layout, and borders |
| Export | 3 | PDF, images, and clipboard copy |
| Slideshow | 6 | Start, stop, navigate, and inspect slideshow state |
| Charts | 7 | Create, inspect, populate, format, and change chart types |
| Animation | 6 | Transitions and animation lifecycle operations |
| Themes | 4 | Theme application, theme colors, and headers/footers |
| Groups | 3 | Group, ungroup, and inspect group items |
| Connectors | 2 | Add and format connectors |
| Hyperlinks | 3 | Add, inspect, and remove hyperlinks |
| Sections | 3 | Add, list, and manage sections |
| Properties | 2 | Read and update presentation metadata |
| Media | 3 | Video, audio, and media settings |
| SmartArt | 3 | Add, modify, and list SmartArt layouts |
| Edit operations | 6 | Undo, redo, and copy shapes or formatting |
| Layout | 7 | Align, distribute, size, background, flip, and merge |
| Effects | 3 | Glow, reflection, and soft-edge effects |
| Comments | 3 | Add, list, and delete comments |
| Advanced | 19 | Tags, fonts, crop, picture operations, selection, icons, URLs, and batch apply |
| Freeform | 7 | Build paths and inspect or modify freeform nodes |
| Design reference | 1 | React Bits-inspired PowerPoint-native design recipes and public reference links |
| **Total** | **165** | |

## Development

From `projects/owencase`:

```powershell
uv sync --group dev
uv run pytest
```

The test suite covers schema strictness, target locking, retry behavior, path constraints, stable shape IDs, slide operations, and validation logic.

## License and credits

Released under the MIT License.

This IBM Bob integration is maintained by [owencase](https://github.com/owencase) and builds on the PowerPoint MCP work from [ykuwai/ppt-mcp](https://github.com/ykuwai/ppt-mcp). It uses [FastMCP](https://github.com/jlowin/fastmcp), [pywin32](https://github.com/mhammond/pywin32), and the [Model Context Protocol](https://modelcontextprotocol.io/).

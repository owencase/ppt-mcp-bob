#!/usr/bin/env node
/**
 * PPT MCP Server
 * Exposes PowerPoint (.pptx) authoring tools to LLM clients via the MCP protocol.
 * The actual file manipulation is delegated to the sibling python-pptx bridge
 * package (packages/ppt-bridge), invoked once per tool call over stdin/stdout JSON.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

// ---------------------------------------------------------------------------
// Resolve the Python bridge script path
//
// This file runs from <package root>/build, so the bridge lives in the sibling
// package packages/ppt-bridge. Expressing it as "package root -> sibling package"
// keeps it correct whether we run the compiled build/ or the sources in src/.
// BRIDGE_SCRIPT env var overrides resolution (useful for non-ASCII paths).
// ---------------------------------------------------------------------------
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PACKAGE_ROOT = path.resolve(__dirname, "..");
const BRIDGE_SCRIPT =
  process.env.BRIDGE_SCRIPT ??
  path.join(PACKAGE_ROOT, "..", "ppt-bridge", "bridge.py");
const PYTHON_BIN = process.env.PYTHON_BIN ?? "python3";

// ---------------------------------------------------------------------------
// Helper: call the Python bridge
// ---------------------------------------------------------------------------
interface BridgeResult {
  success: boolean;
  message?: string;
  data?: unknown;
  error?: string;
}

function callBridge(action: string, params: Record<string, unknown>): BridgeResult {
  const payload = JSON.stringify({ action, params });
  const result = spawnSync(PYTHON_BIN, [BRIDGE_SCRIPT], {
    input: payload,
    encoding: "utf-8",
    timeout: 30_000,
    env: { ...process.env, PYTHONIOENCODING: "utf-8", PYTHONUTF8: "1" },
  });

  if (result.error) {
    return { success: false, error: `Process error: ${result.error.message}` };
  }
  if (result.status !== 0) {
    return { success: false, error: result.stderr?.trim() || "Bridge exited with non-zero status" };
  }

  try {
    return JSON.parse(result.stdout) as BridgeResult;
  } catch {
    return { success: false, error: `Invalid JSON from bridge: ${result.stdout}` };
  }
}

function toolResponse(result: BridgeResult) {
  const text = result.success
    ? result.message ?? JSON.stringify(result.data ?? {})
    : `Error: ${result.error}`;
  return {
    content: [{ type: "text" as const, text }],
    isError: !result.success,
  };
}

// ---------------------------------------------------------------------------
// MCP Server
// ---------------------------------------------------------------------------
const server = new McpServer({ name: "ppt-mcp-server", version: "0.1.0" });

// ── Tool: create_presentation ──────────────────────────────────────────────
server.registerTool(
  "create_presentation",
  {
    description: "Create a new blank PowerPoint presentation and save it to disk.",
    inputSchema: z.object({
      file_path: z.string().describe("Absolute path where the .pptx file will be saved"),
      width_cm: z.number().optional().describe("Slide width in centimetres (default 33.87 = widescreen)"),
      height_cm: z.number().optional().describe("Slide height in centimetres (default 19.05 = widescreen)"),
    }),
  },
  async ({ file_path, width_cm, height_cm }) => {
    return toolResponse(callBridge("create_presentation", { file_path, width_cm, height_cm }));
  }
);

// ── Tool: add_slide ───────────────────────────────────────────────────────
server.registerTool(
  "add_slide",
  {
    description: "Append a new slide to an existing presentation.",
    inputSchema: z.object({
      file_path: z.string().describe("Path to the .pptx file"),
      layout_index: z.number().int().min(0).default(6).describe("Slide layout index (0-based). 6 = blank"),
    }),
  },
  async ({ file_path, layout_index }) => {
    return toolResponse(callBridge("add_slide", { file_path, layout_index }));
  }
);

// ── Tool: add_text_box ────────────────────────────────────────────────────
server.registerTool(
  "add_text_box",
  {
    description: "Add a text box to a specific slide.",
    inputSchema: z.object({
      file_path: z.string().describe("Path to the .pptx file"),
      slide_index: z.number().int().min(0).describe("0-based slide index"),
      text: z.string().describe("Text content"),
      left_cm: z.number().describe("Left position in cm"),
      top_cm: z.number().describe("Top position in cm"),
      width_cm: z.number().describe("Width in cm"),
      height_cm: z.number().describe("Height in cm"),
      font_size_pt: z.number().optional().default(24).describe("Font size in points"),
      bold: z.boolean().optional().default(false),
      color_hex: z.string().optional().default("000000").describe("Font colour as RRGGBB hex string"),
      align: z.enum(["left", "center", "right"]).optional().default("left"),
    }),
  },
  async (args) => {
    return toolResponse(callBridge("add_text_box", args));
  }
);

// ── Tool: add_image ───────────────────────────────────────────────────────
server.registerTool(
  "add_image",
  {
    description: "Insert an image file onto a slide.",
    inputSchema: z.object({
      file_path: z.string().describe("Path to the .pptx file"),
      slide_index: z.number().int().min(0).describe("0-based slide index"),
      image_path: z.string().describe("Absolute path to the image file (PNG/JPG)"),
      left_cm: z.number(),
      top_cm: z.number(),
      width_cm: z.number(),
      height_cm: z.number(),
    }),
  },
  async (args) => {
    return toolResponse(callBridge("add_image", args));
  }
);

// ── Tool: set_background_color ────────────────────────────────────────────
server.registerTool(
  "set_background_color",
  {
    description: "Fill the background of a slide with a solid colour.",
    inputSchema: z.object({
      file_path: z.string().describe("Path to the .pptx file"),
      slide_index: z.number().int().min(0),
      color_hex: z.string().describe("Background colour as RRGGBB hex string, e.g. '1E1E2E'"),
    }),
  },
  async (args) => {
    return toolResponse(callBridge("set_background_color", args));
  }
);

// ── Tool: add_shape ───────────────────────────────────────────────────────
server.registerTool(
  "add_shape",
  {
    description: "Add a filled rectangle or rounded-rectangle shape to a slide.",
    inputSchema: z.object({
      file_path: z.string().describe("Path to the .pptx file"),
      slide_index: z.number().int().min(0),
      shape_type: z.enum(["rectangle", "rounded_rectangle"]).default("rectangle"),
      left_cm: z.number(),
      top_cm: z.number(),
      width_cm: z.number(),
      height_cm: z.number(),
      fill_color_hex: z.string().default("4472C4").describe("Fill colour as RRGGBB hex"),
      line_color_hex: z.string().optional().describe("Border colour as RRGGBB hex (omit for no border)"),
    }),
  },
  async (args) => {
    return toolResponse(callBridge("add_shape", args));
  }
);

// ── Tool: apply_theme ─────────────────────────────────────────────────────
server.registerTool(
  "apply_theme",
  {
    description:
      "Apply a predefined design theme to every slide. Sets the slide background colour and returns the theme's accent palette for use in follow-up add_text_box / add_shape calls.",
    inputSchema: z.object({
      file_path: z.string().describe("Path to the .pptx file"),
      theme: z.enum(["minimal_dark", "minimal_light", "tech_blue", "marketing_warm"]).describe(
        "Theme name. minimal_dark: dark bg + white text. minimal_light: white bg + dark text. tech_blue: navy + accent. marketing_warm: warm tones."
      ),
    }),
  },
  async (args) => {
    return toolResponse(callBridge("apply_theme", args));
  }
);

// ── Tool: save_presentation ───────────────────────────────────────────────
server.registerTool(
  "save_presentation",
  {
    description: "Save (overwrite) the presentation file to disk.",
    inputSchema: z.object({
      file_path: z.string().describe("Path to the .pptx file to save"),
    }),
  },
  async ({ file_path }) => {
    return toolResponse(callBridge("save_presentation", { file_path }));
  }
);

// ── Tool: get_presentation_info ───────────────────────────────────────────
server.registerTool(
  "get_presentation_info",
  {
    description: "Return metadata about an existing presentation: slide count, dimensions, and per-slide shape list.",
    inputSchema: z.object({
      file_path: z.string().describe("Path to the .pptx file"),
    }),
  },
  async ({ file_path }) => {
    return toolResponse(callBridge("get_presentation_info", { file_path }));
  }
);

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("[ppt-mcp-server] running on stdio");
}

main().catch((err) => {
  console.error("[ppt-mcp-server] fatal:", err);
  process.exit(1);
});

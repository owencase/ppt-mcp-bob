#!/usr/bin/env node
/**
 * html-ppt-mcp
 * HTML 기반으로 프레젠테이션을 만드는 MCP 서버.
 *
 * 여기가 시작점입니다. tool 을 추가하려면 아래 registerTool 을 따라 하면 됩니다.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

const server = new McpServer({ name: "html-ppt-mcp", version: "0.1.0" });

// ── Tool 예시 ──────────────────────────────────────────────────────────────
// description 은 LLM 이 이 tool 을 언제 쓸지 판단하는 근거입니다. 구체적으로 씁니다.
server.registerTool(
  "render_slide_html",
  {
    description: "제목과 본문으로 16:9 슬라이드 HTML 한 장을 만든다.",
    inputSchema: z.object({
      title: z.string().describe("슬라이드 제목"),
      body: z.string().describe("본문 텍스트"),
      // 색상은 레포 전체에서 '#' 없는 RRGGBB hex 로 통일합니다.
      bg_hex: z.string().optional().default("0F172A").describe("배경색 (RRGGBB)"),
    }),
  },
  async ({ title, body, bg_hex }) => {
    const html = renderSlide(title, body, bg_hex);
    return { content: [{ type: "text" as const, text: html }] };
  }
);

/** 사용자 입력을 HTML 에 넣기 전 반드시 이스케이프합니다. */
function escapeHtml(text: string): string {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function renderSlide(title: string, body: string, bgHex: string): string {
  return `<!doctype html>
<html><head><meta charset="utf-8"><style>
  body { margin:0; width:1280px; height:720px; background:#${bgHex}; color:#E2E8F0;
         font-family: system-ui, sans-serif; padding:80px; box-sizing:border-box; }
  h1 { font-size:56px; margin:0 0 32px; }
  p  { font-size:28px; line-height:1.6; }
</style></head>
<body><h1>${escapeHtml(title)}</h1><p>${escapeHtml(body)}</p></body></html>`;
}

// ── Boot ───────────────────────────────────────────────────────────────────
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // 로그는 반드시 stderr 로. stdout 은 MCP 프로토콜 전용입니다.
  console.error("[html-ppt-mcp] running on stdio");
}

main().catch((err) => {
  console.error("[html-ppt-mcp] fatal:", err);
  process.exit(1);
});

/**
 * html-ppt-mcp 테스트.
 *
 * 실행:
 *   cd packages/html-ppt-mcp
 *   npm test            # pretest 로 빌드가 먼저 돕니다
 *
 * 지금은 최소 두 개만 있습니다. tool 을 추가하면 여기에도 추가해 주세요.
 * 레퍼런스 예시: packages/mcp-server/test/tools.test.mjs
 */
import { test, describe, before, after } from "node:test";
import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SERVER_ENTRY = path.resolve(__dirname, "..", "build", "index.js");

let client;

before(async () => {
  assert.ok(existsSync(SERVER_ENTRY), `빌드 산출물이 없습니다. 'npm run build' 를 먼저 실행하세요.`);
  client = new Client({ name: "test", version: "1.0.0" });
  await client.connect(new StdioClientTransport({ command: "node", args: [SERVER_ENTRY] }));
});

after(async () => {
  await client?.close();
});

describe("서버", () => {
  // 서버가 부팅 중에 죽으면 LLM 클라이언트에는 그냥 "tool 이 없음" 으로 보입니다.
  // 원인을 찾기 어려운 실패라 여기서 잡습니다.
  test("tool 이 노출된다", async () => {
    const { tools } = await client.listTools();
    assert.ok(tools.length > 0, "tool 이 하나도 없습니다");
    for (const tool of tools) {
      assert.equal(tool.inputSchema.type, "object", `${tool.name}: inputSchema 가 object 가 아닙니다`);
      // description 은 LLM 이 이 tool 을 언제 쓸지 판단하는 유일한 근거입니다.
      assert.ok(tool.description?.length > 10, `${tool.name}: description 이 너무 짧습니다`);
    }
  });

  // HTML 을 만드는 서버라 이스케이프가 빠지면 바로 인젝션이 됩니다.
  // 슬라이드 제목에 <script> 를 넣는 건 사용자가 실수로도 할 수 있습니다.
  test("사용자 입력이 HTML 에 그대로 들어가지 않는다", async () => {
    const result = await client.callTool({
      name: "render_slide_html",
      arguments: { title: "<script>alert(1)</script>", body: "본문 & 기호" },
    });
    const html = result.content[0].text;
    assert.ok(!html.includes("<script>alert(1)</script>"), "제목이 이스케이프되지 않았습니다");
    assert.ok(html.includes("&lt;script&gt;"), "이스케이프된 형태가 보이지 않습니다");
    assert.ok(html.includes("&amp;"), "본문의 & 가 이스케이프되지 않았습니다");
  });
});

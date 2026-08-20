/**
 * mcp-server 테스트.
 *
 * 실행:
 *   cd reference/mcp-server
 *   npm test            # pretest 로 빌드가 먼저 돕니다
 *
 * MCP SDK 의 진짜 Client 로 진짜 서버 프로세스에 붙습니다. LLM 클라이언트가
 * 하는 것과 같은 경로라, 여기서 통과하면 실제로도 붙습니다.
 */
import { test, describe, before, after } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PACKAGE_ROOT = path.resolve(__dirname, "..");
const SERVER_ENTRY = path.join(PACKAGE_ROOT, "build", "index.js");
const BRIDGE_SCRIPT = path.resolve(PACKAGE_ROOT, "..", "ppt-bridge", "bridge.py");

/** venv 가 있으면 그걸 쓰고, 없으면 PYTHON_BIN, 그것도 없으면 python3. */
function resolvePython() {
  if (process.env.PYTHON_BIN) return process.env.PYTHON_BIN;
  const venv = path.resolve(PACKAGE_ROOT, "..", "ppt-bridge", ".venv", "bin", "python");
  return existsSync(venv) ? venv : "python3";
}

let client;

before(async () => {
  assert.ok(existsSync(SERVER_ENTRY), `빌드 산출물이 없습니다: ${SERVER_ENTRY}\n먼저 'npm run build' 를 실행하세요.`);
  client = new Client({ name: "test", version: "1.0.0" });
  await client.connect(new StdioClientTransport({ command: "node", args: [SERVER_ENTRY] }));
});

after(async () => {
  await client?.close();
});

// ───────────────────────────────────────────────────────────────────────────
// 1. tool 이 실제로 노출되는가
//
// registerTool 을 써 놓고 서버가 부팅 중에 죽으면 LLM 클라이언트에는 그냥
// "tool 이 없음"으로 보입니다. 원인을 찾기 어려운 실패라 여기서 잡습니다.
// ───────────────────────────────────────────────────────────────────────────
describe("tool 목록", () => {
  test("9개 tool 이 모두 노출된다", async () => {
    const { tools } = await client.listTools();
    const names = tools.map((t) => t.name).sort();
    assert.deepEqual(names, [
      "add_image",
      "add_shape",
      "add_slide",
      "add_text_box",
      "apply_theme",
      "create_presentation",
      "get_presentation_info",
      "save_presentation",
      "set_background_color",
    ]);
  });

  test("모든 tool 이 file_path 를 받고 설명이 붙어 있다", async () => {
    const { tools } = await client.listTools();
    for (const tool of tools) {
      assert.equal(tool.inputSchema.type, "object", `${tool.name}: inputSchema 가 object 가 아닙니다`);
      assert.ok(
        tool.inputSchema.properties?.file_path,
        `${tool.name}: file_path 파라미터가 없습니다. 모든 액션은 대상 파일을 받습니다.`
      );
      // description 은 LLM 이 이 tool 을 언제 쓸지 판단하는 유일한 근거입니다.
      assert.ok(tool.description?.length > 10, `${tool.name}: description 이 너무 짧습니다`);
    }
  });
});

// ───────────────────────────────────────────────────────────────────────────
// 2. TS ↔ Python 계약 드리프트
//
// 액션 하나를 추가하려면 bridge.py 의 HANDLERS 와 index.ts 의 registerTool
// 을 둘 다 고쳐야 합니다. 한쪽만 고치면 컴파일도 통과하고 테스트도 없으면
// 아무도 모릅니다. 그 상태를 이 테스트가 잡습니다.
//
// 의도적으로 한쪽에만 두는 액션이 생기면, 이 테스트에 예외 목록을 만들고
// 왜 예외인지 주석을 남겨 주세요.
// ───────────────────────────────────────────────────────────────────────────
describe("브릿지 계약", () => {
  test("bridge.py 의 액션 목록과 tool 목록이 일치한다", async (t) => {
    const python = resolvePython();
    const probe = spawnSync(python, [BRIDGE_SCRIPT], {
      input: JSON.stringify({ action: "", params: {} }),
      encoding: "utf-8",
      timeout: 30_000,
    });

    let available;
    try {
      available = JSON.parse(probe.stdout).data.available;
    } catch {
      const hint =
        `브릿지를 실행할 수 없습니다 (python: ${python}).\n` +
        `  cd reference/ppt-bridge && python3 -m venv .venv && .venv/bin/pip install -r requirements.txt\n` +
        `  stderr: ${(probe.stderr ?? "").trim().split("\n").slice(-3).join("\n  ")}`;
      // CI 에서는 건너뛰면 안 됩니다. 로컬에서는 파이썬 준비가 안 됐을 수 있으니 skip.
      if (process.env.CI) assert.fail(hint);
      t.skip(hint);
      return;
    }

    const { tools } = await client.listTools();
    const toolNames = tools.map((x) => x.name).sort();

    assert.deepEqual(
      toolNames,
      available,
      "index.ts 의 tool 과 bridge.py 의 HANDLERS 가 어긋났습니다. 양쪽을 모두 고쳤는지 확인하세요."
    );
  });
});

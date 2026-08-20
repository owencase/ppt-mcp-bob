import test from "node:test";
import assert from "node:assert/strict";
import { escapeHtml, safeName, slideHtml } from "../src/render.mjs";
import { themes } from "../src/themes.mjs";

test("escapeHtml은 태그를 이스케이프한다", () => {
  assert.equal(escapeHtml('<img src=x>'), "&lt;img src=x&gt;");
});

test("safeName은 파일명 금지문자를 제거한다", () => {
  assert.equal(safeName('a:b/c*?'), "a-b-c--");
});

test("slideHtml은 사용자 텍스트를 그대로 넣지 않는다", () => {
  const html = slideHtml({ title: "<script>", body: "안녕" });
  assert.ok(!html.includes("<script>"));
  assert.ok(html.includes("&lt;script&gt;"));
});

test("모든 테마가 필요한 색상 키를 가진다", () => {
  for (const [name, theme] of Object.entries(themes)) {
    for (const key of ["bg", "text", "accent"]) {
      assert.ok(/^[0-9A-F]{6}$/i.test(theme[key]), `${name}.${key} 형식 오류`);
    }
  }
});

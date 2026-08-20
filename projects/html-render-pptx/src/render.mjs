/**
 * 슬라이드 → HTML 렌더링.
 *
 * 여기가 실제 작업 지점입니다. 지금은 최소 구현만 있습니다.
 */
import { getTheme } from "./themes.mjs";

/** HTML 특수문자를 이스케이프합니다. 사용자 입력을 그대로 넣으면 안 됩니다. */
export function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

/** 파일명으로 쓸 수 없는 문자를 '-'로 바꿉니다. */
export function safeName(name) {
  return String(name).replace(/[\\/:*?"<>|]/g, "-");
}

/** 슬라이드 하나를 16:9 HTML로 만듭니다. */
export function slideHtml({ title, body, theme = "midnight" }) {
  const { bg, text, accent } = getTheme(theme);
  return `<!doctype html>
<html><head><meta charset="utf-8"><style>
  body { margin:0; width:1280px; height:720px; background:#${bg}; color:#${text};
         font-family: system-ui, sans-serif; padding:80px; box-sizing:border-box; }
  h1 { color:#${accent}; font-size:56px; margin:0 0 32px; }
  p  { font-size:28px; line-height:1.6; }
</style></head>
<body><h1>${escapeHtml(title)}</h1><p>${escapeHtml(body)}</p></body></html>`;
}

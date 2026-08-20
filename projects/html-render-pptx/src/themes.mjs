/**
 * 슬라이드 테마 정의.
 * 색상은 '#' 없는 RRGGBB hex 문자열로 통일합니다 (레포 전체 규칙).
 */
export const themes = {
  midnight: { bg: "0F172A", text: "E2E8F0", accent: "38BDF8" },
  editorial: { bg: "FFFFFF", text: "1C1917", accent: "F97316" },
};

export function getTheme(name) {
  const theme = themes[name];
  if (!theme) {
    throw new Error(`Unknown theme '${name}'. Available: ${Object.keys(themes).join(", ")}`);
  }
  return theme;
}

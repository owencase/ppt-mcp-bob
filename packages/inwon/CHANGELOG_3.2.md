# 3.2 COM template-editing fixes

- Removed forced `TextFrame2.AutoSize = 2` during COM text replacement.
- Captures and restores template font name, size, bold, italic, color and original AutoSize mode.
- Measures PowerPoint text bounds and tries semantic/local shortening before font reduction.
- Limits automatic font reduction to max 12.5%, max 4pt, never below 14pt.
- Fails with `TEXT_OVERFLOW_UNRESOLVED` instead of resizing/moving shapes.
- Splits design validation into hard errors and warnings.
- Tolerates PowerPoint theme-to-RGB and implicit-to-explicit run normalization.
- Still rejects slide/shape count, geometry, rotation, explicit fill/line, font and large typography changes.
- Records actual replacement text, rewrite method and font-size changes in COM operation manifests.
- Keeps fully visible watch mode and mandatory mode confirmation from 3.1.

Validation performed in this build environment:
- `pytest`: 17 passed
- `scripts/run_tests.py`: passed
- `scripts/smoke_test.py`: generation/routing passed
- Windows PowerPoint COM E2E was not run because this build environment is not Windows with desktop PowerPoint.

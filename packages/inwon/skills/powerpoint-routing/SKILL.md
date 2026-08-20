---
name: powerpoint-routing
description: Route PowerPoint requests between python-pptx generation and Windows PowerPoint COM template editing, with mandatory user confirmation before any file-changing action.
---

# PowerPoint Generation / Template Editing Router

## Purpose

Use this skill for every request to create, make, generate, modify, rewrite, or update a PowerPoint presentation with this MCP.

The system has exactly two execution modes:

- `generate`: create a new deck from scratch with the existing `python-pptx` pipeline.
- `template_com`: keep an existing file in `/template` visually unchanged and modify its presentation content through Microsoft PowerPoint COM automation.

## Non-negotiable mode gate

**Never create or edit a PPT file immediately. Always ask the user which mode to use first, even when the wording makes the intent obvious.**

Before any file-changing PowerPoint tool call:

1. Call `prepare_presentation_task` with the user's original request.
2. Show the returned confirmation question to the user.
3. Stop. Do not call create/edit tools in the same turn.
4. After the user explicitly chooses a mode, call `confirm_presentation_mode` with the returned `confirmation_id` and the selected mode.
5. Pass the resulting one-time `execution_token` to exactly one matching execution tool.

Do not infer consent. Do not turn a recommended mode into confirmation. Do not mint an execution token before the user's reply.

## Intent recommendation

Use these only to make a recommendation in the question; they never remove the confirmation requirement.

- Requests such as “PPT 만들어줘”, “~ 발표자료 생성해줘”, “처음부터 만들어줘” → recommend `generate`.
- Requests such as “~ 템플릿 내용만 수정해줘”, “이 디자인 그대로 내용만 ~ 주제로 바꿔줘”, “기존 PPT의 텍스트만 바꿔줘” → recommend `template_com`.
- Ambiguous requests → do not recommend either mode strongly; ask the same two-choice question.

## Mode: generate

After confirmation:

1. Use `create_presentation` with the `generate` execution token.
2. Do not supply a template path.
3. Use the existing pipeline: research → DeckPlan → semantic QA → python-pptx rendering → render QA.
4. Preserve all existing content grounding, overflow, contrast, image, and layout validation rules.

## Mode: template_com

After confirmation:

1. If needed, call `list_templates` to show the files currently in `/template`.
2. The user must identify the desired template if more than one could match.
3. Use `edit_template_presentation` with the `template_com` execution token.
4. The template must come from `/template`; never use an arbitrary outside path.
5. Never overwrite the original template. Write to a separate output path.
6. Change visible text content through PowerPoint COM. Do not rebuild the deck with python-pptx.
7. **Watch mode is mandatory:** PowerPoint must be visible on screen for the entire COM edit. Never request or pass a hidden/background mode.
8. Before each text replacement, navigate the PowerPoint window to that slide and select the exact text box so the user can see what will change.
9. Pause briefly before and after each replacement (default `step_delay=0.55`, allowed 0.20~5.0 seconds) so the edit is perceptible.
10. After saving, leave the edited presentation open in PowerPoint so the user can inspect the final result.
11. Preserve slide count, shape count, shape geometry, rotation, fill/line styling, images, charts, and other visual structure.
12. Never force `TextFrame2.AutoSize=2`. Preserve the template's existing AutoSize and visible text style after replacement.
13. **Do not fail fast during the first COM pass.** Apply every slide/text box possible, record per-shape errors/overflow, and finish/save the whole deck before QA.
14. Run post-QA only after the full first pass. Repair only the affected slide/text box, not the whole presentation. Default `max_post_qa_rounds=2`; never exceed 3.
15. If text overflows during post-QA, shorten the replacement using only facts already present; only then allow a limited font-size reduction (max 12.5% and max 4pt, never below 14pt). Never resize or move the shape to make text fit.
16. If the same issue signature repeats, stop repair immediately with the cycle breaker. Never restart the whole template job automatically.
17. Final unresolved QA issues must be returned in the manifest (`completed_with_unresolved_issues`, `automatic_restart_blocked=true`) rather than throwing a content/design QA exception that could trigger an external retry loop.
18. **Bob-safe MCP boundary:** expected mode-gate, COM, save, semantic-QA, and post-QA failures must be returned as normal JSON whenever possible, with `tool_call_succeeded=true`, `mcp_transport_error=false`, `do_not_retry=true`. Do not turn them into tool exceptions.
19. If `do_not_retry=true` is returned, never call `edit_template_presentation` again automatically with the same request or execution token. Report the current output/log state to the user instead.
20. Treat PowerPoint theme/RGB or implicit/explicit run normalization as benign. Record actual structure/geometry/color/font changes as final QA errors; limited overflow font reduction is a warning.
21. Use the explicit PowerPoint presentation handle opened for the output file; never rely on whichever PowerPoint window happens to be active.
22. Suppress PowerPoint modal alerts during automated edits while keeping the PowerPoint window visible; collect errors into post-QA instead.
23. If PowerPoint is temporarily busy and COM returns call-rejected/retry-later errors, use the MCP's bounded retry behavior.
24. If Windows, Microsoft PowerPoint, or pywin32 is unavailable, report that COM template editing cannot run on that machine. Do not silently fall back to python-pptx template editing.

## Required user-facing question

The wording may be localized, but it must contain both choices and make the engine difference clear:

> PPT 작업 방식을 먼저 확인할게요. ① 처음부터 새 PPT를 생성할까요? (python-pptx) ② `/template`의 기존 템플릿 디자인은 그대로 두고 내용만 수정할까요? (PowerPoint COM)

## Final checks

Before reporting success, verify:

- A user mode choice occurred before execution.
- The execution token mode matches the tool used.
- `generate` used python-pptx and no template path.
- `template_com` used PowerPoint COM and a `/template` source.
- The original template was not overwritten.
- Template mode may finish with unresolved QA issues. Report `completion_status`, `passed`, `design_preserved`, and `post_validation`; do not automatically run the whole job again.
- Template mode reports `watch_mode.enabled: true`, `powerpoint_visible: true`, and `keep_result_open: true`.

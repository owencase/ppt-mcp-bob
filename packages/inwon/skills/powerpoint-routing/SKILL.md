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
7. Preserve slide count, shape count, shape geometry, rotation, fill/line styling, images, charts, and other visual structure.
8. Use the explicit PowerPoint presentation handle opened for the output file; never rely on whichever PowerPoint window happens to be active.
9. If PowerPoint is temporarily busy and COM returns call-rejected/retry-later errors, use the MCP's bounded retry behavior.
10. If Windows, Microsoft PowerPoint, or pywin32 is unavailable, report that COM template editing cannot run on that machine. Do not silently fall back to python-pptx template editing.

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
- Template mode reports `design_preserved: true`; otherwise treat the operation as failed.

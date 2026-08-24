---
name: powerpoint-routing
description: Route PowerPoint requests between new presentation generation with PPT Master/python-pptx and existing presentation editing with Microsoft PowerPoint COM. Use for any request to create, modify, update, or export a PPT/PPTX deck.
---

# PowerPoint Routing

Classify the request before calling any PowerPoint MCP tool.

## Route selection

- Select **Generate** when no existing PowerPoint file is being used as the
  editable starting point. This includes topic-to-deck and source-to-deck work.
- Select **Edit** when the user supplies an existing `.ppt`, `.pptx`, `.pptm`,
  `.potx`, or `.ppsx` file to open, revise, fill, or save as another file.
- A request to change or fill an existing template is Edit, even when the
  result is saved under a new name.
- If the starting-file intent is genuinely unclear, ask one short question
  before invoking either server.

Exactly one MCP server must be enabled in Bob for the selected route:

- Generate: enable `powerpoint-generate`; disable `powerpoint-edit`.
- Edit: enable `powerpoint-edit`; disable `powerpoint-generate`.

If the selected server is unavailable, stop and tell the user which Bob MCP
toggle to change. Do not fall back to the opposite engine. Do not edit
`.bob/mcp.json` during an ordinary presentation request.

After routing, read only the matching reference:

- Generate: [references/generate.md](references/generate.md)
- Edit: [references/edit.md](references/edit.md)

## Timeout boundary

Do not automatically repeat a timed-out non-idempotent create, edit, or export
call. First inspect project status, the expected output path, or the active
presentation. Bob's Network timeout for the active server should be set to 5
minutes; a timeout is not proof that the underlying operation failed.

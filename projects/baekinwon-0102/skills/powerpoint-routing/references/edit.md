# Edit Mode

Use only tools from `powerpoint-edit`. Windows and Microsoft PowerPoint are
required.

1. Open the exact supplied file with `ppt_open_presentation`.
2. Lock the session target with `ppt_activate_presentation`.
3. Inspect presentation information and existing content before changing it.
4. For an edit without a direct tool, read
   `ppt-edit://capabilities/{operation}` when the exact operation is known. Use
   `ppt_edit_capabilities` with a narrow category/query when discovering it or
   when the client does not expose MCP resources. Read the returned
   `input_schema` before constructing arguments.
5. Apply one operation with `ppt_edit_execute`. Use `ppt_edit_batch` for a
   short ordered sequence only after every operation schema has been checked.
   Batch steps are not transactional: successful earlier edits remain when a
   later step fails.
6. Save with `ppt_save_presentation_as` by default so the source remains
   recoverable. Overwrite with `ppt_save_presentation` only when the user
   explicitly requested in-place modification.

Pass only the fields inside the returned input schema as `arguments`; do not
add the upstream tool's outer `params` wrapper. Do not guess operation names or
arguments, and do not request every capability schema at once.

Never call `ppt_generate_*` tools in this route. If PowerPoint reports a modal
dialog or rejected COM call, close the dialog and retry the single affected
operation; do not replay the entire edit sequence.

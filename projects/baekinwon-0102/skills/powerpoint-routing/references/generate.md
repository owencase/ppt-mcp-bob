# Generate Mode

Use only tools from `powerpoint-generate`. This server exposes seven
`ppt_generate_*` tools and does not launch Microsoft PowerPoint.

1. Read `ppt-guide://overview`, then read only the detailed `ppt-guide` resource
   needed for the deck. If the client does not expose MCP resources, use
   `ppt_generate_read_guide` as the compatibility fallback.
2. Create one project with `ppt_generate_create_project`. Preserve the returned
   absolute `project_path` for every later call.
3. Import user-supplied sources when present.
4. Author each complete SVG slide and write it with
   `ppt_generate_write_slide_svg`. Use stable ordered names such as `P01.svg`.
5. Run `ppt_generate_validate`. Repair blocking findings before export.
6. Run `ppt_generate_export_presentation` once and return its `output_path`.

Never call COM tools in this route. Do not retry project creation with the same
name after a timeout; read the project's `ppt-project` status resource or call
`ppt_generate_get_status` against the expected workspace project first.

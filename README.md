# ppt-mcp-no-powerpoint

**Local PowerPoint Intelligence MCP for template-faithful corporate decks.**

Creates, edits, and validates `.pptx` files **without depending on PowerPoint** —
no COM automation, no Office requirement, and **everything offline and local**:
no network calls, ever. See [DESIGN.md](DESIGN.md) for the full design.

## Status

**M0–M5 done:** deck sessions, reading, placeholder-first authoring, template
intelligence, compliance validation, rendering, brand-governed image
placeholders, template application, and the eval suite.

| Tool | What it does |
|---|---|
| `ppt_open_deck` / `ppt_create_deck` | Open a `.pptx`/`.potx` into an isolated working copy / create blank |
| `ppt_save_deck` / `ppt_close_deck` / `ppt_list_decks` | Session lifecycle |
| `ppt_undo` | Snapshot-based rollback of mutations |
| `ppt_get_deck_overview` | Slide size, masters + layouts, per-slide listing |
| `ppt_get_slide` | Full shape inventory: placeholder roles, geometry, text, notes |
| `ppt_search_deck` | Search slide text and speaker notes |
| `ppt_add_slide` | Create from a layout, map semantic content onto placeholders (dry-run-able) |
| `ppt_set_placeholder_content` | Fill one placeholder by role / idx / shape_id |
| `ppt_edit_text` | replace_text / set_paragraphs / append_paragraph ops |
| `ppt_set_notes` | Speaker notes |
| `ppt_add_table` / `ppt_add_image` / `ppt_add_chart` | Placeholder-first; freeform only behind `allow_freeform=true` |
| `ppt_delete_slide` / `ppt_move_slide` / `ppt_duplicate_slide` | Slide ops (duplicate copies images + notes) |
| `ppt_register_template` | Parse a `.potx`/`.pptx` into a design-system profile (immutable copy) |
| `ppt_list_templates` / `ppt_inspect_template` / `ppt_inspect_layout` | Registry inspection: layouts, placeholder schemas, theme, intent tags |
| `ppt_update_template` | Curate names/versions/intent tags |
| `ppt_extract_theme` | Color + font scheme from a template or open deck |
| `ppt_recommend_layout` | Rank layouts for a slide intent + content shape, with reasons |
| `ppt_map_content_to_placeholders` | Preview the content→placeholder mapping without touching a deck |
| `ppt_validate_compliance` | Rules C01–C10: layout provenance, theme fonts/colors, footers, covered logos, placeholder bypass, overflow, density, off-grid, slide size |
| `ppt_repair_compliance` | Conservative auto-fixes (theme font relink, color snap); dry-run by default |
| `ppt_render_slide` / `ppt_render_deck` | PNG renders via headless LibreOffice (returns the image inline) |
| `ppt_export_pdf` | PDF export |
| `ppt_visual_diff` | Pixel-diff current deck vs a pre-mutation snapshot ("did the logo survive?") |
| `ppt_set_style_profile` / `ppt_get_style_profile` / `ppt_list_style_profiles` / `ppt_delete_style_profile` | Local, versioned brand style profiles for imagery |
| `ppt_compose_image_prompt` | Deterministic prompt assembly: profile + slide context + constraints |
| `ppt_create_image_placeholder` | Governed image slot; full prompt lives in an in-file manifest + notes pointer |
| `ppt_list_image_placeholders` / `ppt_update_image_placeholder` | Manifest inspection/refinement (prompts re-validated against the profile) |
| `ppt_fill_image_placeholder` | Insert the externally generated image; status flips to `generated` |
| `ppt_apply_template` | Re-target a deck onto a template: per-slide fidelity report (dry-run default), placeholder content migrated, orphans carried over + flagged, auto-validation |
| `ppt_extract_template_from_deck` | Recover a design system from a finished deck as a derived template with layout-usage stats |

Rendering needs LibreOffice (`winget install TheDocumentFoundation.LibreOffice`);
everything else works without it. Renders are validation evidence — PowerPoint
remains the fidelity arbiter.

Try `uv run python scripts/make_demo_deck.py` → `examples/demo.pptx`, built
entirely through the tools with zero absolute positioning.

A companion skill for MCP clients lives in
[skills/ppt-deck-authoring](skills/ppt-deck-authoring/SKILL.md).

The 10-question agent eval suite lives in [evals/](evals/README.md).

Next per the [roadmap](DESIGN.md#14-milestones): M6 — built-in starter
templates, slide patterns (business forms), vendored icon sets.

## Quickstart

Requires Python ≥ 3.11 and [uv](https://docs.astral.sh/uv/).

```powershell
uv sync
uv run pytest          # all green
uv run python scripts/smoke_stdio.py   # end-to-end stdio handshake
```

Register with Claude Code:

```powershell
claude mcp add ppt-mcp -- uv --directory H:\repos\ppt-mcp-no-powerpoint run ppt-mcp
```

Or in any MCP client config:

```json
{
  "mcpServers": {
    "ppt-mcp": {
      "command": "uv",
      "args": ["--directory", "H:\\repos\\ppt-mcp-no-powerpoint", "run", "ppt-mcp"]
    }
  }
}
```

Local data lives in `~/.ppt-mcp/` (override with `PPT_MCP_HOME`).

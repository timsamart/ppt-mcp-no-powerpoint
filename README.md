# ppt-mcp-no-powerpoint

**Local PowerPoint Intelligence MCP for template-faithful corporate decks.**

Creates, edits, and validates `.pptx` files **without depending on PowerPoint** —
no COM automation, no Office requirement, and **everything offline and local**:
no network calls, ever. See [DESIGN.md](DESIGN.md) for the full design.

## Status

**M0 + M1 + M2 done:** deck sessions, reading, placeholder-first authoring,
and template intelligence.

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

Try `uv run python scripts/make_demo_deck.py` → `examples/demo.pptx`, built
entirely through the tools with zero absolute positioning.

A companion skill for MCP clients lives in
[skills/ppt-deck-authoring](skills/ppt-deck-authoring/SKILL.md).

Next per the [roadmap](DESIGN.md#14-milestones): compliance + rendering (M3),
brand style profiles (M4).

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

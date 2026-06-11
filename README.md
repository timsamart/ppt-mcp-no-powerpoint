# ppt-mcp-no-powerpoint

**Local PowerPoint Intelligence MCP for template-faithful corporate decks.**

Creates, edits, and validates `.pptx` files **without depending on PowerPoint** —
no COM automation, no Office requirement, and **everything offline and local**:
no network calls, ever. See [DESIGN.md](DESIGN.md) for the full design.

## Status

**M0 (in progress):** deck sessions + reading tools are working.

| Tool | What it does |
|---|---|
| `ppt_open_deck` / `ppt_create_deck` | Open a `.pptx`/`.potx` into an isolated working copy / create blank |
| `ppt_save_deck` / `ppt_close_deck` / `ppt_list_decks` | Session lifecycle |
| `ppt_undo` | Snapshot-based rollback of mutations |
| `ppt_get_deck_overview` | Slide size, masters + layouts, per-slide listing |
| `ppt_get_slide` | Full shape inventory: placeholder roles, geometry, text, notes |
| `ppt_search_deck` | Search slide text and speaker notes |

Next per the [roadmap](DESIGN.md#14-milestones): authoring (M1), template
intelligence (M2), compliance + rendering (M3), brand style profiles (M4).

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

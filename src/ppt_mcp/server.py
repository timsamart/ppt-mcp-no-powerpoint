"""ppt_mcp — FastMCP stdio server (DESIGN.md §10, M0 surface).

Deck lifecycle + reading/inspection tools. All logging goes to stderr
(stdio transport owns stdout). The server makes no network calls, ever.
"""

from __future__ import annotations

import logging
import sys
from typing import Annotated, Any, Literal

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from . import format as fmt
from . import reader
from .errors import PptMcpError
from .sessions import SessionManager
from .store import Store

logging.basicConfig(stream=sys.stderr, level=logging.INFO)
log = logging.getLogger("ppt_mcp")

mcp = FastMCP("ppt_mcp")

store = Store()
sessions = SessionManager(store)

READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=False)
MUTATING = ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=False)
DESTRUCTIVE = ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=False)

ResponseFormat = Literal["markdown", "json"]


def _prs(deck_id: str):
    return reader.load_presentation(sessions.get(deck_id).working_path)


# -- deck lifecycle (§10.1) ----------------------------------------------------


@mcp.tool(annotations=MUTATING)
def ppt_open_deck(
    path: Annotated[str, Field(description="Absolute path to a .pptx or .potx file")],
) -> dict[str, Any]:
    """Open an existing PowerPoint deck for reading/editing. The file is copied
    into an isolated session working copy; the source is not touched until
    ppt_save_deck. Returns the deck_id handle plus a deck overview."""
    session = sessions.open_deck(path)
    overview = reader.deck_overview(reader.load_presentation(session.working_path))
    return {
        "deck_id": session.deck_id,
        "source_path": str(session.source_path),
        "overview": overview,
    }


@mcp.tool(annotations=MUTATING)
def ppt_create_deck() -> dict[str, Any]:
    """Create a new blank deck (template-based creation arrives with the
    template registry). Save it later with ppt_save_deck(path=...)."""
    session = sessions.create_deck()
    return {"deck_id": session.deck_id, "source_path": None}


@mcp.tool(annotations=MUTATING)
def ppt_save_deck(
    deck_id: str,
    path: Annotated[
        str | None,
        Field(description="Target path. Defaults to the path the deck was opened from."),
    ] = None,
) -> dict[str, Any]:
    """Write the deck's working copy to disk. Without a path, saves back to the
    source file the deck was opened from."""
    target = sessions.save_deck(deck_id, path)
    return {"deck_id": deck_id, "saved_to": str(target)}


@mcp.tool(annotations=DESTRUCTIVE)
def ppt_close_deck(deck_id: str) -> dict[str, Any]:
    """Close a deck session and delete its working copy and snapshots.
    Unsaved changes are lost — call ppt_save_deck first if needed."""
    sessions.close_deck(deck_id)
    return {"deck_id": deck_id, "closed": True}


@mcp.tool(annotations=READ_ONLY)
def ppt_list_decks() -> dict[str, Any]:
    """List open deck sessions with their ids and source paths."""
    return {
        "decks": [
            {
                "deck_id": s.deck_id,
                "source_path": str(s.source_path) if s.source_path else None,
                "opened_at": s.opened_at,
                "undo_steps_available": s.snapshot_count,
            }
            for s in sessions.list_sessions()
        ]
    }


@mcp.tool(annotations=MUTATING)
def ppt_undo(
    deck_id: str,
    steps: Annotated[int, Field(ge=1, description="Number of mutations to roll back")] = 1,
) -> dict[str, Any]:
    """Undo the last N mutations on a deck by restoring a pre-mutation snapshot."""
    undone = sessions.undo(deck_id, steps)
    remaining = sessions.get(deck_id).snapshot_count
    if undone == 0:
        return {"deck_id": deck_id, "undone_steps": 0, "note": "Nothing to undo."}
    return {"deck_id": deck_id, "undone_steps": undone, "undo_steps_remaining": remaining}


# -- reading & inspection (§10.2) -----------------------------------------------


@mcp.tool(annotations=READ_ONLY)
def ppt_get_deck_overview(
    deck_id: str,
    response_format: ResponseFormat = "markdown",
) -> str | dict[str, Any]:
    """Deck summary: slide size/aspect ratio, masters with their layouts, and a
    one-line-per-slide listing (index, layout, title)."""
    data = reader.deck_overview(_prs(deck_id))
    if response_format == "json":
        return {"deck_id": deck_id, **data}
    return fmt.overview_markdown(deck_id, data)


@mcp.tool(annotations=READ_ONLY)
def ppt_get_slide(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1, description="1-based, as in the PowerPoint UI")],
    response_format: ResponseFormat = "markdown",
) -> str | dict[str, Any]:
    """Full inventory of one slide: every shape with its shape_id, placeholder
    role/idx, geometry (inches + EMU, layout-inherited where applicable), text
    with indent levels, tables/charts, and speaker notes."""
    data = reader.slide_detail(_prs(deck_id), slide_index)
    if response_format == "json":
        return {"deck_id": deck_id, **data}
    return fmt.slide_markdown(data)


@mcp.tool(annotations=READ_ONLY)
def ppt_search_deck(
    deck_id: str,
    query: Annotated[str, Field(min_length=1, description="Case-insensitive substring")],
    limit: Annotated[int, Field(ge=1, le=200)] = 25,
    offset: Annotated[int, Field(ge=0)] = 0,
    response_format: ResponseFormat = "markdown",
) -> str | dict[str, Any]:
    """Search all slide text and speaker notes. Returns hits with slide index
    and shape_id so results can be addressed directly by follow-up calls."""
    data = reader.search_deck(_prs(deck_id), query, limit=limit, offset=offset)
    if response_format == "json":
        return {"deck_id": deck_id, **data}
    return fmt.search_markdown(data)


def main() -> None:
    log.info("ppt_mcp starting (stdio); data dir: %s", store.root)
    mcp.run()


if __name__ == "__main__":
    main()

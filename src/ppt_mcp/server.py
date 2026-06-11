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
from . import reader, writer
from .errors import PptMcpError
from .models import (
    ChartSpec,
    ContentSpec,
    EditOp,
    ImageRef,
    Position,
    ShapeContent,
    TableSpec,
)
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


def _commit(deck_id: str, prs, tool: str, **provenance: object) -> None:
    """Snapshot-then-write for every mutation: snapshot the pre-state, save
    the changed presentation to the working copy, log provenance."""
    sessions.snapshot(deck_id)
    prs.save(str(sessions.get(deck_id).working_path))
    store.log_provenance(tool, deck_id=deck_id, **provenance)


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


# -- authoring (§10.3) -----------------------------------------------------------


@mcp.tool(annotations=MUTATING)
def ppt_add_slide(
    deck_id: str,
    layout: Annotated[str, Field(description="Layout name; ppt_get_deck_overview lists them")],
    content: ContentSpec,
    position: Annotated[
        int | None, Field(ge=1, description="1-based target position; default: append")
    ] = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Create a slide from a layout and fill its placeholders from semantic
    content (title/subtitle/body/table/image/chart/notes). Content that fits no
    placeholder is reported in unplaced_content — never silently dropped. Use
    dry_run=true to preview the mapping plan."""
    prs = _prs(deck_id)
    layout_obj = writer.resolve_layout(prs, layout)
    plan = writer.plan_content_mapping(layout_obj, content)
    if position is not None and position > len(prs.slides) + 1:
        raise PptMcpError(
            f"position {position} out of range; deck has {len(prs.slides)} slide(s)."
        )
    if dry_run:
        return {"applied": False, "plan": plan}
    new_index = writer.apply_add_slide(prs, layout_obj, content, plan, position)
    _commit(deck_id, prs, "ppt_add_slide", layout=layout_obj.name, slide_index=new_index)
    return {"applied": True, "slide_index": new_index, "plan": plan}


@mcp.tool(annotations=MUTATING)
def ppt_set_placeholder_content(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1)],
    shape_ref: Annotated[
        str, Field(description="Role ('title', 'body', ...), 'idx:N', or a shape_id")
    ],
    content: ShapeContent,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Replace the content of one placeholder, addressed by semantic role,
    placeholder idx, or shape_id. Exactly one content kind must be given."""
    prs = _prs(deck_id)
    slide = reader.get_slide_by_index(prs, slide_index)
    shape = writer.resolve_shape(slide, shape_ref)
    plan = {
        "slide_index": slide_index,
        "shape_id": shape.shape_id,
        "shape_name": shape.name,
        "content_kind": next(k for k, v in content.__dict__.items() if v is not None),
    }
    if dry_run:
        return {"applied": False, "plan": plan}
    result = writer.insert_into_placeholder(shape, content)
    _commit(deck_id, prs, "ppt_set_placeholder_content", slide_index=slide_index)
    return {"applied": True, "plan": plan, **result}


@mcp.tool(annotations=MUTATING)
def ppt_edit_text(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1)],
    shape_ref: str,
    ops: Annotated[
        list[EditOp],
        Field(description="Sequence of replace_text / set_paragraphs / append_paragraph ops"),
    ],
    dry_run: bool = False,
) -> dict[str, Any]:
    """Edit text of one shape with targeted operations. replace_text keeps
    run-level formatting when the match lies within a single run."""
    prs = _prs(deck_id)
    slide = reader.get_slide_by_index(prs, slide_index)
    shape = writer.resolve_shape(slide, shape_ref)
    plan = {
        "slide_index": slide_index,
        "shape_id": shape.shape_id,
        "ops": [op.op for op in ops],
    }
    if dry_run:
        return {"applied": False, "plan": plan}
    result = writer.apply_edit_ops(shape, ops)
    _commit(deck_id, prs, "ppt_edit_text", slide_index=slide_index)
    return {"applied": True, "plan": plan, **result}


@mcp.tool(annotations=MUTATING)
def ppt_set_notes(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1)],
    notes: str,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Set the speaker notes of a slide (replaces existing notes)."""
    prs = _prs(deck_id)
    slide = reader.get_slide_by_index(prs, slide_index)
    if dry_run:
        return {"applied": False, "plan": {"slide_index": slide_index, "set": "notes"}}
    slide.notes_slide.notes_text_frame.text = notes
    _commit(deck_id, prs, "ppt_set_notes", slide_index=slide_index)
    return {"applied": True, "slide_index": slide_index}


def _placed_add(
    deck_id: str,
    slide_index: int,
    content: ShapeContent,
    shape_ref: str | None,
    position: Position | None,
    allow_freeform: bool,
    dry_run: bool,
    freeform_apply,
    tool: str,
) -> dict[str, Any]:
    """Shared placeholder-first/freeform-gated flow for table/image/chart."""
    prs = _prs(deck_id)
    slide = reader.get_slide_by_index(prs, slide_index)
    kind = next(k for k, v in content.__dict__.items() if v is not None)
    if shape_ref is not None:
        shape = writer.resolve_shape(slide, shape_ref)
        plan = {"slide_index": slide_index, "target": f"placeholder shape_id {shape.shape_id}", "kind": kind}
        if dry_run:
            return {"applied": False, "plan": plan}
        result = writer.insert_into_placeholder(shape, content)
        _commit(deck_id, prs, tool, slide_index=slide_index, mode="placeholder")
        return {"applied": True, "plan": plan, **result}
    if not allow_freeform or position is None:
        raise PptMcpError(
            f"No shape_ref given. Either target a suitable placeholder (see "
            f"ppt_get_slide), or pass allow_freeform=true together with an explicit "
            f"position to place the {kind} freely. {writer.FREEFORM_WARNING}"
        )
    plan = {"slide_index": slide_index, "target": "freeform", "kind": kind,
            "warnings": [writer.FREEFORM_WARNING]}
    if dry_run:
        return {"applied": False, "plan": plan}
    new_shape = freeform_apply(slide, position)
    _commit(deck_id, prs, tool, slide_index=slide_index, mode="freeform")
    return {"applied": True, "plan": plan, "shape_id": new_shape.shape_id,
            "warnings": [writer.FREEFORM_WARNING]}


@mcp.tool(annotations=MUTATING)
def ppt_add_table(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1)],
    table: TableSpec,
    shape_ref: Annotated[str | None, Field(description="Target placeholder (preferred)")] = None,
    position: Position | None = None,
    allow_freeform: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Add a table — into a table placeholder when shape_ref is given
    (preferred), or freeform only with allow_freeform=true and a position."""
    return _placed_add(
        deck_id, slide_index, ShapeContent(table=table), shape_ref, position,
        allow_freeform, dry_run,
        lambda slide, pos: writer.add_freeform_table(slide, table, pos),
        "ppt_add_table",
    )


@mcp.tool(annotations=MUTATING)
def ppt_add_image(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1)],
    image: ImageRef,
    shape_ref: Annotated[str | None, Field(description="Target picture placeholder (preferred)")] = None,
    position: Position | None = None,
    allow_freeform: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Add a local image — into a picture placeholder when shape_ref is given
    (preferred), or freeform only with allow_freeform=true and a position."""
    return _placed_add(
        deck_id, slide_index, ShapeContent(image=image), shape_ref, position,
        allow_freeform, dry_run,
        lambda slide, pos: writer.add_freeform_image(slide, image, pos),
        "ppt_add_image",
    )


@mcp.tool(annotations=MUTATING)
def ppt_add_chart(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1)],
    chart: ChartSpec,
    shape_ref: Annotated[str | None, Field(description="Target chart placeholder (preferred)")] = None,
    position: Position | None = None,
    allow_freeform: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Add a native chart (column/bar/line/pie/doughnut/area) — into a chart
    placeholder when shape_ref is given, or freeform with allow_freeform=true."""
    return _placed_add(
        deck_id, slide_index, ShapeContent(chart=chart), shape_ref, position,
        allow_freeform, dry_run,
        lambda slide, pos: writer.add_freeform_chart(slide, chart, pos),
        "ppt_add_chart",
    )


@mcp.tool(annotations=DESTRUCTIVE)
def ppt_delete_slide(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1)],
    dry_run: bool = False,
) -> dict[str, Any]:
    """Delete a slide. Recoverable via ppt_undo while the session is open."""
    prs = _prs(deck_id)
    title = reader.slide_title(reader.get_slide_by_index(prs, slide_index))
    plan = {"delete": {"slide_index": slide_index, "title": title}}
    if dry_run:
        return {"applied": False, "plan": plan}
    writer.delete_slide(prs, slide_index)
    _commit(deck_id, prs, "ppt_delete_slide", slide_index=slide_index)
    return {"applied": True, "plan": plan, "slide_count": len(prs.slides)}


@mcp.tool(annotations=MUTATING)
def ppt_move_slide(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1)],
    to_position: Annotated[int, Field(ge=1)],
    dry_run: bool = False,
) -> dict[str, Any]:
    """Move a slide to a new 1-based position."""
    prs = _prs(deck_id)
    plan = {"move": {"from": slide_index, "to": to_position}}
    if dry_run:
        return {"applied": False, "plan": plan}
    writer.move_slide(prs, slide_index, to_position)
    _commit(deck_id, prs, "ppt_move_slide", slide_index=slide_index, to=to_position)
    return {"applied": True, "plan": plan}


@mcp.tool(annotations=MUTATING)
def ppt_duplicate_slide(
    deck_id: str,
    slide_index: Annotated[int, Field(ge=1)],
    dry_run: bool = False,
) -> dict[str, Any]:
    """Duplicate a slide (shapes, images, notes); the copy lands right after
    the source."""
    prs = _prs(deck_id)
    title = reader.slide_title(reader.get_slide_by_index(prs, slide_index))
    plan = {"duplicate": {"slide_index": slide_index, "title": title}}
    if dry_run:
        return {"applied": False, "plan": plan}
    new_index = writer.duplicate_slide(prs, slide_index)
    _commit(deck_id, prs, "ppt_duplicate_slide", slide_index=slide_index)
    return {"applied": True, "plan": plan, "new_slide_index": new_index}


def main() -> None:
    log.info("ppt_mcp starting (stdio); data dir: %s", store.root)
    mcp.run()


if __name__ == "__main__":
    main()

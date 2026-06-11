"""Markdown renderers for inspection tools (DESIGN.md §11.2).

Markdown is the default response format: skimmable, low-noise, IDs kept
inline so follow-up tool calls can address what they see.
"""

from __future__ import annotations

from typing import Any


def overview_markdown(deck_id: str, data: dict[str, Any]) -> str:
    size = data["slide_size"]
    lines = [
        f"# Deck {deck_id}",
        f"{data['slide_count']} slide(s), "
        f"{size['width_in']} x {size['height_in']} in ({size['aspect_ratio']})",
        "",
        "## Masters & layouts",
    ]
    for master in data["masters"]:
        lines.append(f"- **{master['name']}**: {', '.join(master['layouts'])}")
    lines += ["", "## Slides"]
    for slide in data["slides"]:
        title = slide["title"] or "(no title)"
        notes_flag = " · notes" if slide["has_notes"] else ""
        lines.append(
            f"{slide['slide_index']}. {title} — layout: {slide['layout']}{notes_flag}"
        )
    return "\n".join(lines)


def _shape_line(shape: dict[str, Any]) -> str:
    parts = [f"shape_id={shape['shape_id']}"]
    ph = shape["placeholder"]
    if ph:
        parts.append(f"placeholder role={ph['role']} idx={ph['idx']}")
    else:
        parts.append(shape["shape_type"] or "shape")
    geometry = shape["geometry"]
    if geometry:
        inherited = " (from layout)" if geometry.get("inherited_from_layout") else ""
        parts.append(
            f"at {geometry['left_in']},{geometry['top_in']} "
            f"size {geometry['width_in']}x{geometry['height_in']} in{inherited}"
        )
    if "table" in shape:
        parts.append(f"table {shape['table']['rows']}×{shape['table']['columns']}")
    if "chart" in shape:
        parts.append(f"chart {shape['chart']['chart_type']}")
    return f"- **{shape['name']}** ({', '.join(parts)})"


def slide_markdown(data: dict[str, Any]) -> str:
    lines = [
        f"# Slide {data['slide_index']} — {data['title'] or '(no title)'}",
        f"Layout: {data['layout']} (master: {data['master']})",
        "",
        "## Shapes",
    ]
    for shape in data["shapes"]:
        lines.append(_shape_line(shape))
        if shape["text"]:
            for paragraph in shape["text"]:
                if paragraph["text"]:
                    lines.append(f"  {'  ' * paragraph['level']}- {paragraph['text']}")
    if data["notes"]:
        lines += ["", "## Notes", data["notes"]]
    return "\n".join(lines)


def search_markdown(data: dict[str, Any]) -> str:
    if data["total_count"] == 0:
        return f"No matches for '{data['query']}'."
    lines = [
        f"{data['total_count']} match(es) for '{data['query']}' "
        f"(showing {data['count']} from offset {data['offset']}):"
    ]
    for hit in data["hits"]:
        where = "notes" if hit["in"] == "notes" else f"shape {hit['shape_id']} ({hit['shape_name']})"
        lines.append(f"- slide {hit['slide_index']}, {where}: \"{hit['match']}\"")
    if data["has_more"]:
        lines.append(f"More results available: pass offset={data['next_offset']}.")
    return "\n".join(lines)

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


def template_markdown(entry: dict[str, Any]) -> str:
    theme = entry["theme"]
    accents = ", ".join(
        f"{k}={v}" for k, v in theme["color_scheme"].items() if k.startswith("accent")
    )
    lines = [
        f"# Template {entry['template_id']} — {entry['name']} (v{entry['version']})",
        f"Slide size: {entry['slide_size']['width_in']} x {entry['slide_size']['height_in']} in"
        f" · fonts: major={theme['font_scheme'].get('major')}, minor={theme['font_scheme'].get('minor')}",
        f"Accent colors: {accents}",
        "",
        "## Layouts",
    ]
    for layout in entry["layouts"]:
        roles = ", ".join(
            ph["role"] for ph in layout["placeholders"]
            if ph["role"] not in ("footer", "date", "slide_number")
        ) or "(no content placeholders)"
        tags = ", ".join(layout["intent_tags"]) or "-"
        lines.append(
            f"- **{layout['layout_id']}** '{layout['name']}' — placeholders: {roles} · tags: {tags}"
        )
    return "\n".join(lines)


def layout_markdown(template_id: str, layout: dict[str, Any]) -> str:
    lines = [
        f"# Layout {layout['layout_id']} — {layout['name']} (template {template_id})",
        f"Intent tags: {', '.join(layout['intent_tags']) or '-'}",
        f"Capacity (estimated): ~{layout['capacity']['body_bullets']} bullets, "
        f"~{layout['capacity']['chars_per_line']} chars/line",
        "",
        "## Placeholders",
    ]
    for ph in layout["placeholders"]:
        geometry = (
            f"at {ph['left_in']},{ph['top_in']} size {ph['width_in']}x{ph['height_in']} in"
            if ph["left_in"] is not None
            else "(geometry inherited from master)"
        )
        lines.append(f"- idx:{ph['idx']} role={ph['role']} type={ph['type']} {geometry}")
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

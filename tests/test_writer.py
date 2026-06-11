"""M1 authoring tests, driven through the server tool functions so the
dry-run/snapshot/commit flow is exercised, not just the writer internals."""

import pytest

from ppt_mcp import server
from ppt_mcp.errors import PptMcpError
from ppt_mcp.models import (
    ChartSpec,
    ChartSeries,
    ContentSpec,
    ImageRef,
    Paragraph,
    Position,
    ReplaceTextOp,
    AppendParagraphOp,
    ShapeContent,
    TableSpec,
)


@pytest.fixture()
def deck():
    deck_id = server.ppt_create_deck()["deck_id"]
    yield deck_id
    server.ppt_close_deck(deck_id)


@pytest.fixture()
def opened(sample_deck):
    deck_id = server.ppt_open_deck(str(sample_deck))["deck_id"]
    yield deck_id
    server.ppt_close_deck(deck_id)


BULLETS = [
    Paragraph(text="First point"),
    Paragraph(text="Detail", level=1),
    Paragraph(text="Second point"),
]


def test_add_slide_dry_run_does_not_mutate(deck):
    result = server.ppt_add_slide(
        deck, "Title and Content",
        ContentSpec(title="T", body=BULLETS), dry_run=True,
    )
    assert result["applied"] is False
    assert len(result["plan"]["placements"]) == 2
    assert server.ppt_get_deck_overview(deck, response_format="json")["slide_count"] == 0


def test_add_slide_fills_placeholders_and_notes(deck):
    result = server.ppt_add_slide(
        deck, "Title and Content",
        ContentSpec(title="Governance", body=BULLETS, notes="Speak slowly."),
    )
    assert result["applied"] is True
    slide = server.ppt_get_slide(deck, result["slide_index"], response_format="json")
    assert slide["title"] == "Governance"
    body = next(s for s in slide["shapes"] if s["placeholder"]["role"] == "body")
    assert [(p["text"], p["level"]) for p in body["text"]] == [
        ("First point", 0), ("Detail", 1), ("Second point", 0),
    ]
    assert slide["notes"] == "Speak slowly."


def test_add_slide_reports_unplaced_content(deck):
    result = server.ppt_add_slide(
        deck, "Title Only",
        ContentSpec(title="T", table=TableSpec(headers=["a"], rows=[["1"]])),
    )
    assert len(result["plan"]["unplaced_content"]) == 1
    assert "table" in result["plan"]["unplaced_content"][0]


def test_add_slide_image_into_picture_placeholder(deck, sample_image):
    result = server.ppt_add_slide(
        deck, "Picture with Caption",
        ContentSpec(title="Vision", image=ImageRef(path=str(sample_image), alt_text="blue box")),
    )
    assert result["plan"]["unplaced_content"] == []
    slide = server.ppt_get_slide(deck, result["slide_index"], response_format="json")
    assert any(
        s["placeholder"] and s["placeholder"]["role"] == "picture" for s in slide["shapes"]
    )


def test_add_slide_unknown_layout_lists_options(deck):
    with pytest.raises(PptMcpError, match="Title Slide"):
        server.ppt_add_slide(deck, "Nonexistent", ContentSpec(title="x"))


def test_set_placeholder_content_by_role(opened):
    server.ppt_set_placeholder_content(
        opened, 1, "title", ShapeContent(text="New Title")
    )
    assert server.ppt_get_slide(opened, 1, response_format="json")["title"] == "New Title"


def test_set_placeholder_ambiguous_role(deck):
    server.ppt_add_slide(deck, "Two Content", ContentSpec(title="T"))
    with pytest.raises(PptMcpError, match="ambiguous"):
        server.ppt_set_placeholder_content(deck, 1, "body", ShapeContent(text="x"))
    # disambiguating by idx works
    server.ppt_set_placeholder_content(deck, 1, "idx:1", ShapeContent(text="left"))


def test_edit_text_replace_and_append(opened):
    result = server.ppt_edit_text(
        opened, 2, "body",
        [
            ReplaceTextOp(op="replace_text", find="decision", replace="escalation"),
            AppendParagraphOp(op="append_paragraph", text="New bullet", level=0),
        ],
    )
    assert result["replacements"] == 1
    slide = server.ppt_get_slide(opened, 2, response_format="json")
    body = next(s for s in slide["shapes"] if s["placeholder"]["role"] == "body")
    texts = [p["text"] for p in body["text"]]
    assert "No escalation authority" in texts
    assert texts[-1] == "New bullet"


def test_add_table_requires_gate_then_works_freeform(opened):
    spec = TableSpec(headers=["Risk", "Owner"], rows=[["Data", "CTO"], ["Audit", "CISO"]])
    with pytest.raises(PptMcpError, match="allow_freeform"):
        server.ppt_add_table(opened, 3, spec)
    result = server.ppt_add_table(
        opened, 3, spec, allow_freeform=True,
        position=Position(left=1, top=1, width=6, height=2),
    )
    assert result["applied"] is True
    assert "C06" in result["warnings"][0]
    slide = server.ppt_get_slide(opened, 3, response_format="json")
    assert any("table" in s for s in slide["shapes"])


def test_add_chart_freeform(opened):
    result = server.ppt_add_chart(
        opened, 3,
        ChartSpec(chart_type="column", categories=["Q1", "Q2"],
                  series=[ChartSeries(name="Rev", values=[1.0, 2.0])]),
        allow_freeform=True, position=Position(left=1, top=3, width=5, height=3),
    )
    assert result["applied"] is True
    slide = server.ppt_get_slide(opened, 3, response_format="json")
    assert any("chart" in s for s in slide["shapes"])


def test_delete_move_and_undo(opened):
    assert server.ppt_delete_slide(opened, 3)["slide_count"] == 2
    server.ppt_move_slide(opened, 2, 1)
    overview = server.ppt_get_deck_overview(opened, response_format="json")
    assert overview["slides"][0]["title"] == "Key Risks"
    # undo the move, then the delete
    server.ppt_undo(opened, 2)
    overview = server.ppt_get_deck_overview(opened, response_format="json")
    assert overview["slide_count"] == 3
    assert overview["slides"][0]["title"] == "Quarterly Business Review"


def test_duplicate_slide_with_picture(deck, sample_image, tmp_path):
    server.ppt_add_slide(
        deck, "Picture with Caption",
        ContentSpec(title="Original", image=ImageRef(path=str(sample_image)),
                    notes="dup me"),
    )
    result = server.ppt_duplicate_slide(deck, 1)
    assert result["new_slide_index"] == 2
    copy = server.ppt_get_slide(deck, 2, response_format="json")
    assert copy["title"] == "Original"
    assert any(
        s["placeholder"] and s["placeholder"]["role"] == "picture" for s in copy["shapes"]
    )
    assert copy["notes"] == "dup me"
    # the copied image relationship must survive a save/reopen round-trip
    out = tmp_path / "dup.pptx"
    server.ppt_save_deck(deck, str(out))
    reopened = server.ppt_open_deck(str(out))
    slide = server.ppt_get_slide(reopened["deck_id"], 2, response_format="json")
    assert any(
        s["placeholder"] and s["placeholder"]["role"] == "picture" for s in slide["shapes"]
    )
    server.ppt_close_deck(reopened["deck_id"])


def test_m1_exit_criterion_ten_slides_no_freeform(deck, sample_image):
    """M1 exit: a 10-slide deck built only from layouts and placeholders —
    zero absolutely-positioned shapes."""
    img = ImageRef(path=str(sample_image))
    slides = [
        ("Title Slide", ContentSpec(title="AI Strategy", subtitle="Board Edition")),
        ("Title and Content", ContentSpec(title="Agenda", body=BULLETS)),
        ("Section Header", ContentSpec(title="Part 1")),
        ("Title and Content", ContentSpec(title="Findings", body=BULLETS)),
        ("Picture with Caption", ContentSpec(title="Vision", image=img)),
        ("Two Content", ContentSpec(title="Compare", body=BULLETS)),
        ("Title and Content", ContentSpec(title="Risks", body=BULLETS, notes="n")),
        ("Section Header", ContentSpec(title="Part 2")),
        ("Title and Content", ContentSpec(title="Roadmap", body=BULLETS)),
        ("Title Only", ContentSpec(title="Thank You")),
    ]
    for layout, spec in slides:
        result = server.ppt_add_slide(deck, layout, spec)
        assert result["applied"] is True
    overview = server.ppt_get_deck_overview(deck, response_format="json")
    assert overview["slide_count"] == 10
    for i in range(1, 11):
        detail = server.ppt_get_slide(deck, i, response_format="json")
        for shape in detail["shapes"]:
            assert shape["placeholder"] is not None, (
                f"slide {i} shape {shape['name']} is not a placeholder"
            )

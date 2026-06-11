import pytest

from ppt_mcp import reader
from ppt_mcp.errors import SlideIndexError


@pytest.fixture()
def prs(sample_deck):
    return reader.load_presentation(sample_deck)


def test_overview(prs):
    data = reader.deck_overview(prs)
    assert data["slide_count"] == 3
    assert data["slide_size"]["aspect_ratio"] in ("16:9", "4:3")
    assert len(data["masters"]) == 1
    assert "Title Slide" in data["masters"][0]["layouts"]
    assert data["slides"][0]["title"] == "Quarterly Business Review"
    assert data["slides"][1]["has_notes"] is True
    assert data["slides"][2]["title"] is None


def test_slide_detail_roles_and_text(prs):
    data = reader.slide_detail(prs, 2)
    roles = {s["placeholder"]["role"] for s in data["shapes"] if s["placeholder"]}
    assert "title" in roles
    assert "body" in roles
    body = next(s for s in data["shapes"] if s["placeholder"] and s["placeholder"]["role"] == "body")
    texts = [(p["text"], p["level"]) for p in body["text"]]
    assert ("Unclear data ownership", 0) in texts
    assert ("Mitigation: governance board", 1) in texts
    assert data["notes"] == "Stress the audit trail gap."


def test_slide_detail_inherited_geometry(prs):
    data = reader.slide_detail(prs, 1)
    title = next(s for s in data["shapes"] if s["placeholder"] and s["placeholder"]["role"] == "title")
    geometry = title["geometry"]
    assert geometry is not None, "placeholder geometry should resolve via layout"
    assert geometry["width_in"] > 0


def test_slide_detail_freeform_shape(prs):
    data = reader.slide_detail(prs, 3)
    box = next(s for s in data["shapes"] if s["placeholder"] is None)
    assert box["geometry"]["left_in"] == 1.0
    assert box["text"][0]["text"] == "Freeform annotation"


def test_slide_index_bounds(prs):
    with pytest.raises(SlideIndexError, match="1..3"):
        reader.slide_detail(prs, 4)
    with pytest.raises(SlideIndexError):
        reader.slide_detail(prs, 0)


def test_search_hits_shapes_and_notes(prs):
    data = reader.search_deck(prs, "audit")
    assert data["total_count"] == 1
    assert data["hits"][0]["in"] == "notes"

    data = reader.search_deck(prs, "ownership")
    assert data["hits"][0]["slide_index"] == 2
    assert data["hits"][0]["shape_id"] is not None


def test_search_pagination(prs):
    all_hits = reader.search_deck(prs, "e")  # matches many shapes
    assert all_hits["total_count"] > 2
    page = reader.search_deck(prs, "e", limit=2, offset=0)
    assert page["count"] == 2
    assert page["has_more"] is True
    next_page = reader.search_deck(prs, "e", limit=2, offset=page["next_offset"])
    assert next_page["hits"] != page["hits"]

"""M3 render tests — skipped wholesale when LibreOffice is absent."""

import pytest

from ppt_mcp import server
from ppt_mcp.models import ShapeContent
from ppt_mcp.render import find_soffice

pytestmark = pytest.mark.skipif(
    find_soffice() is None, reason="LibreOffice not installed"
)


@pytest.fixture(scope="module")
def deck(tmp_path_factory):
    from conftest import build_sample_deck

    path = build_sample_deck(tmp_path_factory.mktemp("render") / "render.pptx")
    deck_id = server.ppt_open_deck(str(path))["deck_id"]
    yield deck_id
    server.ppt_close_deck(deck_id)


def test_render_deck_produces_pngs(deck):
    result = server.ppt_render_deck(deck, dpi=72)
    assert len(result["slides"]) == 3
    from PIL import Image

    with Image.open(result["slides"]["1"]) as im:
        # 10in x 7.5in at 72 dpi (default template is 4:3)
        assert im.size[0] in (720, 960)  # 4:3 or 16:9
        assert im.size[0] > im.size[1]


def test_render_single_slide_returns_image(deck):
    parts = server.ppt_render_slide(deck, 2, dpi=72)
    from mcp.server.fastmcp import Image as MCPImage

    assert any(isinstance(p, MCPImage) for p in parts)


def test_export_pdf(deck, tmp_path):
    target = tmp_path / "out.pdf"
    result = server.ppt_export_pdf(deck, str(target))
    assert target.is_file()
    assert target.stat().st_size > 1000
    assert result["exported_to"].endswith("out.pdf")


def test_visual_diff_detects_title_change(deck):
    # mutate slide 1 (creates snapshot 1)
    server.ppt_set_placeholder_content(
        deck, 1, "title", ShapeContent(text="COMPLETELY DIFFERENT TITLE")
    )
    diff = server.ppt_visual_diff(deck, snapshot=1, dpi=72)
    assert "1" in diff["changed_slides"]
    assert "2" not in diff["changed_slides"]
    assert "3" not in diff["changed_slides"]

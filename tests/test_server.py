"""Smoke tests for the MCP tool surface. Tool functions are invoked directly
(FastMCP's decorator returns them unchanged); the registered inventory is
checked through the server's own list_tools."""

import anyio

from ppt_mcp import server

EXPECTED_TOOLS = {
    "ppt_open_deck",
    "ppt_create_deck",
    "ppt_save_deck",
    "ppt_close_deck",
    "ppt_list_decks",
    "ppt_undo",
    "ppt_get_deck_overview",
    "ppt_get_slide",
    "ppt_search_deck",
}


def test_tool_inventory_and_annotations():
    tools = anyio.run(server.mcp.list_tools)
    by_name = {t.name: t for t in tools}
    assert EXPECTED_TOOLS <= set(by_name)
    assert by_name["ppt_get_slide"].annotations.readOnlyHint is True
    assert by_name["ppt_close_deck"].annotations.destructiveHint is True
    for tool in by_name.values():
        assert tool.annotations.openWorldHint is False  # local-only guarantee
        assert tool.description, f"{tool.name} is missing a description"


def test_open_read_save_flow(sample_deck, tmp_path):
    opened = server.ppt_open_deck(str(sample_deck))
    deck_id = opened["deck_id"]
    assert opened["overview"]["slide_count"] == 3

    listed = server.ppt_list_decks()
    assert any(d["deck_id"] == deck_id for d in listed["decks"])

    slide_md = server.ppt_get_slide(deck_id, 2)
    assert "Key Risks" in slide_md
    slide_json = server.ppt_get_slide(deck_id, 2, response_format="json")
    assert slide_json["notes"] == "Stress the audit trail gap."

    search_md = server.ppt_search_deck(deck_id, "governance")
    assert "slide 2" in search_md

    saved = server.ppt_save_deck(deck_id, str(tmp_path / "out.pptx"))
    assert (tmp_path / "out.pptx").is_file()
    assert saved["saved_to"].endswith("out.pptx")

    assert server.ppt_undo(deck_id)["undone_steps"] == 0  # no mutations yet

    server.ppt_close_deck(deck_id)
    assert not any(d["deck_id"] == deck_id for d in server.ppt_list_decks()["decks"])


def test_overview_markdown(sample_deck):
    opened = server.ppt_open_deck(str(sample_deck))
    md = server.ppt_get_deck_overview(opened["deck_id"])
    assert "Quarterly Business Review" in md
    assert "Masters & layouts" in md
    server.ppt_close_deck(opened["deck_id"])

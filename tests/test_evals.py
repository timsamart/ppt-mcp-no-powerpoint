"""Guards the eval suite: every answer in evals/questions.xml must match what
the tools actually return for the committed fixture deck."""

import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from ppt_mcp import server

REPO = Path(__file__).parent.parent
FIXTURE = REPO / "evals" / "fixtures" / "eval_deck.pptx"
QUESTIONS = REPO / "evals" / "questions.xml"


@pytest.fixture(scope="module")
def answers():
    root = ET.parse(QUESTIONS).getroot()
    return [qa.findtext("answer") for qa in root.findall("qa_pair")]


@pytest.fixture(scope="module")
def deck():
    deck_id = server.ppt_open_deck(str(FIXTURE))["deck_id"]
    yield deck_id
    server.ppt_close_deck(deck_id)


@pytest.fixture(scope="module")
def template_id(deck):
    return server.ppt_extract_template_from_deck(deck, name="Eval Consistency")["template_id"]


def test_eval_answers_are_consistent(deck, template_id, answers):
    overview = server.ppt_get_deck_overview(deck, response_format="json")
    entry = server.ppt_inspect_template(template_id, response_format="json")
    validation = server.ppt_validate_compliance(deck, template_id)
    c02 = [f for f in validation["findings"] if f["rule"] == "C02"]
    slide2 = server.ppt_get_slide(deck, 2, response_format="json")
    body = next(
        s for s in slide2["shapes"]
        if s["placeholder"] and s["placeholder"]["role"] == "body"
    )
    picture_layouts = [
        lo["name"] for lo in entry["layouts"]
        if any(ph["role"] == "picture" for ph in lo["placeholders"])
    ]
    hits = server.ppt_search_deck(deck, "audit trail", response_format="json")

    computed = [
        overview["slide_size"]["aspect_ratio"],
        overview["slides"][-1]["title"],
        str(len(entry["layouts"])),
        picture_layouts[0],
        entry["theme"]["font_scheme"]["major"],
        entry["theme"]["color_scheme"]["accent1"],
        str(hits["hits"][0]["slide_index"]),
        str(len([p for p in body["text"] if p["text"].strip()])),
        c02[0]["message"].split("'")[1],
        str(c02[0]["slide_index"]),
    ]
    assert computed == answers

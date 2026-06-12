"""Solve every eval question against the fixture deck through the MCP tool
functions and print the ground-truth answers (mcp-builder phase: answer
verification). Run: uv run python scripts/compute_eval_answers.py"""

import os
import tempfile
from pathlib import Path

os.environ.setdefault("PPT_MCP_HOME", str(Path(tempfile.mkdtemp(prefix="ppt-mcp-eval-"))))

from ppt_mcp import server

DECK = Path(__file__).parent.parent / "evals" / "fixtures" / "eval_deck.pptx"


def main() -> None:
    deck = server.ppt_open_deck(str(DECK))["deck_id"]
    overview = server.ppt_get_deck_overview(deck, response_format="json")
    print("aspect_ratio:", overview["slide_size"]["aspect_ratio"])
    print("last_slide_title:", overview["slides"][-1]["title"])

    extracted = server.ppt_extract_template_from_deck(deck, name="Eval Template")
    template_id = extracted["template_id"]
    print("layout_count:", extracted["layouts"])

    entry = server.ppt_inspect_template(template_id, response_format="json")
    picture_layouts = [
        lo["name"] for lo in entry["layouts"]
        if any(ph["role"] == "picture" for ph in lo["placeholders"])
    ]
    print("picture_layouts:", picture_layouts)
    theme = entry["theme"]
    print("major_font:", theme["font_scheme"]["major"])
    print("accent1:", theme["color_scheme"]["accent1"])

    hits = server.ppt_search_deck(deck, "audit trail", response_format="json")
    print("audit_trail_slide:", hits["hits"][0]["slide_index"])

    slide2 = server.ppt_get_slide(deck, 2, response_format="json")
    body = next(
        s for s in slide2["shapes"]
        if s["placeholder"] and s["placeholder"]["role"] == "body"
    )
    print("slide2_bullets:", len([p for p in body["text"] if p["text"].strip()]))

    validation = server.ppt_validate_compliance(deck, template_id)
    c02 = [f for f in validation["findings"] if f["rule"] == "C02"]
    print("c02_count:", len(c02))
    print("c02_messages:", [f["message"] for f in c02])
    print("c02_slide:", [f["slide_index"] for f in c02])

    server.ppt_close_deck(deck)


if __name__ == "__main__":
    main()

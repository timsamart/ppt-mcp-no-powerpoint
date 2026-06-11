"""Build examples/demo.pptx through the MCP tool functions — layout-driven,
zero freeform shapes. Open the result in PowerPoint as the manual
ground-truth check (DESIGN.md §15). Run: uv run python scripts/make_demo_deck.py"""

import os
import tempfile
from pathlib import Path

os.environ.setdefault("PPT_MCP_HOME", str(Path(tempfile.gettempdir()) / "ppt-mcp-demo"))

from ppt_mcp import server
from ppt_mcp.models import ContentSpec, Paragraph


def p(*texts: str) -> list[Paragraph]:
    return [Paragraph(text=t) for t in texts]


RISKS = [
    Paragraph(text="Unclear data ownership"),
    Paragraph(text="Mitigation: governance board", level=1),
    Paragraph(text="No audit trail"),
]

SLIDES = [
    ("Title Slide", ContentSpec(title="ppt-mcp Demo Deck", subtitle="Built without PowerPoint (M1)")),
    ("Title and Content", ContentSpec(title="Agenda", body=p("Why", "What", "How"))),
    ("Section Header", ContentSpec(title="Part 1: Findings")),
    ("Title and Content", ContentSpec(title="Key Risks", body=RISKS, notes="Stress the audit gap.")),
    ("Two Content", ContentSpec(title="Comparison", body=RISKS)),
    ("Section Header", ContentSpec(title="Part 2: Plan")),
    ("Title and Content", ContentSpec(title="Roadmap", body=p("M2: templates", "M3: compliance", "M4: brand styles"))),
    ("Title Only", ContentSpec(title="Thank You")),
]


def main() -> None:
    out = Path(__file__).parent.parent / "examples" / "demo.pptx"
    out.parent.mkdir(exist_ok=True)
    deck = server.ppt_create_deck()["deck_id"]
    for layout, spec in SLIDES:
        result = server.ppt_add_slide(deck, layout, spec)
        assert result["applied"], result
    server.ppt_save_deck(deck, str(out))
    server.ppt_close_deck(deck)
    print(f"saved {out}")


if __name__ == "__main__":
    main()

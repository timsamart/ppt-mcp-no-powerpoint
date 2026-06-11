"""Error types whose messages are written for the calling agent: state the
problem, then the next step that resolves it (DESIGN.md §11.2)."""


class PptMcpError(Exception):
    """Base class for all tool-facing errors."""


class DeckNotFoundError(PptMcpError):
    def __init__(self, deck_id: str, known_ids: list[str]):
        known = ", ".join(known_ids) if known_ids else "none"
        super().__init__(
            f"Unknown deck_id '{deck_id}'. Open decks: {known}. "
            "Use ppt_open_deck or ppt_create_deck first, or ppt_list_decks to see sessions."
        )


class SlideIndexError(PptMcpError):
    def __init__(self, slide_index: int, slide_count: int):
        super().__init__(
            f"slide_index {slide_index} is out of range; the deck has {slide_count} "
            f"slide(s), indexed 1..{slide_count}."
        )

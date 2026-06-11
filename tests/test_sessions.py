from pathlib import Path

import pytest

from ppt_mcp.errors import DeckNotFoundError, PptMcpError
from ppt_mcp.sessions import SessionManager
from ppt_mcp.store import Store


def test_open_copies_to_working_and_leaves_source_untouched(manager, sample_deck):
    before = sample_deck.read_bytes()
    session = manager.open_deck(sample_deck)
    assert session.working_path.is_file()
    assert session.working_path != sample_deck
    assert sample_deck.read_bytes() == before


def test_open_rejects_missing_and_wrong_type(manager, tmp_path):
    with pytest.raises(PptMcpError, match="File not found"):
        manager.open_deck(tmp_path / "nope.pptx")
    bad = tmp_path / "doc.docx"
    bad.write_bytes(b"x")
    with pytest.raises(PptMcpError, match="Unsupported file type"):
        manager.open_deck(bad)


def test_unknown_deck_id_error_lists_open_decks(manager, sample_deck):
    session = manager.open_deck(sample_deck)
    with pytest.raises(DeckNotFoundError, match=session.deck_id):
        manager.get("deck_bogus")


def test_save_defaults_to_source(manager, sample_deck):
    session = manager.open_deck(sample_deck)
    session.working_path.write_bytes(session.working_path.read_bytes() + b" ")
    target = manager.save_deck(session.deck_id)
    assert target == sample_deck
    assert sample_deck.read_bytes() == session.working_path.read_bytes()


def test_save_created_deck_requires_path(manager, tmp_path):
    session = manager.create_deck()
    with pytest.raises(PptMcpError, match="explicit target path"):
        manager.save_deck(session.deck_id)
    out = manager.save_deck(session.deck_id, tmp_path / "new.pptx")
    assert out.is_file()


def test_save_refuses_registered_template_dir(manager, store, sample_deck):
    session = manager.open_deck(sample_deck)
    protected = store.templates_dir / "tpl_x" / "source.potx"
    with pytest.raises(PptMcpError, match="immutable"):
        manager.save_deck(session.deck_id, protected)


def test_snapshot_undo_roundtrip(manager, sample_deck):
    session = manager.open_deck(sample_deck)
    original = session.working_path.read_bytes()

    manager.snapshot(session.deck_id)
    session.working_path.write_bytes(original + b"1")
    manager.snapshot(session.deck_id)
    session.working_path.write_bytes(original + b"22")

    assert manager.undo(session.deck_id, 1) == 1
    assert session.working_path.read_bytes() == original + b"1"
    assert manager.undo(session.deck_id, 5) == 1  # clamped to what exists
    assert session.working_path.read_bytes() == original
    assert manager.undo(session.deck_id) == 0  # nothing left


def test_sessions_persist_across_restart(store, sample_deck):
    first = SessionManager(store)
    session = first.open_deck(sample_deck)

    second = SessionManager(store)  # simulated restart
    restored = second.get(session.deck_id)
    assert restored.source_path == session.source_path
    assert restored.working_path.is_file()


def test_close_removes_session_dir(manager, sample_deck):
    session = manager.open_deck(sample_deck)
    manager.close_deck(session.deck_id)
    assert not session.session_dir.exists()
    with pytest.raises(DeckNotFoundError):
        manager.get(session.deck_id)


def test_provenance_log_written(store: Store, manager, sample_deck):
    manager.open_deck(sample_deck)
    text = store.provenance_path.read_text(encoding="utf-8")
    assert "ppt_open_deck" in text

"""Deck session manager (DESIGN.md §4.1).

Decks are addressed by opaque `deck_id` handles. Opening a deck copies it
into a session working directory; all edits hit the working copy and only
`save_deck` writes back. Every mutation snapshots first, so undo is a file
restore. Sessions survive server restarts via per-session manifest files.
"""

from __future__ import annotations

import json
import shutil
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from .errors import DeckNotFoundError, PptMcpError
from .store import Store

DECK_SUFFIXES = {".pptx", ".potx"}
MANIFEST_NAME = "manifest.json"
WORKING_NAME = "working.pptx"
SNAPSHOTS_DIRNAME = "snapshots"


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


@dataclass
class DeckSession:
    deck_id: str
    session_dir: Path
    source_path: Path | None  # None for decks created from scratch
    opened_at: str
    snapshot_count: int = 0
    closed: bool = field(default=False, repr=False)

    @property
    def working_path(self) -> Path:
        return self.session_dir / WORKING_NAME

    @property
    def snapshots_dir(self) -> Path:
        return self.session_dir / SNAPSHOTS_DIRNAME

    def manifest(self) -> dict:
        return {
            "deck_id": self.deck_id,
            "source_path": str(self.source_path) if self.source_path else None,
            "opened_at": self.opened_at,
            "snapshot_count": self.snapshot_count,
        }

    def write_manifest(self) -> None:
        path = self.session_dir / MANIFEST_NAME
        path.write_text(json.dumps(self.manifest(), indent=2), encoding="utf-8")


class SessionManager:
    def __init__(self, store: Store):
        self.store = store
        self._sessions: dict[str, DeckSession] = {}
        self._restore_persisted_sessions()

    # -- lifecycle -----------------------------------------------------------

    def open_deck(self, path: str | Path) -> DeckSession:
        source = Path(path).expanduser().resolve()
        if not source.is_file():
            raise PptMcpError(
                f"File not found: '{source}'. Provide an absolute path to an "
                "existing .pptx or .potx file."
            )
        if source.suffix.lower() not in DECK_SUFFIXES:
            raise PptMcpError(
                f"Unsupported file type '{source.suffix}'. Supported: "
                f"{', '.join(sorted(DECK_SUFFIXES))}."
            )
        session = self._new_session(source_path=source)
        shutil.copy2(source, session.working_path)
        session.write_manifest()
        self.store.log_provenance("ppt_open_deck", deck_id=session.deck_id, source=source)
        return session

    def create_deck(self) -> DeckSession:
        from pptx import Presentation  # deferred: keeps module import light

        session = self._new_session(source_path=None)
        Presentation().save(str(session.working_path))
        session.write_manifest()
        self.store.log_provenance("ppt_create_deck", deck_id=session.deck_id)
        return session

    def save_deck(self, deck_id: str, path: str | Path | None = None) -> Path:
        session = self.get(deck_id)
        if path is not None:
            target = Path(path).expanduser().resolve()
        elif session.source_path is not None:
            target = session.source_path
        else:
            raise PptMcpError(
                f"Deck '{deck_id}' was created from scratch and has no source path. "
                "Call ppt_save_deck with an explicit target path ending in .pptx."
            )
        if target.suffix.lower() not in DECK_SUFFIXES:
            raise PptMcpError(
                f"Target path must end in .pptx or .potx, got '{target.name}'."
            )
        if self.store.is_protected_path(target):
            raise PptMcpError(
                "Refusing to overwrite a registered template source — registered "
                "templates are immutable. Save to a different path."
            )
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(session.working_path, target)
        self.store.log_provenance("ppt_save_deck", deck_id=deck_id, target=target)
        return target

    def close_deck(self, deck_id: str) -> None:
        session = self.get(deck_id)
        session.closed = True
        del self._sessions[deck_id]
        shutil.rmtree(session.session_dir, ignore_errors=True)
        self.store.log_provenance("ppt_close_deck", deck_id=deck_id)

    def list_sessions(self) -> list[DeckSession]:
        return sorted(self._sessions.values(), key=lambda s: s.opened_at)

    def get(self, deck_id: str) -> DeckSession:
        session = self._sessions.get(deck_id)
        if session is None:
            raise DeckNotFoundError(deck_id, list(self._sessions))
        return session

    # -- snapshots / undo ------------------------------------------------------

    def snapshot(self, deck_id: str) -> int:
        """Snapshot the working file before a mutation. Returns snapshot number."""
        session = self.get(deck_id)
        session.snapshots_dir.mkdir(exist_ok=True)
        session.snapshot_count += 1
        shutil.copy2(
            session.working_path,
            session.snapshots_dir / f"{session.snapshot_count}.pptx",
        )
        session.write_manifest()
        return session.snapshot_count

    def undo(self, deck_id: str, steps: int = 1) -> int:
        """Restore the working file from `steps` snapshots back. Returns the
        number of steps actually undone."""
        session = self.get(deck_id)
        if steps < 1:
            raise PptMcpError("steps must be >= 1.")
        available = session.snapshot_count
        if available == 0:
            return 0
        steps = min(steps, available)
        restore_from = session.snapshots_dir / f"{available - steps + 1}.pptx"
        shutil.copy2(restore_from, session.working_path)
        for n in range(available - steps + 1, available + 1):
            (session.snapshots_dir / f"{n}.pptx").unlink(missing_ok=True)
        session.snapshot_count = available - steps
        session.write_manifest()
        self.store.log_provenance("ppt_undo", deck_id=deck_id, steps=steps)
        return steps

    # -- internals -------------------------------------------------------------

    def _new_session(self, source_path: Path | None) -> DeckSession:
        deck_id = f"deck_{uuid.uuid4().hex[:8]}"
        session_dir = self.store.sessions_dir / deck_id
        session_dir.mkdir(parents=True)
        session = DeckSession(
            deck_id=deck_id,
            session_dir=session_dir,
            source_path=source_path,
            opened_at=_now_iso(),
        )
        self._sessions[deck_id] = session
        return session

    def _restore_persisted_sessions(self) -> None:
        for manifest_path in self.store.sessions_dir.glob(f"*/{MANIFEST_NAME}"):
            try:
                data = json.loads(manifest_path.read_text(encoding="utf-8"))
                session = DeckSession(
                    deck_id=data["deck_id"],
                    session_dir=manifest_path.parent,
                    source_path=Path(data["source_path"]) if data.get("source_path") else None,
                    opened_at=data.get("opened_at", _now_iso()),
                    snapshot_count=int(data.get("snapshot_count", 0)),
                )
                if session.working_path.is_file():
                    self._sessions[session.deck_id] = session
            except (json.JSONDecodeError, KeyError, OSError):
                continue  # a broken session dir must not take the server down

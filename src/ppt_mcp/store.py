"""Local storage layout (DESIGN.md §12).

Everything lives under one inspectable directory: `~/.ppt-mcp` by default,
overridable with the PPT_MCP_HOME environment variable. No database, no
network — plain files only.
"""

from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path

ENV_HOME = "PPT_MCP_HOME"


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class Store:
    def __init__(self, root: str | Path | None = None):
        self.root = Path(root or os.environ.get(ENV_HOME) or Path.home() / ".ppt-mcp")
        self.sessions_dir = self.root / "sessions"
        self.registry_dir = self.root / "registry"
        self.templates_dir = self.root / "templates"
        self.style_profiles_dir = self.root / "style_profiles"
        self.patterns_dir = self.root / "patterns"
        self.assets_dir = self.root / "assets"
        self.renders_dir = self.root / "renders"
        self.logs_dir = self.root / "logs"
        for d in (
            self.sessions_dir,
            self.registry_dir,
            self.templates_dir,
            self.style_profiles_dir,
            self.patterns_dir,
            self.assets_dir,
            self.renders_dir,
            self.logs_dir,
        ):
            d.mkdir(parents=True, exist_ok=True)

    @property
    def provenance_path(self) -> Path:
        return self.logs_dir / "provenance.jsonl"

    def log_provenance(self, tool: str, **fields: object) -> None:
        """Append one provenance record (DESIGN.md §13). Never raises."""
        record = {"at": _now_iso(), "tool": tool, **fields}
        try:
            with self.provenance_path.open("a", encoding="utf-8") as fh:
                fh.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")
        except OSError:
            pass

    def is_protected_path(self, path: Path) -> bool:
        """Registered template sources are immutable (DESIGN.md §13)."""
        try:
            path.resolve().relative_to(self.templates_dir.resolve())
            return True
        except ValueError:
            return False

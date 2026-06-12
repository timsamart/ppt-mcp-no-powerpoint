"""Brand Style Prompt Registry (DESIGN.md §4.5, §6.2).

Style profiles are local JSON files describing a corporate visual language
for imagery. They are **configuration, not instruction**: profile text is
interpolated into output prompt strings only — it never drives server
behavior, never executes, never enables network access.

Prompt composition is deterministic template assembly; the calling agent may
refine the scene clause, and every write is re-validated against the
profile's forbidden-motif and color rules.
"""

from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from typing import Any

from .errors import PptMcpError
from .store import Store

STANDARD_NEGATIVE = (
    "no readable text, no fake brand logos, no distorted hands, no glowing "
    "brains, no neon cyberpunk aesthetic, no generic handshake stock photo"
)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")


def _bump_version(version: str) -> str:
    parts = version.split(".")
    try:
        parts[-1] = str(int(parts[-1]) + 1)
        return ".".join(parts)
    except ValueError:
        return version + ".1"


class StyleProfileRegistry:
    def __init__(self, store: Store):
        self.store = store
        self.dir = store.style_profiles_dir

    def _path(self, profile_id: str):
        return self.dir / f"{profile_id}.json"

    def set(self, name: str, system_prompt: str, metadata: dict[str, Any] | None = None) -> dict:
        """Create or update (upsert): updating bumps the version and keeps a
        change history so generated prompts stay traceable (§13)."""
        metadata = metadata or {}
        profile_id = f"sp_{_slug(name)}"
        path = self._path(profile_id)
        now = _now_iso()
        known_fields = {
            "allowed_colors": metadata.get("allowed_colors", []),
            "forbidden_motifs": metadata.get("forbidden_motifs", []),
            "negative_prompt_base": metadata.get("negative_prompt_base", STANDARD_NEGATIVE),
            "media_type": metadata.get("media_type", "photography"),
            "composition": metadata.get("composition"),
            "text_in_image": metadata.get("text_in_image", "forbidden"),
            "logo_usage": metadata.get("logo_usage", "forbidden"),
            "default_aspect_ratio": metadata.get("default_aspect_ratio", "16:9"),
        }
        if path.is_file():
            profile = json.loads(path.read_text(encoding="utf-8"))
            profile["history"].append(
                {"version": profile["version"], "updated_at": profile["updated_at"]}
            )
            profile.update(
                name=name,
                system_prompt=system_prompt,
                version=_bump_version(profile["version"]),
                updated_at=now,
                **known_fields,
            )
        else:
            profile = {
                "profile_id": profile_id,
                "name": name,
                "version": "1.0",
                "system_prompt": system_prompt,
                "created_at": now,
                "updated_at": now,
                "history": [],
                **known_fields,
            }
        path.write_text(json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")
        self.store.log_provenance(
            "ppt_set_style_profile", profile_id=profile_id, version=profile["version"]
        )
        return profile

    def get(self, profile_id: str) -> dict:
        path = self._path(profile_id)
        if not path.is_file():
            known = ", ".join(p.stem for p in self.dir.glob("sp_*.json")) or "none"
            raise PptMcpError(
                f"Unknown profile_id '{profile_id}'. Profiles: {known}. "
                "Create one with ppt_set_style_profile."
            )
        return json.loads(path.read_text(encoding="utf-8"))

    def list(self) -> list[dict]:
        return [
            json.loads(p.read_text(encoding="utf-8"))
            for p in sorted(self.dir.glob("sp_*.json"))
        ]

    def delete(self, profile_id: str) -> None:
        self.get(profile_id)  # existence check with actionable error
        self._path(profile_id).unlink()
        self.store.log_provenance("ppt_delete_style_profile", profile_id=profile_id)

    # -- validation & composition --------------------------------------------------

    @staticmethod
    def validate_prompt(profile: dict, prompt: str) -> list[str]:
        """Check a prompt against the profile's rules. Returns violations."""
        violations = []
        lowered = prompt.lower()
        for motif in profile.get("forbidden_motifs", []):
            if motif.lower() in lowered:
                violations.append(f"forbidden motif: '{motif}'")
        allowed = {c.upper() for c in profile.get("allowed_colors", [])}
        if allowed:
            for hex_color in set(re.findall(r"#[0-9a-fA-F]{6}", prompt)):
                if hex_color.upper() not in allowed:
                    violations.append(
                        f"color {hex_color} is not in the profile's allowed colors"
                    )
        return violations

    @staticmethod
    def compose_prompt(
        profile: dict,
        image_intent: str,
        slide_context: dict[str, Any],
        aspect_ratio: str | None = None,
        constraints: str | None = None,
    ) -> dict[str, Any]:
        """Deterministic prompt assembly — no LLM (DESIGN.md §6.1)."""
        parts = [profile["system_prompt"].strip().rstrip(".") + "."]
        parts.append(f"Scene: {image_intent.strip().rstrip('.')}.")
        title = slide_context.get("title")
        if title:
            parts.append(f"The slide is titled '{title}'.")
        body_summary = slide_context.get("body_summary")
        if body_summary:
            parts.append(f"Key points on the slide: {body_summary}.")
        if profile.get("composition"):
            parts.append(f"Composition: {profile['composition']}.")
        if profile.get("allowed_colors"):
            parts.append(f"Subtle accent colors {', '.join(profile['allowed_colors'])}.")
        if constraints:
            parts.append(constraints.strip().rstrip(".") + ".")
        prompt = " ".join(parts)
        alt_text = image_intent.strip()
        if title:
            alt_text += f" — illustrating '{title}'"
        return {
            "prompt": prompt,
            "negative_prompt": profile.get("negative_prompt_base", STANDARD_NEGATIVE),
            "alt_text": alt_text,
            "aspect_ratio": aspect_ratio or profile.get("default_aspect_ratio", "16:9"),
            "provenance": {
                "style_profile_id": profile["profile_id"],
                "style_profile_version": profile["version"],
                "created_by_tool": "ppt_compose_image_prompt",
                "created_at": _now_iso(),
            },
        }

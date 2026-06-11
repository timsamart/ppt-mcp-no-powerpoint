"""Layout recommendation (DESIGN.md §5.3) — deterministic scoring, no LLM.

score = w_name · name/intent match        (fuzzy vs layout name + intent tags)
      + w_fit  · structural fit           (spec sections ↔ available roles)
      + w_cap  · capacity fit             (bullet count vs layout capacity)
      − w_waste · waste penalty           (content placeholders left empty)

Results carry reasons; the calling agent makes the final call.
"""

from __future__ import annotations

from typing import Any

from rapidfuzz import fuzz

from .models import ContentSpec

W_NAME, W_FIT, W_CAP, W_WASTE = 0.40, 0.40, 0.10, 0.10

_CONTENT_ROLES = ("title", "subtitle", "body", "table", "picture", "chart")


# A direct layout-name hit beats an inferred intent-tag hit: tags are partly
# structural guesses, the name is the template author's own label.
_TAG_DISCOUNT = 0.9


def _name_score(intent: str, layout: dict[str, Any]) -> tuple[float, str | None]:
    intent_lower = intent.lower()
    best = fuzz.partial_ratio(intent_lower, layout["name"].lower()) / 100
    reason = f"name '{layout['name']}' matches intent" if best >= 0.7 else None
    for tag in layout["intent_tags"]:
        tag_score = _TAG_DISCOUNT * fuzz.partial_ratio(intent_lower, tag.replace("_", " ")) / 100
        if tag_score > best:
            best = tag_score
            reason = f"intent tag '{tag}' matches"
    return best, reason


def _sections(spec: ContentSpec) -> dict[str, str]:
    """Spec section -> required placeholder role."""
    wanted = {}
    if spec.title is not None:
        wanted["title"] = "title"
    if spec.subtitle is not None:
        wanted["subtitle"] = "subtitle"
    if spec.body is not None:
        wanted["body"] = "body"
    if spec.table is not None:
        wanted["table"] = "table"
    if spec.image is not None:
        wanted["image"] = "picture"
    if spec.chart is not None:
        wanted["chart"] = "chart"
    return wanted


def score_layout(layout: dict[str, Any], intent: str, spec: ContentSpec) -> dict[str, Any]:
    roles_available: list[str] = [
        ph["role"]
        for ph in layout["placeholders"]
        if ph["role"] in _CONTENT_ROLES
    ]
    wanted = _sections(spec)

    name_score, name_reason = _name_score(intent, layout)

    pool = list(roles_available)
    placed, missing = [], []
    for section, role in wanted.items():
        if role in pool:
            pool.remove(role)
            placed.append(section)
        else:
            missing.append(section)
    fit_score = len(placed) / len(wanted) if wanted else 1.0

    cap_score = 1.0
    cap_reason = None
    if spec.body is not None and layout["capacity"]["body_bullets"]:
        max_bullets = layout["capacity"]["body_bullets"]
        if len(spec.body) > max_bullets:
            cap_score = max(0.0, 1 - (len(spec.body) - max_bullets) / max_bullets)
            cap_reason = f"{len(spec.body)} bullets exceed estimated capacity ({max_bullets})"

    waste = len(pool) / len(roles_available) if roles_available else 0.0

    score = W_NAME * name_score + W_FIT * fit_score + W_CAP * cap_score - W_WASTE * waste

    reasons = []
    if name_reason:
        reasons.append(name_reason)
    if placed:
        reasons.append(f"fits {', '.join(placed)}")
    if missing:
        reasons.append(f"no placeholder for {', '.join(missing)}")
    if cap_reason:
        reasons.append(cap_reason)
    if pool:
        reasons.append(f"{len(pool)} placeholder(s) would stay empty")

    return {
        "layout_id": layout["layout_id"],
        "name": layout["name"],
        "confidence": round(max(0.0, min(1.0, score)), 2),
        "reason": "; ".join(reasons) or "generic fallback",
        "missing_sections": missing,
    }


def recommend_layouts(
    layouts: list[dict[str, Any]],
    intent: str,
    spec: ContentSpec,
    top_n: int = 3,
) -> list[dict[str, Any]]:
    scored = [score_layout(layout, intent, spec) for layout in layouts]
    scored.sort(key=lambda r: r["confidence"], reverse=True)
    return scored[:top_n]

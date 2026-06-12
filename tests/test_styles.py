"""M4 tests: style profiles, prompt composition, governed image placeholders,
and the in-file manifest's save/reopen survival."""

import pytest

from ppt_mcp import server
from ppt_mcp.errors import PptMcpError
from ppt_mcp.models import ContentSpec, ImagePlaceholderSpec, Paragraph, Position

PROFILE_META = {
    "allowed_colors": ["#1A6BCC", "#E0A030"],
    "forbidden_motifs": ["robot hands", "neon cyberpunk"],
    "composition": "circular collaboration motifs, natural light",
}


@pytest.fixture()
def profile_id():
    result = server.ppt_set_style_profile(
        "ACME Corporate Photography",
        "Photorealistic, modern enterprise, calm confidence, diverse professionals.",
        PROFILE_META,
    )
    yield result["profile_id"]
    try:
        server.ppt_delete_style_profile(result["profile_id"])
    except PptMcpError:
        pass


@pytest.fixture()
def deck(sample_deck):
    deck_id = server.ppt_open_deck(str(sample_deck))["deck_id"]
    yield deck_id
    server.ppt_close_deck(deck_id)


def test_profile_upsert_bumps_version(profile_id):
    profile = server.ppt_get_style_profile(profile_id)
    assert profile["version"] == "1.0"
    server.ppt_set_style_profile(
        "ACME Corporate Photography", "Updated visual language.", PROFILE_META
    )
    updated = server.ppt_get_style_profile(profile_id)
    assert updated["version"] == "1.1"
    assert updated["history"][0]["version"] == "1.0"
    assert updated["system_prompt"] == "Updated visual language."


def test_compose_prompt_uses_slide_context(deck, profile_id):
    bundle = server.ppt_compose_image_prompt(
        deck, 2, "leadership team reviewing a strategy wall", profile_id
    )
    prompt = bundle["prompt"]
    assert "Photorealistic" in prompt                 # profile system prompt
    assert "leadership team" in prompt                # intent
    assert "Key Risks" in prompt                      # slide title context
    assert "#1A6BCC" in prompt                        # allowed colors
    assert "no readable text" in bundle["negative_prompt"]
    assert bundle["provenance"]["style_profile_id"] == profile_id
    assert bundle["provenance"]["style_profile_version"] == "1.0"


def test_forbidden_motif_rejected(deck, profile_id):
    with pytest.raises(PptMcpError, match="robot hands"):
        server.ppt_create_image_placeholder(
            deck, 1,
            ImagePlaceholderSpec(
                image_intent="x", profile_id=profile_id,
                prompt="cool scene with robot hands",
                allow_freeform=True,
                position=Position(left=1, top=1, width=4, height=3),
            ),
        )


def test_picture_placeholder_slot_and_fill(profile_id, sample_image):
    deck = server.ppt_create_deck()["deck_id"]
    try:
        server.ppt_add_slide(deck, "Picture with Caption", ContentSpec(title="Vision"))
        created = server.ppt_create_image_placeholder(
            deck, 1,
            ImagePlaceholderSpec(
                image_intent="vision scene", profile_id=profile_id, shape_ref="picture"
            ),
        )
        assert created["applied"] is True
        record = created["record"]
        assert record["status"] == "pending"
        assert record["target_kind"] == "picture_placeholder"
        assert record["style_profile_id"] == profile_id
        # notes pointer, not on-slide prompt
        slide = server.ppt_get_slide(deck, 1, response_format="json")
        assert "vision scene" in slide["notes"]
        assert all(
            "Photorealistic" not in (p["text"] if s["text"] else "")
            for s in slide["shapes"] for p in (s["text"] or [])
        )
        # fill with the generated image
        filled = server.ppt_fill_image_placeholder(
            deck, 1, str(record["shape_id"]), str(sample_image)
        )
        assert filled["status"] == "generated"
        listed = server.ppt_list_image_placeholders(deck, status="generated")
        assert listed["count"] == 1
        assert listed["image_placeholders"][0]["image_path"] == str(sample_image)
    finally:
        server.ppt_close_deck(deck)


def test_freeform_slot_with_label(deck, profile_id):
    created = server.ppt_create_image_placeholder(
        deck, 3,
        ImagePlaceholderSpec(
            image_intent="collaboration scene", profile_id=profile_id,
            allow_freeform=True, position=Position(left=1, top=2, width=5, height=3),
        ),
    )
    slide = server.ppt_get_slide(deck, 3, response_format="json")
    label_shape = next(s for s in slide["shapes"] if s["shape_id"] == created["shape_id"])
    assert label_shape["text"][0]["text"] == "Image placeholder: collaboration scene"


def test_manifest_survives_save_and_reopen(deck, profile_id, tmp_path):
    server.ppt_create_image_placeholder(
        deck, 1,
        ImagePlaceholderSpec(
            image_intent="hero image", profile_id=profile_id,
            allow_freeform=True, position=Position(left=1, top=1, width=4, height=3),
        ),
    )
    out = tmp_path / "roundtrip.pptx"
    server.ppt_save_deck(deck, str(out))
    reopened = server.ppt_open_deck(str(out))["deck_id"]
    try:
        listed = server.ppt_list_image_placeholders(reopened)
        assert listed["count"] == 1
        record = listed["image_placeholders"][0]
        assert record["image_intent"] == "hero image"
        assert record["prompt"].startswith("Photorealistic")
    finally:
        server.ppt_close_deck(reopened)


def test_update_placeholder_prompt_validated(deck, profile_id):
    created = server.ppt_create_image_placeholder(
        deck, 1,
        ImagePlaceholderSpec(
            image_intent="scene", profile_id=profile_id,
            allow_freeform=True, position=Position(left=1, top=1, width=4, height=3),
        ),
    )
    ref = str(created["shape_id"])
    with pytest.raises(PptMcpError, match="neon cyberpunk"):
        server.ppt_update_image_placeholder(
            deck, 1, ref, {"prompt": "scene in neon cyberpunk style"}
        )
    updated = server.ppt_update_image_placeholder(
        deck, 1, ref, {"prompt": "calm modern office scene", "status": "approved"}
    )
    assert updated["record"]["status"] == "approved"


def test_missing_record_error_is_actionable(deck):
    with pytest.raises(PptMcpError, match="ppt_list_image_placeholders"):
        server.ppt_update_image_placeholder(deck, 1, "title", {"status": "approved"})

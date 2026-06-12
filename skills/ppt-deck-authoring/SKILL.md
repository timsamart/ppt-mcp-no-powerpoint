---
name: ppt-deck-authoring
description: Workflow guidance for building template-faithful PowerPoint decks with the ppt-mcp server (tools prefixed ppt_). Use whenever creating or editing .pptx files through ppt-mcp — covers template-first deck creation, layout choice, placeholder filling, and verification.
---

# Building decks with ppt-mcp

ppt-mcp is a local, offline MCP server for PowerPoint. Its core principle:
**slides are born from template layouts and filled through placeholders** —
never assembled from absolutely-positioned textboxes.

## The template-first decision tree

1. **A registered corporate template exists?** (`ppt_list_templates`)
   → `ppt_create_deck(template_id=...)` and use its layouts. Always preferred.
2. **The user hands you a corporate .potx/.pptx?**
   → `ppt_register_template(path)` first, then as above.
3. **No template at all?**
   → `ppt_create_deck()` gives a bare default deck. Say so — output will not
   look corporate.

## Per-slide workflow

1. Describe the slide's content as a `content` spec: `title`, `subtitle`,
   `body` (paragraphs with `level`), `table`, `image`, `chart`, `notes`.
2. Ask `ppt_recommend_layout(template_id, slide_intent, content)` — it returns
   ranked layouts with reasons. Treat it as advice; check the reasons.
3. Optionally preview with `ppt_map_content_to_placeholders`.
4. `ppt_add_slide(deck_id, layout, content)` — it fills placeholders by
   semantic role. **Check `unplaced_content` in the response**: content that
   fits no placeholder is reported, not silently dropped. Pick a better layout
   rather than forcing freeform shapes.
5. Pass `dry_run=true` on any mutating tool to see the change plan first.

## Editing existing decks

- `ppt_open_deck(path)` works on a copy; nothing touches the source until
  `ppt_save_deck`. `ppt_undo` rolls back mutations.
- Inspect before editing: `ppt_get_deck_overview`, then `ppt_get_slide` —
  every shape is addressable by role (`"title"`), placeholder idx
  (`"idx:1"`), or `shape_id`.
- Prefer `ppt_set_placeholder_content` / `ppt_edit_text` over rebuilding
  slides. Keep slides attached to their existing layouts.

## Corporate imagery

Decks should not get random stock photos. The governed path:

1. Store the corporate visual language once:
   `ppt_set_style_profile(name, system_prompt, metadata)` (allowed colors,
   forbidden motifs, composition rules).
2. Per image: `ppt_compose_image_prompt(deck, slide, intent, profile_id)` —
   refine the scene wording if needed.
3. `ppt_create_image_placeholder` — targets a picture placeholder (preferred);
   the full prompt travels inside the file (manifest + notes pointer), never
   on the visible slide.
4. Generate the image with whatever tool the user uses, then
   `ppt_fill_image_placeholder(deck, slide, shape_ref, image_path)`.

## What to avoid

- `allow_freeform=true` exists for table/image/chart placement but bypasses
  the template's grammar (compliance rule C06). Use it only when no layout
  has a suitable placeholder and say so in your summary.
- Never overwrite a registered template; save decks to new paths when in doubt.

## Verification — always close the loop

After building or editing:

1. `ppt_validate_compliance(deck_id, template_id)` — fix errors before
   delivering. `ppt_repair_compliance` auto-fixes font/color drift (preview
   with the default dry_run first, then `dry_run=false`).
2. `ppt_render_slide(deck_id, n)` — look at the slides you changed.
3. After risky edits, `ppt_visual_diff(deck_id, snapshot)` confirms nothing
   else changed (logos, footers).
4. `ppt_export_pdf` when the user wants a shareable artifact.

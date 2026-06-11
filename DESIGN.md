# ppt-mcp-no-powerpoint — Design Document

**Local PowerPoint Intelligence MCP for template-faithful corporate decks.**

A Model Context Protocol server that creates, edits, and validates `.pptx` files
**without depending on PowerPoint** — no COM automation, no Office requirement,
no cloud, no network. Everything runs offline on local files. (If PowerPoint
happens to be installed, it serves one role only: a ground-truth viewer for
humans to inspect outputs. The server never touches it programmatically.)
Its differentiator is not "generate slides"; it is **governed slide production**:
slides that look like they were born inside the company's template, not smuggled
in through customs.

- Status: Draft v1
- Date: 2026-06-11
- Repo: `ppt-mcp-no-powerpoint`

---

## 1. Product thesis

For corporate PowerPoint, template/theme/master support is the core requirement,
not an enhancement. PowerPoint's own file format (PresentationML) places slide
masters, slide layouts, and themes at the structural heart of every valid
document. A corporate template is a **design system encoded as a file**:

```text
Corporate template
├─ theme            (color scheme, font scheme, effects)
├─ slide masters    (logos, footers, brand structure, recurring elements)
├─ slide layouts    (title, section divider, two-column, chart, timeline, …)
└─ placeholders     (title, subtitle, body, picture, chart, table, footer, …)
```

The server therefore treats every operation as template-first:

1. Inspect the approved template.
2. Choose the best layout for each slide intent.
3. Fill placeholders by semantic role — never raw coordinates when a
   placeholder exists.
4. Preserve theme and master inheritance.
5. Render and validate.
6. Report deviations.

### Goals

- **G1 — Template intelligence.** Import, parse, and obey corporate templates
  (`.potx`, `.pptx`, optionally `.thmx`): masters, layouts, placeholders,
  themes, footers, logos.
- **G2 — Placeholder-native authoring.** Map semantic content
  (title/body/table/chart/image) onto layout placeholders instead of absolute
  positioning.
- **G3 — Compliance validation & repair.** Detect and optionally fix drift from
  the approved template: wrong fonts, off-theme colors, bypassed placeholders,
  text overflow, lost footers.
- **G4 — Brand style profiles for imagery.** Store corporate visual-style
  prompts locally; emit governed image *placeholders with generation prompts*
  rather than generating images.
- **G5 — Verification without Office.** Render decks to PNG/PDF via headless
  LibreOffice for visual checks; degrade gracefully when it is absent.
- **G6 — Offline and local, always.** All state on disk in inspectable files.
  The server makes **no network calls, ever** — this is a hard guarantee, not
  a default setting. All processing happens on the local machine.

### Non-goals

- **No image generation.** The server produces prompts and placeholder shapes;
  an external tool/model generates pixels. (A fill tool inserts the finished
  file.)
- **No PowerPoint COM / Office automation.** That is the repo's name and its
  contract. A locally installed PowerPoint may be used by *humans* (or a
  separate computer-use agent) to visually inspect outputs — the server itself
  never invokes, scripts, or depends on it.
- **No cloud rendering, no telemetry, no network access of any kind.**
- **No in-server LLM.** The server is deterministic; semantic judgment lives in
  the calling agent (see §6.1). MCP *sampling* is a future option, not a
  dependency.
- **No real-time co-editing / OneDrive sync.**

---

## 2. Technology decisions

### 2.1 Language & core library: Python + python-pptx + lxml

| Option | Verdict | Why |
|---|---|---|
| **python-pptx + lxml** | ✅ core | Only mature OSS library that *reads and edits existing* `.pptx`/`.potx` including masters, layouts, placeholders, and theme parts. Explicitly models "new slide inherits placeholders from its layout" — exactly our G2. lxml gives a raw-OOXML escape hatch for what python-pptx doesn't expose (theme color resolution, custom XML parts, footer fields). |
| PptxGenJS (TypeScript) | ❌ | Generation-only. Importing existing presentations/templates is an explicitly unimplemented non-roadmap feature — fatal for G1. |
| Aspose.Slides | ⚠️ optional | Commercial; high fidelity. Design for pluggability, do not depend on it. |
| PowerPoint COM | ❌ | Excluded by contract even where PowerPoint is installed; Windows-only; flaky in automation. Installed PowerPoint is reserved for manual visual inspection of outputs (see §15). |

The MCP best-practice default of TypeScript is overridden here because the
domain library decides the project: rewriting template-grade OOXML parsing in
TS is months of work; python-pptx is battle-tested.

### 2.2 MCP framework & transport

- **SDK:** official `mcp` Python SDK (FastMCP server API), Pydantic v2 models
  for every tool input/output.
- **Transport:** **stdio** (local, single-user, subprocess of the client). All
  logging to stderr. Streamable HTTP is out of scope for v1.
- **Server name:** `ppt_mcp`. **Tool prefix:** `ppt_` (snake_case,
  action-oriented) so tools don't collide with other servers.

### 2.3 Rendering: headless LibreOffice (optional dependency)

`soffice --headless --convert-to pdf|png` renders decks for visual validation,
thumbnails, and overflow ground truth. Detection at startup; if absent, all
render-dependent tools return a clear "rendering unavailable — install
LibreOffice" error and the rest of the server works normally. Geometric checks
(§7.3) use Pillow font metrics as a render-free approximation.

### 2.4 Supporting libraries

- `Pillow` — text-extent estimation, thumbnail post-processing, visual diff.
- `rapidfuzz` — fuzzy layout-name matching for recommendations.
- `pydantic` — schemas; `platformdirs` — data dir resolution.

---

## 3. Architecture overview

```text
┌────────────────────────────────────────────────────────────┐
│                    MCP tool surface (stdio)                 │
│ deck · authoring · template · compliance · style · render · │
│                      patterns · icons                       │
├────────────────────────────────────────────────────────────┤
│ Session Manager          │ Template Intelligence Layer      │
│  deck handles, snapshots │  registry, parser, layout        │
│  dry-run planner, undo   │  recommender, placeholder mapper │
├──────────────────────────┼──────────────────────────────────┤
│ Brand Style Registry     │ Compliance Engine                │
│  profiles, prompt        │  rule checks, findings,          │
│  composer, provenance    │  repair planner                  │
├──────────────────────────┴──────────────────────────────────┤
│ Built-in libraries: starter templates · slide patterns      │
│                     (forms) · icon sets (vendored, offline) │
├─────────────────────────────────────────────────────────────┤
│            Core OOXML engine (python-pptx + lxml)           │
├─────────────────────────────────────────────────────────────┤
│  Render service (LibreOffice headless, optional) + cache    │
├─────────────────────────────────────────────────────────────┤
│        Local store  ~/.ppt-mcp/  (json + files, no DB)      │
└─────────────────────────────────────────────────────────────┘
```

**Design rule:** every mutating tool goes through the dry-run planner — it
first computes a structured *change plan*, and only applies it when
`dry_run=false` (see §11.1).

---

## 4. Core concepts & data model

### 4.1 Deck sessions

Decks are addressed by opaque `deck_id` handles, not raw paths, so the agent
cannot accidentally clobber the source file:

- `ppt_open_deck(path)` copies the file into a session working dir; edits hit
  the working copy; `ppt_save_deck` writes back (or to a new path).
- Every applied change plan first snapshots the working file →
  `ppt_undo(deck_id, steps)` is cheap and reliable.
- Sessions persist across server restarts (manifest on disk) and are listed by
  `ppt_list_decks`.

### 4.2 Shape & placeholder addressing

Stable addressing is what makes multi-turn editing safe:

- **`shape_id`** — the OOXML `<p:cNvPr id>` (stable across edits within a file).
- **`placeholder_idx`** — the layout-inherited placeholder index.
- **`role`** — semantic role (`title`, `subtitle`, `body`, `picture`, `chart`,
  `table`, `footer`, `date`, `slide_number`, …) derived from placeholder type +
  layout analysis.

Tools accept a `shape_ref` union: role string (resolved against the slide's
layout), placeholder idx, or shape id. Responses always echo all three.

### 4.3 `content_spec` — the semantic content model

The agent describes *what the slide says*; the server decides *where it goes*:

```jsonc
{
  "intent": "risk_overview",          // free-text slide intent
  "title": "Key AI Governance Risks",
  "subtitle": null,
  "body": [                            // outline with nesting levels
    { "text": "Unclear data ownership", "level": 0 },
    { "text": "No model decision authority", "level": 0 }
  ],
  "table":  { "headers": [...], "rows": [...] },   // optional
  "chart":  { "type": "bar", "categories": [...], "series": [...] },
  "images": [ { "image_intent": "vision scene", "alt_text": "..." } ],
  "notes":  "Speaker notes…",
  "footer_overrides": null
}
```

Sections present in the spec but unmappable to the chosen layout are returned
as `unplaced_content` warnings — never silently dropped, never silently
free-floated onto the slide.

### 4.4 Template registry entry

```jsonc
{
  "template_id": "tpl_a1b2c3",
  "name": "ACME Corporate 2026", "version": "2.1", 
  "source_path": "...", "sha256": "...",
  "aspect_ratios": ["16:9"],
  "masters": [ { "master_id": "...", "name": "...", "layout_ids": [...] } ],
  "layouts": [
    {
      "layout_id": "lyt_exec_summary",
      "name": "Executive Summary",
      "master_id": "...",
      "placeholders": [
        { "idx": 0, "type": "title", "role": "title",
          "position_emu": {...}, "size_emu": {...},
          "default_font": "ACME Sans", "max_estimated_chars": 90 }
      ],
      "intent_tags": ["executive_summary", "management_summary"],  // inferred + user-editable
      "capacity": { "body_bullets": 5, "max_level": 2 }
    }
  ],
  "theme": {
    "color_scheme": { "dk1": "...", "lt1": "...", "accent1": "#4567FC", ... },
    "font_scheme":  { "major": "...", "minor": "..." },
    "effects": "ref"
  },
  "brand_rules": { "footer_required": true, "logo_shapes": [...], "notes": "..." },
  "registered_at": "...", "metadata": { ... }
}
```

Parsed once at registration, cached as `parsed.json`, invalidated by file hash.

### 4.5 Brand style profile

```jsonc
{
  "profile_id": "sp_acme_corporate", "name": "ACME Corporate Photography",
  "version": "1.2.0",
  "system_prompt": "Photorealistic, modern enterprise, calm confidence, …",
  "allowed_colors": ["#1A6BCC", "#E0A030", "#7FB54A"],
  "forbidden_motifs": ["robot hands", "glowing brains", "neon cyberpunk",
                        "generic handshake stock photo"],
  "text_in_image": "forbidden", "logo_usage": "forbidden",
  "media_type": "photography", "realism": "photorealistic",
  "composition": "circular collaboration motifs, natural light",
  "diversity_rules": "...", "default_aspect_ratio": "16:9",
  "negative_prompt_base": "no readable text, no fake logos, …",
  "example_prompts": [ ... ],
  "created_at": "...", "updated_at": "...",
  "history": [ { "version": "1.1.0", "changed": "...", "at": "..." } ]
}
```

Profiles are **configuration, not instruction**: their text is never executed,
never overrides server behavior, and cannot enable network access. Every
emitted image prompt records `profile_id` + `version` for provenance.

### 4.6 Image placeholder manifest

Long prompts must not pollute the visible slide. Each image placeholder is:

- **On the slide:** the layout's picture placeholder (preferred) or a labeled
  rectangle — short human-readable label only
  (`"Image placeholder: AI strategy collaboration scene"`).
- **In the file:** full record in a **custom XML part**
  (`/customXml/pptmcp-manifest.xml`, namespace `urn:ppt-mcp:manifest:v1`) so it
  survives round-trips through PowerPoint itself; a one-line pointer in
  speaker notes for human discoverability.

Record fields: slide index, shape_id, role, image_intent, prompt,
negative_prompt, aspect_ratio, placement + target size, alt_text, caption
suggestion, generation status (`pending|generated|approved`),
style_profile_id + version, created_by_tool, created_at, template layout name.

This is the evidence trail for "why does the deck suddenly look like a
Scandinavian bank hired a meditation coach?"

---

## 5. Template Intelligence Layer

### 5.1 Three template modes

| Mode | Input | Behavior | Status |
|---|---|---|---|
| **1. Native template** (default) | Clean `.potx`/template `.pptx` | Use existing masters/layouts/placeholders. Never invent design unless asked. | M2 |
| **2. Template extraction** | A finished corporate deck ("Final_v13_really_final.pptx") | Cluster slides by layout usage + recurring geometry, infer a layout registry, register it as a derived template (flagged `derived: true`). | M5 |
| **3. Theme-only** | Colors/fonts only (or `.thmx`) | Generate clean slides from theme; responses carry a standing warning that this is weaker than official masters. | M4 |

### 5.2 Template parsing

At `ppt_register_template`:

1. Unzip; walk `ppt/slideMasters/*`, `ppt/slideLayouts/*`, `ppt/theme/*` with
   python-pptx, dropping to lxml for theme color resolution, footer/field
   detection, and non-placeholder recurring shapes (logo candidates =
   identical picture/shape elements appearing on master or ≥80% of layouts).
2. Build the registry entry (§4.4) including a **placeholder schema** per
   layout and **capacity estimates** (how many bullets/chars fit, from
   placeholder geometry + default font metrics).
3. Infer `intent_tags` per layout from layout names (multilingual synonym
   table: "Agenda", "Inhalt", "Two Content", "Section Header", …) and
   placeholder composition (a layout with one 2×2 grid of content boxes tags
   `matrix`). Tags are stored and user-editable via `ppt_update_template`.
4. Render layout thumbnails (if LibreOffice present) for agent inspection.

### 5.3 Layout recommendation (`ppt_recommend_layout`)

Deterministic scoring, returned ranked with reasons — the calling LLM makes
the final call:

```text
score(layout, intent, content_spec) =
    w1 · name/intent match        (rapidfuzz vs layout name + intent_tags)
  + w2 · structural fit           (spec sections ↔ placeholder roles available)
  + w3 · capacity fit             (bullet count/char count vs layout capacity)
  − w4 · waste penalty            (required placeholders the spec leaves empty)
```

```jsonc
{ "recommendations": [
  { "layout_id": "lyt_risk_matrix", "name": "Risk Matrix", "confidence": 0.91,
    "reason": "intent tag 'risk' matched; title + 2x2 content placeholders fit 4 items" },
  { "layout_id": "lyt_three_col", "name": "Three Pillars", "confidence": 0.78,
    "reason": "fits 3 items but lacks severity/impact structure" }
] }
```

### 5.4 Placeholder mapping (`ppt_map_content_to_placeholders`)

Pure function: `(template_id, layout_id, content_spec) → mapping plan`.
Resolution order: explicit role match → placeholder type match → positional
convention (first body placeholder gets `body`). Output lists each placement,
predicted overflow risk, and `unplaced_content`. `ppt_add_slide` runs the same
mapper internally, so the agent can preview the mapping first or trust it.

---

## 6. Where the intelligence lives

### 6.1 Deterministic server, semantic client

The server never calls an LLM. Anything requiring judgment is split:

- **Server provides:** complete structured facts (layout schemas, theme,
  capacity numbers, compliance findings) + deterministic heuristics
  (recommendation scores, prompt assembly).
- **Client agent provides:** the actual semantic decisions, using those facts.

This keeps the server testable, fast, offline, and free of API keys. If
in-server generation is ever wanted, MCP **sampling** (server→client LLM
callback) is the upgrade path — noted as future work, not designed in.

### 6.2 Image prompt composition (`ppt_compose_image_prompt`)

Deterministic template assembly, not generation. Inputs: slide context (title,
intent, surrounding content — read from the deck), image intent, style
profile, constraints, aspect ratio/placement from the target placeholder.
Output:

```jsonc
{
  "prompt": "<style system prompt> + <slide-specific scene clause> + <composition/color clauses>",
  "negative_prompt": "<profile negative base + constraint additions>",
  "alt_text": "...", "aspect_ratio": "16:9",
  "provenance": { "style_profile_id": "...", "style_profile_version": "1.2.0",
                   "created_by_tool": "ppt_compose_image_prompt", "created_at": "..." }
}
```

The calling agent is free to refine the scene clause before writing it into
the placeholder — the server validates the final prompt against profile rules
(forbidden motifs as substring/regex checks, color allowlist) on write.

---

## 7. Compliance Engine

### 7.1 Rule set (v1)

| ID | Check | Severity |
|---|---|---|
| C01 | Slide's layout/master belongs to the reference template | error |
| C02 | Fonts resolve to theme font scheme (or documented exceptions) | error |
| C03 | Explicit colors are theme colors or within allowed delta | warning |
| C04 | Required footer/date/slide-number fields present where master defines them | error |
| C05 | Logo / recurring master elements not covered, deleted, or displaced | error |
| C06 | Placeholder bypass: free textbox/picture overlapping an empty placeholder of matching type | warning |
| C07 | Text overflow: content exceeds placeholder bounds (metric estimate; render-confirmed when available) | error |
| C08 | Density: content exceeds layout capacity estimate | warning |
| C09 | Manually positioned objects off the layout grid (> tolerance from any placeholder edge/guide) | info |
| C10 | Slide size / aspect ratio matches template | error |

Findings are structured: `{rule, severity, slide_index, shape_ref, message,
auto_fixable, suggested_fix}`.

### 7.2 Repair (`ppt_repair_compliance`)

Takes findings (or re-runs validation), produces a change plan per finding;
`strategy` ∈ `conservative` (only zero-risk fixes: re-link fonts to theme,
restore footers, snap colors to nearest theme color) | `aggressive` (also
migrate bypassing textboxes into placeholders, reflow overflow by shrink/
split). Always dry-run-first; never auto-applies `aggressive`.

### 7.3 Overflow detection without rendering

Estimate text extent with Pillow font metrics (font, size, wrap width from
placeholder geometry, autofit settings parsed from OOXML). Mark findings
`confidence: estimated`; when LibreOffice is available, `ppt_validate_…`
optionally renders affected slides and upgrades findings to
`confidence: rendered`.

### 7.4 Template application (`ppt_apply_template`)

Re-targeting an existing deck to a template is the riskiest operation in the
product. Therefore:

- `dry_run` **defaults to `true`**; the report enumerates per-slide fidelity
  risks: layout re-mapping choices, position/font/color changes, orphaned
  shapes, charts/tables likely to drift, footers gained/lost.
- Apply = clone deck into the template's master/layout set, re-attach each
  slide to the best-matching layout (same scorer as §5.3), move placeholder
  content across, keep orphans grouped + flagged.
- Post-apply, `ppt_validate_compliance` runs automatically and is included in
  the response.

---

## 8. Rendering & validation service

- `soffice` invocation wrapped with timeout + temp profile dir (parallel-safe).
- PDF for whole-deck export; per-slide PNG (via PDF→raster) for thumbnails,
  visual diff, and overflow confirmation.
- Render cache keyed by `(working file hash, slide, dpi)` under
  `~/.ppt-mcp/renders/`.
- Renders are returned as MCP image content (thumbnail-sized by default) plus
  the file path for full resolution.
- `ppt_visual_diff(deck_id, snapshot_ref)` — pixel diff (Pillow) between the
  current state and a named snapshot; returns changed-region boxes per slide.
  Used to verify "the logo survived my edit."

---

## 9. Built-in libraries: starter templates, slide patterns, icons, skill

Everything in this section is **vendored at build time and fully offline at
runtime** — a build script in the repo fetches/produces the assets once;
released artifacts contain them; the running server never downloads anything.
Built-in assets live read-only inside the installed package; user-added
templates, patterns, and icon sets go under `~/.ppt-mcp/`.

### 9.1 Starter template library (typical business design choices)

Many users have no clean corporate template — they should still get governed,
layout-driven decks instead of free-floating textboxes. The repo ships a small
library of starter `.potx` templates **built in-repo from scratch** (no
third-party deck content, so licensing is clean):

- **3–4 visual directions:** clean corporate light, consulting minimal,
  dark executive, neutral document-style.
- **Each with the full standard layout set:** title, agenda, section divider,
  one/two/three-column content, comparison, quote/statement, timeline/roadmap,
  KPI/dashboard, 2×2 matrix (risk/effort-impact), team/org, architecture/
  diagram, decision request, table, appendix, closing.
- Pre-registered in the template registry with `builtin: true`, hand-curated
  `intent_tags`, capacity data, and thumbnails.

Builtins flow through exactly the same Template Intelligence path as corporate
templates — no special-casing. When a corporate template is registered it wins
by default; builtins remain available as explicit choices and as the fallback
for theme-only mode (the theme is injected into a builtin's master/layout
skeleton, which strengthens mode 3 considerably).

### 9.2 Slide pattern library (business forms)

A **pattern** is a template-independent semantic recipe for a recurring
"form-like" business slide. Shipped patterns (v1):

```text
project_charter        one-pager: goal, scope, team, milestones, budget, risks
status_report          traffic lights, progress, accomplishments, risks, next steps
decision_request       situation, options (≥2) with pros/cons, recommendation, ask
raid_log               risks / assumptions / issues / dependencies table
risk_register          probability × impact table with owners and mitigations
okr_sheet              objective + key results with progress indicators
swot                   2×2 strengths/weaknesses/opportunities/threats
stakeholder_matrix     influence × interest grid
action_items           minutes / action list: owner, due date, status
roadmap_lite           phased timeline with milestones and swimlanes
```

Each pattern defines:

1. **A content schema** — a typed extension of `content_spec` with named
   fields (`status.overall`, `decision.options[].pros`, …) and validation
   rules ("decision_request requires ≥ 2 options").
2. **A rendering recipe** — how fields map onto the *target template's*
   layout placeholders and structured tables, with all styling derived from
   the template's theme (e.g. traffic-light red/amber/green resolved to theme
   or standard semantic colors, never hard-coded brand values).
3. **Capacity rules** — when content exceeds the form (12 actions on an
   action list), the dry-run plan proposes splitting or compressing.

Patterns render against any registered template — the corporate template
first, a builtin as fallback when no corporate layout fits. They are plain
data files (`patterns/*.json`); users drop their own into
`~/.ppt-mcp/patterns/`. This is deliberately *not* a fourth template mode:
patterns are content recipes that ride on whatever template is active.

### 9.3 Icon library (offline, vendored)

Corporate decks need icons constantly (process steps, KPI markers, section
markers). Bundled, license-clean, offline:

- **Default set: Google Material Symbols** (Apache 2.0 — safe to vendor).
  Architecture supports additional sets (e.g. Lucide/ISC, Tabler/MIT) dropped
  into the assets dir; each set ships SVGs + `index.json` (name, synonyms,
  tags, category) + its LICENSE/attribution file.
- **Search:** `ppt_search_icons` does fuzzy name/tag lookup (rapidfuzz over
  the index) and returns small preview renders.
- **Insertion pipeline:** recolor the SVG (Material/Lucide icons are
  single-fill paths, so recoloring is a fill swap) → rasterize at high DPI via
  a vendored `resvg` static binary (no cairo dependency chain) → insert as
  PNG; optionally also embed the native SVG part (`svgBlip`, supported by
  PowerPoint 2016+) with the PNG as fallback.
- **Governance:** default icon color comes from the theme (text or accent
  slots); compliance rule C03 applies to icon colors like any other explicit
  color; icon insertions are recorded in provenance like all mutations.

### 9.4 Companion agent skill

An MCP server defines *capabilities*; agents still need the *procedure*. The
repo ships a companion skill (Anthropic `SKILL.md` format, under
`skills/ppt-deck-authoring/`) that teaches any skill-capable client the
intended workflow:

- the template-first decision tree (corporate → builtin → theme-only),
- `content_spec` authoring and `recommend_layout`/mapping preview usage,
- the validate → repair → render verification loop,
- the image-placeholder / brand-style-prompt workflow,
- when to reach for patterns vs. plain layouts vs. freeform (almost never).

Keeping the skill in-repo means tool changes and workflow guidance evolve in
lockstep, and every MCP client gets the same governed behavior instead of
reinventing it. A skeleton of the skill is drafted at M2 and finalized at M6
against real usage.

---

## 10. MCP tool surface

~42 tools in 8 groups. All names prefixed `ppt_`. Consolidations vs. the
initial brainstorm: `list_slide_masters` + `list_slide_layouts` folded into
`ppt_inspect_template` (one call, less context burn); `compare_template_usage`
folded into `ppt_validate_compliance` (same computation, one entry point);
style-profile create/update unified as upsert with automatic versioning.

### 10.1 Deck lifecycle

| Tool | Signature | Notes |
|---|---|---|
| `ppt_create_deck` | `(template_id?, theme_only_spec?) → deck_id` | From registered template (mode 1), theme-only (mode 3), or blank. |
| `ppt_open_deck` | `(path) → deck_id + overview` | Copies to session working dir. |
| `ppt_save_deck` | `(deck_id, path?)` | Default: back to source path; refuses to overwrite a registered template file. |
| `ppt_close_deck` | `(deck_id, discard?)` | |
| `ppt_list_decks` | `() → sessions` | |
| `ppt_undo` | `(deck_id, steps=1)` | Snapshot-based. |

### 10.2 Reading & inspection

| Tool | Signature |
|---|---|
| `ppt_get_deck_overview` | `(deck_id) → slide count, sections, template fingerprint match, masters in use` |
| `ppt_get_slide` | `(deck_id, slide_index, response_format?) → shapes, placeholders (role/idx/id), text, notes, layout ref` |
| `ppt_search_deck` | `(deck_id, query, limit?, offset?) → text hits with slide/shape refs` |

### 10.3 Authoring

| Tool | Signature |
|---|---|
| `ppt_add_slide` | `(deck_id, layout_ref, content_spec, position?, dry_run?)` — the workhorse; runs mapper §5.4 |
| `ppt_set_placeholder_content` | `(deck_id, slide_index, shape_ref, content, dry_run?)` |
| `ppt_edit_text` | `(deck_id, slide_index, shape_ref, ops[], dry_run?)` — replace/insert/delete runs, bullet ops |
| `ppt_set_notes` | `(deck_id, slide_index, notes)` |
| `ppt_add_table` / `ppt_add_chart` / `ppt_add_image` | placeholder-first; free placement requires `allow_freeform=true` and emits a C06-style warning |
| `ppt_delete_slide` / `ppt_move_slide` / `ppt_duplicate_slide` | `(deck_id, slide_index, …)` |

### 10.4 Template registry & intelligence

| Tool | Signature |
|---|---|
| `ppt_register_template` | `(path, name?, version?, metadata?) → template_id` |
| `ppt_list_templates` | `(limit?, offset?)` |
| `ppt_inspect_template` | `(template_id) → masters, layouts (names + intent_tags + placeholder summary), theme, brand rules` |
| `ppt_inspect_layout` | `(template_id, layout_id) → full placeholder schema, capacity, thumbnail` |
| `ppt_update_template` | `(template_id, patch)` — edit intent_tags, brand rules, name/version |
| `ppt_extract_theme` | `(template_id \| deck_id) → color scheme, font scheme` |
| `ppt_extract_template_from_deck` | `(deck_id, name?) → derived template_id` (mode 2, M5) |
| `ppt_recommend_layout` | `(template_id, slide_intent, content_spec) → ranked layouts + reasons` |
| `ppt_map_content_to_placeholders` | `(template_id, layout_id, content_spec) → mapping plan, unplaced_content` |

### 10.5 Compliance

| Tool | Signature |
|---|---|
| `ppt_validate_compliance` | `(deck_id, template_id?, slides?, rules?) → findings[]` (template defaults to the deck's registered fingerprint match) |
| `ppt_repair_compliance` | `(deck_id, findings?, strategy?, dry_run=true)` |
| `ppt_apply_template` | `(deck_id, template_id, strategy?, dry_run=true) → fidelity risk report \| applied + validation` |

### 10.6 Brand style profiles & image placeholders

| Tool | Signature |
|---|---|
| `ppt_set_style_profile` | `(name, system_prompt, metadata?) → profile_id, version` — upsert; bumps version on change |
| `ppt_get_style_profile` / `ppt_list_style_profiles` / `ppt_delete_style_profile` | |
| `ppt_compose_image_prompt` | `(deck_id, slide_index, image_intent, profile_id, constraints?) → prompt bundle (§6.2)` |
| `ppt_create_image_placeholder` | `(deck_id, slide_index, spec, dry_run?)` — uses layout picture placeholder when available |
| `ppt_list_image_placeholders` | `(deck_id, status?) → manifest records` |
| `ppt_update_image_placeholder` | `(deck_id, slide_index, shape_ref, patch, dry_run?)` — validated against profile rules |
| `ppt_fill_image_placeholder` | `(deck_id, slide_index, shape_ref, image_path)` — inserts the externally generated image, sets status `generated`, preserves crop/aspect rules |

### 10.7 Rendering & export

| Tool | Signature |
|---|---|
| `ppt_render_slide` | `(deck_id, slide_index, dpi?) → image content + path` |
| `ppt_render_deck` | `(deck_id, dpi?) → thumbnails` |
| `ppt_export_pdf` | `(deck_id, path)` |
| `ppt_visual_diff` | `(deck_id, snapshot_ref) → changed regions per slide` |

### 10.8 Patterns & icons

| Tool | Signature |
|---|---|
| `ppt_list_patterns` | `(limit?, offset?) → pattern ids, names, one-line purpose` |
| `ppt_get_pattern` | `(pattern_id) → content schema, field docs, example` |
| `ppt_add_slide_from_pattern` | `(deck_id, pattern_id, content, template_id?, position?, dry_run?)` — validates content against the pattern schema, renders via the active template (§9.2) |
| `ppt_list_icon_sets` | `() → sets with license info` |
| `ppt_search_icons` | `(query, set_id?, limit?) → ranked icons (id, name, tags) + preview images` |
| `ppt_insert_icon` | `(deck_id, slide_index, icon_id, target, color?, size?, dry_run?)` — `target` = shape_ref or position; color defaults to theme |

---

## 11. Cross-cutting tool conventions

### 11.1 Universal `dry_run` change plans

Every mutating tool accepts `dry_run` (default `false` for routine authoring,
**`true` for high-risk tools**: `repair_compliance`, `apply_template`). A dry
run returns the structured change plan the real run would apply:

```jsonc
{ "plan": [ { "op": "fill_placeholder", "slide": 4, "shape_ref": {...},
              "summary": "set title to 'Key AI Governance Risks'" } ],
  "warnings": [ "body has 7 bullets; layout capacity ≈ 5 (C08)" ],
  "unplaced_content": [] }
```

The applied response = same plan + `applied: true` + post-state summary.

### 11.2 Responses, pagination, errors

- `response_format`: `"markdown"` (default, human-skimmable) | `"json"`
  (complete structured data). Inspection tools support both.
- List tools: `limit` (default 25) / `offset`, returning `total_count`,
  `has_more`, `next_offset`.
- Errors are actionable and in-result (not protocol errors):
  `"Layout 'lyt_exec' not found in template tpl_a1b2c3. Call
  ppt_inspect_template to list valid layout_ids."`
- Annotations on every tool: `readOnlyHint` for §10.2/10.4-inspection/10.7
  and pattern/icon search,
  `destructiveHint: true` only for `delete_slide`, `close_deck(discard)`,
  `delete_style_profile`, `apply_template`; `openWorldHint: false` everywhere
  (local-only).

### 11.3 Units

API uses points for font sizes and **inches** (float) for geometry by default,
with optional `{value, unit}` objects (`in|cm|pt|emu`). Internally everything
is EMU. Responses include both EMU and inches.

---

## 12. Storage layout

```text
~/.ppt-mcp/                      (platformdirs-resolved; overridable via env)
├─ config.json                   (render engine path, defaults — no network option exists)
├─ registry/templates.json       (index; entries in §4.4 form)
├─ templates/<template_id>/
│  ├─ source.potx                (immutable copy)
│  ├─ parsed.json
│  └─ thumbnails/*.png
├─ style_profiles/<profile_id>.json
├─ patterns/*.json                (user-added; builtins ship in the package)
├─ assets/icons/<set_id>/         (user-added sets: svg/, index.json, LICENSE)
├─ sessions/<deck_id>/
│  ├─ manifest.json              (source path, opened_at, template link)
│  ├─ working.pptx
│  └─ snapshots/<n>.pptx
├─ renders/<hash>/…
└─ logs/provenance.jsonl         (append-only: tool, args digest, profile/template versions)
```

Boring, inspectable, no database, fully local.

---

## 13. Security & governance

- **Offline enforced:** no outbound network calls anywhere in the codebase —
  no update checks, no telemetry, no font/CDN fetches. CI greps the dependency
  tree and source for network APIs (`requests`, `httpx`, `urllib`, sockets) to
  keep the guarantee honest.
- **Path safety:** all user paths normalized + checked against directory
  traversal; session working copies prevent accidental source corruption;
  registered template sources are immutable.
- **Prompt-as-config:** style profile text is data. It is interpolated into
  output strings only — never into shell commands, file paths, or server
  control flow. Validation rules (forbidden motifs etc.) run *on* it, are
  never run *by* it.
- **Provenance:** every generated artifact (slide from layout, image prompt,
  repair) logs tool, timestamp, template/profile id + version to
  `provenance.jsonl` and, where applicable, the in-file manifest (§4.6).
- **Subprocess hygiene:** `soffice` called with fixed argv (no shell), temp
  profile, timeout, and output confined to the renders dir.

---

## 14. Milestones

| | Scope | Exit criterion |
|---|---|---|
| **M0** | Project scaffold, FastMCP stdio server, session manager, storage layout, `open/create/save/close/list/get_slide/search` | Open a real corporate deck, read every slide's placeholders + text via MCP Inspector |
| **M1** | Authoring: `add_slide` (against the deck's own layouts), placeholder filling, text ops, notes, table/image/chart, delete/move/duplicate, undo, dry-run planner | Build a 10-slide deck from a `.potx`'s own layouts with zero absolute positioning |
| **M2** | Template Intelligence: registry, parsing (§5.2), `inspect_*`, `recommend_layout`, `map_content_to_placeholders`, `create_deck_from_template`; first builtin starter template registered (§9.1); companion skill skeleton (§9.4) | "Create an executive summary slide" → correct layout chosen and filled on 3 distinct real-world templates |
| **M3** | Compliance engine C01–C10, `repair_compliance`, LibreOffice rendering, overflow confirmation, `render_*`, `export_pdf`, `visual_diff` | Detects seeded violations (wrong font, deleted footer, overflow, bypassed placeholder) with zero false negatives on the test corpus |
| **M4** | Brand style registry, image prompt composition, image placeholder manifest (custom XML part), `fill_image_placeholder`, theme-only mode | Deck with governed image placeholders survives a PowerPoint round-trip with manifest intact |
| **M5** | `apply_template` with fidelity report, template extraction from messy decks, eval suite (§15), docs | 10-question eval ≥ 8/10 with a stock MCP client |
| **M6** | Built-in libraries complete: full starter template set, slide pattern library + `add_slide_from_pattern`, icon library + search/insert pipeline, companion skill finalized | `status_report` pattern renders correctly on a corporate template and two builtins; icon search + themed insert works fully offline; a skill-guided agent builds a governed deck end-to-end |

---

## 15. Testing & evaluation

- **Unit:** parsers and mappers against a fixture corpus of templates —
  Office-default templates + several real-world corporate-style `.potx` built
  for the repo (varied layout naming languages, 4:3 and 16:9).
- **Property tests:** open→edit→save→reopen round-trips never lose masters,
  layouts, theme parts, or the custom XML manifest (compare part inventories
  + key XML invariants).
- **Compliance seeding:** scripts inject known violations into clean decks;
  the validator must find exactly them.
- **Render tests:** gated on LibreOffice presence in CI; golden-image
  thumbnails with perceptual-diff tolerance.
- **PowerPoint ground truth (manual, dev machines):** where PowerPoint is
  installed, each milestone's output decks are opened in real PowerPoint to
  verify they load without repair prompts and render as intended. This is a
  human/visual checkpoint — never automated via COM — and is the fidelity
  arbiter when LibreOffice and PowerPoint disagree.
- **Agent evals (per mcp-builder methodology):** 10 read-only,
  independently verifiable questions over a fixture template + deck (e.g.
  "Which layout in template X has a picture placeholder wider than 6 inches,
  and what is its name?"), run with the MCP evaluation harness.

---

## 16. Risks & open questions

| Risk | Mitigation |
|---|---|
| python-pptx gaps (some chart types, exotic placeholder inheritance) | lxml escape hatch is a first-class internal API, not a hack; contribute upstream where sane |
| LibreOffice render fidelity ≠ PowerPoint fidelity | Treat renders as *validation evidence*, not pixel truth; document the caveat in tool descriptions; locally installed PowerPoint is the manual ground-truth arbiter (§15) |
| Layout `intent_tags` inference quality on creatively-named corporate layouts | Tags are user-editable (`ppt_update_template`); recommendation always returns reasons so the agent can overrule |
| Tool count (~36) burning agent context | Concise descriptions, markdown-default responses, consolidated inspect tools; revisit grouping after M2 usage data |
| Template extraction (mode 2) is genuinely hard | Scoped to M5; ship value without it; start with "cluster by explicit layout usage" before inferring from geometry |
| Pattern visual quality varies across arbitrary corporate templates | Recipes use only placeholder geometry + theme values; dry-run reports fit score and falls back to a builtin template when no corporate layout fits |
| Icon rendering fidelity (`resvg` raster vs native SVG) | Dual embed (svgBlip + PNG fallback); golden-image tests per icon set; sets are single-fill paths, the simplest SVG class |

Open questions (decide before the affected milestone):

1. **Chart fidelity strategy (M1):** python-pptx chart support is partial —
   define the supported chart matrix and the fallback (native table or
   flagged image placeholder) for the rest.
2. **Theme color tolerance (M3):** exact-match only, or ΔE distance with a
   configurable threshold for C03?
3. **Multi-deck merge** (`combine decks`, reuse slides across decks) — demand
   exists, but out of scope until after M5.
4. **MCP resources** (expose templates/thumbnails as MCP resources in addition
   to tools) — nice for clients that surface resources; evaluate at M2.

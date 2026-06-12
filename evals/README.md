# ppt-mcp evaluation suite

Ten read-only, independently verifiable questions over a deterministic
fixture deck, following the MCP server evaluation methodology (DESIGN.md §15).

## Files

- `questions.xml` — the 10 QA pairs (answers verified by string comparison)
- `fixtures/eval_deck.pptx` — generated fixture; do not edit by hand

## Regenerating / verifying

```powershell
uv run python scripts/make_eval_fixtures.py      # rebuild the fixture deck
uv run python scripts/compute_eval_answers.py    # solve via the MCP tools, print ground truth
```

`compute_eval_answers.py` answers every question through the actual server
tool functions — if it prints values matching `questions.xml`, the suite is
consistent.

## Running against an agent

Connect the server (`uv run ppt-mcp`, stdio) to an MCP client, ask each
question in a fresh session, and compare the agent's answer to the
`<answer>` element. The M5 exit criterion is ≥ 8/10 with a stock client.

Note: the fixture deck intentionally contains exactly one compliance
violation (Comic Sans MS on the slide-5 title) for the C02 questions.

"""End-to-end smoke check: launch the server over stdio, initialize, list
tools, open a generated deck, and read a slide — exactly as an MCP client
would. Run with: uv run python scripts/smoke_stdio.py"""

import asyncio
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, str(Path(__file__).parent.parent / "tests"))

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


async def main() -> None:
    from conftest import build_sample_deck

    deck_path = build_sample_deck(Path(tempfile.mkdtemp()) / "smoke.pptx")

    params = StdioServerParameters(command="uv", args=["run", "ppt-mcp"])
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            tools = await session.list_tools()
            names = sorted(t.name for t in tools.tools)
            print(f"tools ({len(names)}):", ", ".join(names))

            opened = await session.call_tool("ppt_open_deck", {"path": str(deck_path)})
            deck_id = opened.structuredContent["deck_id"]
            print("opened:", deck_id)

            slide = await session.call_tool(
                "ppt_get_slide", {"deck_id": deck_id, "slide_index": 2}
            )
            print(slide.content[0].text)

            await session.call_tool("ppt_close_deck", {"deck_id": deck_id})
            print("closed:", deck_id)


if __name__ == "__main__":
    asyncio.run(main())

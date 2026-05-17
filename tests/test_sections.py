from __future__ import annotations

import pytest

pytestmark = pytest.mark.anyio


async def test_section_listing_and_update(invoke_tool, copy_fixture):
    path = copy_fixture("with_sections.docx")
    listed = await invoke_tool("list_sections", path=str(path))
    updated = await invoke_tool(
        "set_section_page_setup",
        path=str(path),
        section_index=1,
        orientation="portrait",
        paper_size="a4",
        left_margin_points=36,
    )

    assert listed["status"] == "ok"
    assert len(listed["result"]["sections"]) == 2
    assert updated["status"] == "ok"
    assert updated["result"]["section"]["orientation"] == "portrait"

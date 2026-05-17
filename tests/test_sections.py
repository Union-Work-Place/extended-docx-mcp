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


async def test_fractional_section_margins_are_serialized_safely(invoke_tool, fixtures_dir):
    path = fixtures_dir / "fractional_sections.docx"

    listed = await invoke_tool("list_sections", path=str(path))
    read = await invoke_tool("read_docx", path=str(path), include_sections=True, include_tables=False, include_comments=False, include_revisions=False)

    assert listed["status"] == "ok"
    assert listed["result"]["sections"][1]["left_margin_points"] == pytest.approx(1984.251968503937 / 20)
    assert read["status"] == "ok"
    assert read["result"]["sections"][1]["left_margin_points"] == pytest.approx(1984.251968503937 / 20)

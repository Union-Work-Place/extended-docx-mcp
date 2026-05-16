from __future__ import annotations

import pytest

pytestmark = pytest.mark.anyio


async def test_style_listing_and_application(invoke_tool, copy_fixture):
    path = copy_fixture("with_styles.docx")
    listed = await invoke_tool("list_styles", path=str(path), include_builtin=False)
    created = await invoke_tool(
        "create_or_update_style",
        path=str(path),
        style_name="BodyNote",
        font_name="Calibri",
        font_size_points=11,
        space_after_points=8,
    )
    applied = await invoke_tool("apply_paragraph_style", path=str(path), style_name="BodyNote", paragraph_index=1)

    assert listed["status"] == "ok"
    assert any(style["name"] == "Callout" for style in listed["result"]["styles"])
    assert created["status"] == "ok"
    assert created["result"]["style"]["name"] == "BodyNote"
    assert applied["status"] == "ok"
    assert applied["result"]["style_name"] == "BodyNote"

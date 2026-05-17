from __future__ import annotations

import pytest

pytestmark = pytest.mark.anyio


async def test_table_reading_tools(invoke_tool, fixtures_dir):
    tables = await invoke_tool("list_tables", path=str(fixtures_dir / "with_tables.docx"))
    cell = await invoke_tool(
        "get_table_cell_content",
        path=str(fixtures_dir / "with_tables.docx"),
        table_index=0,
        row_index=1,
        cell_index=0,
    )

    assert tables["status"] == "ok"
    assert len(tables["result"]["tables"]) == 2
    assert cell["status"] == "ok"
    assert cell["result"]["text"] == "Budget"


async def test_table_editing_tools(invoke_tool, copy_fixture):
    path = copy_fixture("with_tables.docx")
    inserted = await invoke_tool(
        "insert_table",
        path=str(path),
        after_paragraph=0,
        data=[["A", "B"], ["1", "2"]],
        track_changes=False,
    )
    updated = await invoke_tool(
        "update_table_cell",
        path=str(path),
        table_index=0,
        row_index=1,
        cell_index=1,
        text="110",
        track_changes=False,
    )
    formatted = await invoke_tool("set_table_format", path=str(path), table_index=0, alignment="center", allow_auto_fit=True)

    assert inserted["status"] == "ok"
    assert inserted["result"]["rows"] == 2
    assert updated["status"] == "ok"
    assert updated["result"]["text"] == "110"
    assert formatted["status"] == "ok"
    assert formatted["result"]["table"]["alignment"] == "center"

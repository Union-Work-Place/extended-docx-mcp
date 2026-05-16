from __future__ import annotations

import pytest

pytestmark = pytest.mark.anyio


async def test_range_tools_cover_read_replace_delete_and_block_insert(invoke_tool, copy_fixture):
    path = copy_fixture("simple.docx")
    paragraph_range = await invoke_tool("get_paragraph_range", path=str(path), start_paragraph=1, end_paragraph=2)
    replaced = await invoke_tool(
        "replace_text_in_range",
        path=str(path),
        start_paragraph=1,
        end_paragraph=2,
        find_text="budget",
        replace_with="forecast",
    )
    inserted = await invoke_tool(
        "insert_block_after_paragraph",
        path=str(path),
        after_paragraph=1,
        block={"type": "heading", "level": 2, "text": "Inserted block"},
    )
    deleted = await invoke_tool("delete_paragraph_range", path=str(path), start_paragraph=2, end_paragraph=2)

    assert paragraph_range["status"] == "ok"
    assert len(paragraph_range["result"]["paragraphs"]) == 2
    assert replaced["status"] == "ok"
    assert replaced["result"]["replacements"] == 1
    assert inserted["status"] == "ok"
    assert inserted["result"]["block_type"] == "heading"
    assert deleted["status"] == "ok"
    assert deleted["result"]["deleted_count"] == 1


async def test_batch_tools_cover_text_format_and_table_updates(invoke_tool, copy_fixture):
    text_path = copy_fixture("simple.docx")
    replace_result = await invoke_tool(
        "batch_replace_text",
        path=str(text_path),
        replacements=[
            {"find_text": "Alpha", "replace_with": "Beta"},
            {"find_text": "draft", "replace_with": "approved"},
        ],
    )
    format_result = await invoke_tool(
        "batch_set_paragraph_format",
        path=str(text_path),
        paragraph_indices=[1, 2],
        alignment="center",
        space_after_points=12,
    )
    table_path = copy_fixture("with_tables.docx")
    table_result = await invoke_tool(
        "batch_update_table_cells",
        path=str(table_path),
        updates=[
            {"table_index": 0, "row_index": 1, "cell_index": 1, "text": "125"},
            {"table_index": 1, "row_index": 1, "cell_index": 1, "text": "Published"},
        ],
    )
    verify = await invoke_tool("get_table_cell_content", path=str(table_path), table_index=1, row_index=1, cell_index=1)

    assert replace_result["status"] == "ok"
    assert replace_result["result"]["operations"] == 2
    assert format_result["status"] == "ok"
    assert format_result["result"]["updated"] == 2
    assert table_result["status"] == "ok"
    assert table_result["result"]["updated"] == 2
    assert verify["result"]["text"] == "Published"

from __future__ import annotations

import pytest

pytestmark = pytest.mark.anyio


async def test_write_docx_and_plain_editing_flow(invoke_tool, tmp_path):
    target = tmp_path / "generated.docx"
    created = await invoke_tool(
        "write_docx",
        path=str(target),
        blocks=[
            {"type": "heading", "level": 1, "text": "Generated file"},
            {"type": "paragraph", "text": "budget draft"},
        ],
    )
    replaced = await invoke_tool("replace_text", path=str(target), find_text="draft", replace_with="approved", track_changes=False)
    inserted = await invoke_tool("insert_paragraph", path=str(target), after_paragraph=1, text="Inserted body paragraph", track_changes=False)
    deleted = await invoke_tool("delete_paragraph", path=str(target), paragraph_index=2, track_changes=False)

    assert created["status"] == "ok"
    assert created["result"]["blocks_written"] == 2
    assert replaced["status"] == "ok"
    assert replaced["result"]["replacements"] == 1
    assert inserted["status"] == "ok"
    assert inserted["result"]["inserted_after"] == 1
    assert deleted["status"] == "ok"
    assert deleted["result"]["deleted_paragraph"] == 2


async def test_tracked_replace_creates_revisions(invoke_tool, copy_fixture):
    path = copy_fixture("simple.docx")
    replace_result = await invoke_tool(
        "replace_text",
        path=str(path),
        find_text="draft",
        replace_with="approved",
        track_changes=True,
        author="QA",
    )
    revisions = await invoke_tool("list_revisions", path=str(path))

    assert replace_result["status"] == "ok"
    assert replace_result["result"]["replacements"] == 1
    assert revisions["status"] == "ok"
    assert revisions["result"]["revisions"]

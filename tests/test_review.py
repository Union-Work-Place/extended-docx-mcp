from __future__ import annotations

import pytest

pytestmark = pytest.mark.anyio


async def test_comments_and_revisions_are_listed(invoke_tool, fixtures_dir):
    comments = await invoke_tool("list_comments", path=str(fixtures_dir / "with_review.docx"))
    revisions = await invoke_tool("list_revisions", path=str(fixtures_dir / "with_review.docx"))
    details = await invoke_tool("get_revision_details", path=str(fixtures_dir / "with_review.docx"), revision_index=0, context_paragraphs=1)

    assert comments["status"] == "ok"
    assert comments["result"]["comments"]
    assert revisions["status"] == "ok"
    assert revisions["result"]["revisions"]
    assert details["status"] == "ok"
    assert details["result"]["revision"]["index"] == 0


async def test_comment_mutations_and_revision_resolution(invoke_tool, copy_fixture):
    path = copy_fixture("with_review.docx")
    added = await invoke_tool(
        "add_comment",
        path=str(path),
        paragraph_index=0,
        comment_text="Top-level note",
        author="Lead",
        initials="LD",
    )
    replied = await invoke_tool(
        "add_comment_reply",
        path=str(path),
        comment_index=0,
        reply_text="Acknowledged",
        author="Lead",
        initials="LD",
    )
    accepted = await invoke_tool("accept_all_revisions", path=str(path))

    rejected_path = copy_fixture("with_review.docx")
    rejected = await invoke_tool("reject_all_revisions", path=str(rejected_path))

    assert added["status"] == "ok"
    assert replied["status"] == "ok"
    assert accepted["status"] == "ok"
    assert accepted["result"]["remaining"] == 0
    assert rejected["status"] == "ok"
    assert rejected["result"]["remaining"] == 0

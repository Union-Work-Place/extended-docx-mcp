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


async def test_comment_tools_use_canonical_paragraph_indices(invoke_tool, copy_fixture):
    path = copy_fixture("table_before_target.docx")

    match_result = await invoke_tool("find_text_occurrences", path=str(path), target_text="UNIQUE_TARGET_AFTER_TABLE")
    comment_result = await invoke_tool(
        "add_comment_to_matching_text",
        path=str(path),
        target_text="UNIQUE_TARGET_AFTER_TABLE",
        comment_text="Canonical index comment",
        author="QA",
        initials="QA",
        paragraph_index=match_result["result"]["matches"][0]["paragraph_index"],
    )
    extract_result = await invoke_tool("extract_text", path=str(path), start_paragraph=4, count=2)
    comments_result = await invoke_tool("list_comments", path=str(path))

    assert match_result["status"] == "ok"
    assert match_result["result"]["matches"][0]["paragraph_index"] == 5
    assert comment_result["status"] == "ok"
    assert comment_result["result"]["paragraph_index"] == 5
    assert extract_result["status"] == "ok"
    assert extract_result["result"]["paragraphs"][1]["text"] == "UNIQUE_TARGET_AFTER_TABLE"
    assert comments_result["status"] == "ok"
    assert comments_result["result"]["comments"][0]["text"] == "Canonical index comment"


async def test_revision_details_use_canonical_context_indices(invoke_tool, copy_fixture):
    path = copy_fixture("table_before_target.docx")
    replace_result = await invoke_tool(
        "replace_text",
        path=str(path),
        find_text="UNIQUE_TARGET_AFTER_TABLE",
        replace_with="UPDATED_TARGET",
        track_changes=True,
        author="QA",
    )
    details = await invoke_tool("get_revision_details", path=str(path), revision_index=0, context_paragraphs=1)

    assert replace_result["status"] == "ok"
    assert details["status"] == "ok"
    assert details["result"]["revision"]["paragraph"]["index"] == 5
    assert [item["index"] for item in details["result"]["revision"]["context"]] == [4, 5]

from __future__ import annotations

import pytest

pytestmark = pytest.mark.anyio


async def test_read_docx_returns_document_window(invoke_tool, fixtures_dir):
    result = await invoke_tool("read_docx", path=str(fixtures_dir / "complex.docx"), paragraph_count=10)
    assert result["status"] == "ok"
    assert result["result"]["counts"]["tables"] == 1
    assert result["result"]["counts"]["sections"] == 2
    assert result["metadata"]["path"].endswith("complex.docx")


async def test_extract_and_find_tools_cover_text_queries(invoke_tool, fixtures_dir):
    extracted = await invoke_tool("extract_text", path=str(fixtures_dir / "simple.docx"), start_paragraph=0, count=2)
    matches = await invoke_tool("find_text_occurrences", path=str(fixtures_dir / "simple.docx"), target_text="budget")
    paragraphs = await invoke_tool("find_paragraphs", path=str(fixtures_dir / "with_styles.docx"), style_name="Callout")

    assert extracted["status"] == "ok"
    assert extracted["result"]["paragraphs"][1]["text"] == "Alpha project budget is 100."
    assert matches["status"] == "ok"
    assert matches["result"]["returned"] == 1
    assert paragraphs["status"] == "ok"
    assert paragraphs["result"]["returned"] == 1

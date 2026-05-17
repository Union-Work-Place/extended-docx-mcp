"""Range and block editing tools exposed through FastMCP."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from ops.document_ops import (
    find_paragraph,
    insert_structured_block_after,
    iter_paragraphs,
    normalize_paragraph_range,
    paragraph_to_dict,
)
from ops.package_io import load_document, save_document
from ops.text_ops import replace_in_paragraph_plain
from toolsets.response_schema import tool_response


def register_range_tools(server: FastMCP) -> None:
    """Register range and block editing helpers."""

    @server.tool()
    @tool_response("get_paragraph_range")
    def get_paragraph_range(
        path: str,
        start_paragraph: int,
        end_paragraph: int,
        include_runs: bool = True,
    ) -> dict[str, Any]:
        doc, resolved = load_document(path)
        paragraphs = iter_paragraphs(doc)
        start, end = normalize_paragraph_range(len(paragraphs), start_paragraph, end_paragraph)
        return {
            "path": str(resolved),
            "engine": "python-docx",
            "start_paragraph": start,
            "end_paragraph": end,
            "paragraphs": [paragraph_to_dict(paragraphs[index], index, include_runs) for index in range(start, end + 1)],
        }

    @server.tool()
    @tool_response("replace_text_in_range")
    def replace_text_in_range(
        path: str,
        start_paragraph: int,
        end_paragraph: int,
        find_text: str,
        replace_with: str,
        output_path: str | None = None,
        match_case: bool = False,
        find_whole_words_only: bool = False,
    ) -> dict[str, Any]:
        doc, source_path = load_document(path)
        paragraphs = iter_paragraphs(doc)
        start, end = normalize_paragraph_range(len(paragraphs), start_paragraph, end_paragraph)
        replacements = 0
        changed_paragraphs = 0
        for index in range(start, end + 1):
            changed = replace_in_paragraph_plain(
                paragraphs[index],
                find_text,
                replace_with,
                match_case=match_case,
                find_whole_words_only=find_whole_words_only,
            )
            if changed:
                replacements += changed
                changed_paragraphs += 1
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "start_paragraph": start,
            "end_paragraph": end,
            "replacements": replacements,
            "paragraphs_changed": changed_paragraphs,
        }

    @server.tool()
    @tool_response("delete_paragraph_range")
    def delete_paragraph_range(
        path: str,
        start_paragraph: int,
        end_paragraph: int,
        output_path: str | None = None,
    ) -> dict[str, Any]:
        doc, source_path = load_document(path)
        paragraphs = iter_paragraphs(doc)
        start, end = normalize_paragraph_range(len(paragraphs), start_paragraph, end_paragraph)
        deleted_text = [paragraphs[index].text for index in range(start, end + 1)]
        for index in range(end, start - 1, -1):
            paragraphs[index]._element.getparent().remove(paragraphs[index]._element)
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "start_paragraph": start,
            "end_paragraph": end,
            "deleted_count": len(deleted_text),
            "deleted_text": deleted_text,
        }

    @server.tool()
    @tool_response("insert_block_after_paragraph")
    def insert_block_after_paragraph(
        path: str,
        block: dict[str, Any],
        output_path: str | None = None,
        after_paragraph: int | None = None,
        anchor_text: str | None = None,
    ) -> dict[str, Any]:
        doc, source_path = load_document(path)
        target, actual_index = find_paragraph(doc, after_paragraph, anchor_text)
        inserted = insert_structured_block_after(target, block, 0)
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "after_paragraph": actual_index,
            "block_type": str(block.get("type", "paragraph")).lower(),
            "inserted_text": getattr(inserted, "text", None),
        }

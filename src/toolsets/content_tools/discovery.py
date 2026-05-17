"""Read-only content discovery tools exposed through FastMCP."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from ops.package_io import load_document, read_zip_xml, resolve_path
from ops.review import (
    iter_document_paragraphs_xml,
    list_comments_xml,
    list_revisions_xml,
    paragraph_revision_view,
)
from ops.structure_ops import iter_paragraphs, iter_tables, paragraph_to_dict, section_to_dict, table_to_dict
from ops.text_ops import matches_for_text
from toolsets.response_schema import tool_response


def register_content_discovery_tools(server: FastMCP) -> None:
    """Register read-oriented content tools on a FastMCP server.

    Args:
        server: MCP server to populate.
    """

    @server.tool()
    @tool_response("read_docx")
    def read_docx(
        path: str,
        start_paragraph: int = 0,
        paragraph_count: int = 100,
        include_runs: bool = True,
        include_tables: bool = True,
        include_comments: bool = True,
        include_revisions: bool = True,
        include_sections: bool = True,
        include_text: bool = True,
        max_text_chars: int = 12000,
    ) -> dict[str, Any]:
        """Read a DOCX file as a structured model.

        Args:
            path: Path to the DOCX file.
            start_paragraph: Starting paragraph index for the returned window.
            paragraph_count: Maximum number of paragraphs to return.
            include_runs: Include run-level metadata for returned paragraphs.
            include_tables: Include serialized table data.
            include_comments: Include serialized comments from OOXML.
            include_revisions: Include serialized tracked revisions from OOXML.
            include_sections: Include serialized section/page setup data.
            include_text: Include the full document text summary.
            max_text_chars: Maximum number of full-document text characters to return.

        Returns:
            Structured document payload for downstream MCP clients.
        """

        if start_paragraph < 0:
            raise ValueError("start_paragraph must be >= 0")
        if paragraph_count < 0:
            raise ValueError("paragraph_count must be >= 0")
        doc, resolved = load_document(path)
        paragraphs = iter_paragraphs(doc)
        document_root = read_zip_xml(resolved, "word/document.xml")
        paragraph_xml = iter_document_paragraphs_xml(document_root) if document_root is not None else []
        end = min(len(paragraphs), start_paragraph + paragraph_count)
        paragraph_items = []
        for index in range(start_paragraph, end):
            item = paragraph_to_dict(paragraphs[index], index, include_runs)
            if index < len(paragraph_xml):
                revision_view = paragraph_revision_view(paragraph_xml[index])
                item["text"] = revision_view["visible_text"]
                item["text_with_revisions"] = revision_view["annotated_text"]
                item["inserted_text"] = revision_view["inserted_text"]
                item["deleted_text"] = revision_view["deleted_text"]
            paragraph_items.append(item)
        full_text = ""
        if include_text:
            if paragraph_xml:
                full_text = "\n".join(paragraph_revision_view(paragraph)["visible_text"] for paragraph in paragraph_xml).strip()
            else:
                full_text = "\n".join(paragraph.text for paragraph in paragraphs).strip()
        tables = iter_tables(doc)
        comments = list_comments_xml(resolved) if include_comments else []
        revisions = list_revisions_xml(resolved) if include_revisions else []
        result: dict[str, Any] = {
            "path": str(resolved),
            "engine": "python-docx+lxml",
            "paragraph_window": {
                "start": start_paragraph,
                "count": paragraph_count,
                "returned": len(paragraph_items),
                "total": len(paragraphs),
            },
            "paragraphs": paragraph_items,
            "counts": {
                "paragraphs": len(paragraphs),
                "tables": len(tables),
                "sections": len(doc.sections),
                "comments": len(comments) if include_comments else None,
                "revisions": len(revisions) if include_revisions else None,
            },
        }
        if include_text:
            result["text"] = full_text[:max_text_chars]
            result["text_truncated"] = len(full_text) > max_text_chars
        if include_tables:
            result["tables"] = [table_to_dict(table, table_index, include_cells=True) for table_index, table in enumerate(tables)]
        if include_comments:
            result["comments"] = comments
        if include_revisions:
            result["revisions"] = revisions
        if include_sections:
            result["sections"] = [section_to_dict(section, index) for index, section in enumerate(doc.sections)]
        return result

    @server.tool()
    @tool_response("extract_text")
    def extract_text(path: str, start_paragraph: int = 0, count: int = 50) -> dict[str, Any]:
        """Extract a paragraph window with revision-aware text.

        Args:
            path: Path to the DOCX file.
            start_paragraph: Starting paragraph index for the returned window.
            count: Maximum number of paragraphs to return.

        Returns:
            Paragraph window with plain and annotated text.
        """

        doc, resolved = load_document(path)
        paragraphs = iter_paragraphs(doc)
        document_root = read_zip_xml(resolved, "word/document.xml")
        paragraph_xml = iter_document_paragraphs_xml(document_root) if document_root is not None else []
        end = min(len(paragraphs), start_paragraph + count)
        items = []
        for index in range(start_paragraph, end):
            revision_view = paragraph_revision_view(paragraph_xml[index]) if index < len(paragraph_xml) else None
            items.append(
                {
                    "index": index,
                    "text": revision_view["visible_text"] if revision_view else paragraphs[index].text,
                    "text_with_revisions": revision_view["annotated_text"] if revision_view else paragraphs[index].text,
                }
            )
        return {
            "path": str(resolved),
            "engine": "python-docx+lxml",
            "start_paragraph": start_paragraph,
            "returned": len(items),
            "total_paragraphs": len(paragraphs),
            "paragraphs": items,
        }

    @server.tool()
    @tool_response("find_text_occurrences")
    def find_text_occurrences(
        path: str,
        target_text: str,
        match_case: bool = False,
        find_whole_words_only: bool = False,
        max_results: int = 100,
        paragraph_index: int | None = None,
    ) -> dict[str, Any]:
        """Find text occurrences before replacement or comment insertion.

        Args:
            path: Path to the DOCX file.
            target_text: Text to search for.
            match_case: Whether matching should be case-sensitive.
            find_whole_words_only: Whether only whole words should match.
            max_results: Maximum number of occurrences to return.
            paragraph_index: Optional paragraph index to constrain the search.

        Returns:
            Matching paragraph indices, offsets and previews.
        """

        if max_results <= 0:
            raise ValueError("max_results must be > 0")
        doc, resolved = load_document(path)
        paragraphs = iter_paragraphs(doc)
        if paragraph_index is not None and (paragraph_index < 0 or paragraph_index >= len(paragraphs)):
            raise IndexError(f"Paragraph index out of range: {paragraph_index}")
        indexed_paragraphs = [(paragraph_index, paragraphs[paragraph_index])] if paragraph_index is not None else list(enumerate(paragraphs))
        matches: list[dict[str, Any]] = []
        for current_index, paragraph in indexed_paragraphs:
            paragraph_matches = matches_for_text(paragraph.text, target_text, match_case, find_whole_words_only)
            for match in paragraph_matches:
                matches.append(
                    {
                        "paragraph_index": current_index,
                        "start_offset": match.start(),
                        "end_offset": match.end(),
                        "text": match.group(0),
                        "paragraph_preview": paragraph.text[:300],
                    }
                )
                if len(matches) >= max_results:
                    return {"path": str(resolved), "engine": "python-docx", "returned": len(matches), "matches": matches}
        return {"path": str(resolved), "engine": "python-docx", "returned": len(matches), "matches": matches}

    @server.tool()
    @tool_response("find_paragraphs")
    def find_paragraphs(
        path: str,
        search_text: str | None = None,
        style_name: str | None = None,
        match_case: bool = False,
        max_results: int = 100,
    ) -> dict[str, Any]:
        """Find paragraphs by text snippet, style name, or both.

        Args:
            path: Path to the DOCX file.
            search_text: Optional text snippet to match inside paragraphs.
            style_name: Optional paragraph style name to match.
            match_case: Whether text matching should be case-sensitive.
            max_results: Maximum number of paragraphs to return.

        Returns:
            Matching paragraph descriptors without run-level payload.
        """

        if search_text is None and style_name is None:
            raise ValueError("Provide search_text or style_name")
        if max_results <= 0:
            raise ValueError("max_results must be > 0")
        doc, resolved = load_document(path)
        matches: list[dict[str, Any]] = []
        for index, paragraph in enumerate(iter_paragraphs(doc)):
            text_matches = True
            style_matches = True
            if search_text is not None:
                text_matches = bool(matches_for_text(paragraph.text, search_text, match_case, False))
            if style_name is not None:
                style_matches = bool(paragraph.style and paragraph.style.name == style_name)
            if text_matches and style_matches:
                matches.append(paragraph_to_dict(paragraph, index, include_runs=False))
                if len(matches) >= max_results:
                    break
        return {"path": str(resolved), "engine": "python-docx", "returned": len(matches), "paragraphs": matches}

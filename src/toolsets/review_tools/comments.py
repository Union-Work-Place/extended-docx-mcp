"""Comment-management tools exposed through FastMCP."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from config import W
from ops.package_io import read_zip_xml, resolve_path, save_zip_parts, serialize_xml
from ops.review import (
    add_comment_node,
    anchor_comment_to_paragraph,
    anchor_comment_to_range,
    ensure_comments_parts,
    find_text_range_xml,
    iter_document_paragraphs_xml,
    list_comments_xml,
    next_comment_id,
    w_attr,
)
from toolsets.response_schema import tool_response


def _resolve_comment_anchor(root: Any, paragraph_index: int | None, anchor_text: str | None) -> tuple[int, Any]:
    """Resolve a paragraph target for comment insertion.

    Args:
        root: Parsed ``word/document.xml`` root.
        paragraph_index: Optional paragraph index.
        anchor_text: Optional anchor text.

    Returns:
        Matching paragraph index and paragraph XML node.
    """

    paragraphs = iter_document_paragraphs_xml(root)
    if paragraph_index is not None:
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            raise IndexError(f"Paragraph index out of range: {paragraph_index}")
        return paragraph_index, paragraphs[paragraph_index]
    if anchor_text:
        matches = [(index, paragraph) for index, paragraph in enumerate(paragraphs) if anchor_text in "".join(paragraph.itertext())]
        if not matches:
            raise ValueError(f"Anchor text was not found in document: {anchor_text}")
        return matches[0]
    raise ValueError("Provide paragraph_index or anchor_text")


def register_comment_tools(server: FastMCP) -> None:
    """Register comment-related tools on a FastMCP server.

    Args:
        server: MCP server to populate.
    """

    @server.tool()
    @tool_response("add_comment")
    def add_comment(
        path: str,
        comment_text: str,
        author: str,
        initials: str,
        output_path: str | None = None,
        paragraph_index: int | None = None,
        anchor_text: str | None = None,
    ) -> dict[str, Any]:
        """Add a Word comment anchored to a paragraph.

        Args:
            path: Path to the DOCX file.
            comment_text: Comment text body.
            author: Comment author name.
            initials: Author initials.
            output_path: Optional alternate output path.
            paragraph_index: Optional absolute paragraph index.
            anchor_text: Optional paragraph anchor text.

        Returns:
            Summary of the comment insertion.
        """

        source_path = resolve_path(path)
        root = read_zip_xml(source_path, "word/document.xml")
        if root is None:
            raise ValueError("DOCX package is missing word/document.xml")
        actual_index, paragraph = _resolve_comment_anchor(root, paragraph_index, anchor_text)
        content_types, rels, comments_root = ensure_comments_parts(source_path)
        comment_id = next_comment_id(comments_root)
        add_comment_node(comments_root, comment_id, comment_text, author, initials)
        anchor_comment_to_paragraph(paragraph, comment_id)
        saved_to = save_zip_parts(
            source_path,
            output_path,
            {
                "[Content_Types].xml": serialize_xml(content_types),
                "word/_rels/document.xml.rels": serialize_xml(rels),
                "word/comments.xml": serialize_xml(comments_root),
                "word/document.xml": serialize_xml(root),
            },
        )
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "lxml-ooxml",
            "paragraph_index": actual_index,
            "comment_id": comment_id,
            "comment": comment_text,
        }

    @server.tool()
    @tool_response("add_comment_to_text_range")
    def add_comment_to_text_range(
        path: str,
        comment_text: str,
        author: str,
        initials: str,
        start_offset: int,
        end_offset: int,
        output_path: str | None = None,
        paragraph_index: int | None = None,
        anchor_text: str | None = None,
    ) -> dict[str, Any]:
        """Add a Word comment to an explicit text range inside a paragraph.

        Args:
            path: Path to the DOCX file.
            comment_text: Comment text body.
            author: Comment author name.
            initials: Author initials.
            start_offset: Inclusive character start offset.
            end_offset: Exclusive character end offset.
            output_path: Optional alternate output path.
            paragraph_index: Optional absolute paragraph index.
            anchor_text: Optional paragraph anchor text.

        Returns:
            Summary of the comment insertion.
        """

        source_path = resolve_path(path)
        root = read_zip_xml(source_path, "word/document.xml")
        if root is None:
            raise ValueError("DOCX package is missing word/document.xml")
        actual_index, paragraph = _resolve_comment_anchor(root, paragraph_index, anchor_text)
        content_types, rels, comments_root = ensure_comments_parts(source_path)
        comment_id = next_comment_id(comments_root)
        add_comment_node(comments_root, comment_id, comment_text, author, initials)
        anchor_comment_to_range(paragraph, comment_id, start_offset, end_offset)
        saved_to = save_zip_parts(
            source_path,
            output_path,
            {
                "[Content_Types].xml": serialize_xml(content_types),
                "word/_rels/document.xml.rels": serialize_xml(rels),
                "word/comments.xml": serialize_xml(comments_root),
                "word/document.xml": serialize_xml(root),
            },
        )
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "lxml-ooxml",
            "paragraph_index": actual_index,
            "comment_id": comment_id,
            "start_offset": start_offset,
            "end_offset": end_offset,
            "comment": comment_text,
        }

    @server.tool()
    @tool_response("add_comment_to_matching_text")
    def add_comment_to_matching_text(
        path: str,
        target_text: str,
        comment_text: str,
        author: str,
        initials: str,
        occurrence_index: int = 0,
        output_path: str | None = None,
        paragraph_index: int | None = None,
        anchor_text: str | None = None,
    ) -> dict[str, Any]:
        """Find text and attach a comment to the matching range.

        Args:
            path: Path to the DOCX file.
            target_text: Text to locate.
            comment_text: Comment text body.
            author: Comment author name.
            initials: Author initials.
            occurrence_index: Zero-based match occurrence.
            output_path: Optional alternate output path.
            paragraph_index: Optional paragraph restriction.
            anchor_text: Optional paragraph anchor text.

        Returns:
            Summary of the comment insertion.
        """

        source_path = resolve_path(path)
        root = read_zip_xml(source_path, "word/document.xml")
        if root is None:
            raise ValueError("DOCX package is missing word/document.xml")
        paragraph, actual_index, start_offset, end_offset = find_text_range_xml(
            root,
            target_text,
            occurrence_index=occurrence_index,
            paragraph_index=paragraph_index,
            anchor_text=anchor_text,
        )
        content_types, rels, comments_root = ensure_comments_parts(source_path)
        comment_id = next_comment_id(comments_root)
        add_comment_node(comments_root, comment_id, comment_text, author, initials)
        anchor_comment_to_range(paragraph, comment_id, start_offset, end_offset)
        saved_to = save_zip_parts(
            source_path,
            output_path,
            {
                "[Content_Types].xml": serialize_xml(content_types),
                "word/_rels/document.xml.rels": serialize_xml(rels),
                "word/comments.xml": serialize_xml(comments_root),
                "word/document.xml": serialize_xml(root),
            },
        )
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "lxml-ooxml",
            "paragraph_index": actual_index,
            "target_text": target_text,
            "comment_id": comment_id,
            "start_offset": start_offset,
            "end_offset": end_offset,
            "comment": comment_text,
        }

    @server.tool()
    @tool_response("add_comment_reply")
    def add_comment_reply(
        path: str,
        comment_index: int,
        reply_text: str,
        author: str,
        initials: str,
        output_path: str | None = None,
    ) -> dict[str, Any]:
        """Reply to an existing comment.

        Args:
            path: Path to the DOCX file.
            comment_index: Comment index in ``list_comments`` order.
            reply_text: Reply body text.
            author: Reply author name.
            initials: Reply author initials.
            output_path: Optional alternate output path.

        Returns:
            Summary of the reply insertion.
        """

        source_path = resolve_path(path)
        content_types, rels, comments_root = ensure_comments_parts(source_path)
        comments = comments_root.findall(f"{W}comment")
        if comment_index < 0 or comment_index >= len(comments):
            raise IndexError(f"Comment index out of range: {comment_index}")
        parent_id = int(comments[comment_index].get(w_attr("id"), "0"))
        reply_id = next_comment_id(comments_root)
        add_comment_node(comments_root, reply_id, reply_text, author, initials, parent_id=parent_id)
        saved_to = save_zip_parts(
            source_path,
            output_path,
            {
                "[Content_Types].xml": serialize_xml(content_types),
                "word/_rels/document.xml.rels": serialize_xml(rels),
                "word/comments.xml": serialize_xml(comments_root),
            },
        )
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "lxml-ooxml",
            "comment_index": comment_index,
            "parent_comment_id": parent_id,
            "reply_comment_id": reply_id,
            "reply": reply_text,
        }

    @server.tool()
    @tool_response("list_comments")
    def list_comments(path: str, include_replies: bool = True) -> dict[str, Any]:
        """Return comments and replies with metadata.

        Args:
            path: Path to the DOCX file.
            include_replies: Whether reply comments should be included.

        Returns:
            Serialized comments payload.
        """

        resolved = resolve_path(path)
        comments = list_comments_xml(resolved)
        if not include_replies:
            comments = [comment for comment in comments if comment.get("parent_id") is None]
        return {"path": str(resolved), "engine": "lxml-ooxml", "comments": comments}

"""Document creation and paragraph-level editing tools exposed through FastMCP."""

from __future__ import annotations

import datetime as dt
from typing import Any

from lxml import etree
from mcp.server.fastmcp import FastMCP

from docx import Document

from docx.shared import Pt

from ops.package_io import load_document, load_or_create_document, read_zip_xml, resolve_path, save_document, save_zip_parts, serialize_xml
from ops.review import (
    append_revision_pair,
    canonical_paragraph_refs,
    clear_paragraph_content,
    enable_track_revisions_part,
    iter_document_paragraphs_xml,
    list_revisions_xml,
    new_text_run,
    new_w_element,
    next_revision_id,
    replace_in_paragraph_tracked,
    w_attr,
)
from ops.structure_ops import (
    apply_run_format,
    find_paragraph,
    insert_paragraph_after,
    iter_paragraphs,
    iter_tables,
    validate_structured_blocks,
    write_structured_block,
)
from ops.text_ops import normalize_mapping_value, replace_in_paragraph_plain
from config import PX_ALIGNMENTS
from toolsets.response_schema import tool_response


def register_content_editing_tools(server: FastMCP) -> None:
    """Register creation and paragraph-editing tools on a FastMCP server.

    Args:
        server: MCP server to populate.
    """

    @server.tool()
    @tool_response("write_docx")
    def write_docx(
        path: str,
        blocks: list[dict[str, Any]],
        output_path: str | None = None,
        mode: str = "replace",
        track_changes: bool = False,
        author: str = "DOCX MCP",
        title: str | None = None,
        subject: str | None = None,
    ) -> dict[str, Any]:
        """Create, replace or append DOCX content from structured blocks.

        Args:
            path: Target DOCX path.
            blocks: Structured content blocks to write.
            output_path: Optional alternate output path.
            mode: Either ``replace`` or ``append``.
            track_changes: Reserved compatibility flag for future append tracking.
            author: Reserved author label for future append tracking.
            title: Optional document title property.
            subject: Optional document subject property.

        Returns:
            Summary of the write operation.
        """

        normalized_mode = mode.lower()
        if normalized_mode not in {"replace", "append"}:
            raise ValueError("mode must be 'replace' or 'append'")
        validate_structured_blocks(blocks)
        existing_doc, source_path, created_from_missing = load_or_create_document(path)
        doc = existing_doc if normalized_mode == "append" else Document()
        if title is not None:
            doc.core_properties.title = title
        if subject is not None:
            doc.core_properties.subject = subject
        for index, block in enumerate(blocks):
            write_structured_block(doc, block, index)
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "mode": normalized_mode,
            "created_from_missing": created_from_missing,
            "blocks_written": len(blocks),
            "paragraphs": len(iter_paragraphs(doc)),
            "tables": len(iter_tables(doc)),
            "track_changes_requested": track_changes,
            "track_changes_supported": False,
            "author": author if track_changes else None,
        }

    @server.tool()
    @tool_response("replace_text")
    def replace_text(
        path: str,
        find_text: str,
        replace_with: str,
        output_path: str | None = None,
        match_case: bool = False,
        find_whole_words_only: bool = False,
        track_changes: bool = True,
        author: str = "DOCX MCP",
    ) -> dict[str, Any]:
        """Replace text in paragraphs and table cells.

        Args:
            path: Path to the DOCX file.
            find_text: Text to search for.
            replace_with: Replacement text.
            output_path: Optional alternate output path.
            match_case: Whether matching should be case-sensitive.
            find_whole_words_only: Whether only whole words should match.
            track_changes: Whether to write OOXML tracked changes instead of plain replacement.
            author: Revision author label for tracked replacements.

        Returns:
            Summary of the replacement operation.
        """

        if track_changes:
            source_path = resolve_path(path)
            if not source_path.exists():
                raise FileNotFoundError(f"Document was not found: {source_path}")
            root = read_zip_xml(source_path, "word/document.xml")
            if root is None:
                raise ValueError("DOCX package is missing word/document.xml")
            changed = 0
            paragraphs_changed = 0
            revision_id = next_revision_id(root)
            timestamp = dt.datetime.now(dt.timezone.utc).isoformat()
            for paragraph in iter_document_paragraphs_xml(root):
                paragraph_changed, revision_id = replace_in_paragraph_tracked(
                    paragraph,
                    find_text,
                    replace_with,
                    match_case,
                    find_whole_words_only,
                    revision_id,
                    author,
                    timestamp,
                )
                if paragraph_changed:
                    changed += paragraph_changed
                    paragraphs_changed += 1
            parts = {"word/document.xml": serialize_xml(root)}
            settings_part = enable_track_revisions_part(source_path)
            if settings_part is not None:
                parts["word/settings.xml"] = settings_part
            saved_to = save_zip_parts(source_path, output_path, parts)
            return {
                "path": str(source_path),
                "saved_to": str(saved_to),
                "engine": "lxml-ooxml",
                "replacements": changed,
                "paragraphs_changed": paragraphs_changed,
                "track_changes_requested": True,
                "track_changes_supported": True,
                "revisions": len(list_revisions_xml(saved_to)),
            }

        doc, source_path = load_document(path)
        changed = 0
        paragraphs_changed = 0
        for paragraph in iter_paragraphs(doc):
            paragraph_changed = replace_in_paragraph_plain(
                paragraph,
                find_text,
                replace_with,
                match_case=match_case,
                find_whole_words_only=find_whole_words_only,
            )
            if paragraph_changed:
                changed += paragraph_changed
                paragraphs_changed += 1
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "replacements": changed,
            "paragraphs_changed": paragraphs_changed,
            "track_changes_requested": False,
            "track_changes_supported": True,
            "revisions": len(list_revisions_xml(saved_to)),
        }

    @server.tool()
    @tool_response("insert_paragraph")
    def insert_paragraph(
        path: str,
        text: str,
        output_path: str | None = None,
        after_paragraph: int | None = None,
        anchor_text: str | None = None,
        track_changes: bool = True,
        author: str = "DOCX MCP",
    ) -> dict[str, Any]:
        """Insert a paragraph after another paragraph.

        Args:
            path: Path to the DOCX file.
            text: Inserted paragraph text.
            output_path: Optional alternate output path.
            after_paragraph: Optional absolute paragraph index to insert after.
            anchor_text: Optional paragraph anchor text.
            track_changes: Whether to write the insertion as a revision.
            author: Revision author label.

        Returns:
            Summary of the insertion operation.
        """

        if track_changes:
            source_path = resolve_path(path)
            doc, _ = load_document(path)
            root = read_zip_xml(source_path, "word/document.xml")
            if root is None:
                raise ValueError("DOCX package is missing word/document.xml")
            paragraphs = canonical_paragraph_refs(doc, root)
            if after_paragraph is not None:
                if after_paragraph < 0 or after_paragraph >= len(paragraphs):
                    raise IndexError(f"Paragraph index out of range: {after_paragraph}")
                target = paragraphs[after_paragraph].paragraph
                actual_index = after_paragraph
            elif anchor_text:
                matches = [(ref.index, ref.paragraph) for ref in paragraphs if anchor_text in "".join(ref.paragraph.itertext())]
                if not matches:
                    raise ValueError(f"Anchor text was not found in document: {anchor_text}")
                actual_index, target = matches[0]
            else:
                raise ValueError("Provide after_paragraph or anchor_text")
            paragraph_xml = etree.Element(target.tag)
            revision = new_w_element(
                "ins",
                {
                    w_attr("id"): next_revision_id(root),
                    w_attr("author"): author,
                    w_attr("date"): dt.datetime.now(dt.timezone.utc).isoformat(),
                },
            )
            revision.append(new_text_run(text))
            paragraph_xml.append(revision)
            target.addnext(paragraph_xml)
            parts = {"word/document.xml": serialize_xml(root)}
            settings_part = enable_track_revisions_part(source_path)
            if settings_part is not None:
                parts["word/settings.xml"] = settings_part
            saved_to = save_zip_parts(source_path, output_path, parts)
            return {
                "path": str(source_path),
                "saved_to": str(saved_to),
                "engine": "lxml-ooxml",
                "inserted_after": actual_index,
                "text": text,
                "track_changes_supported": True,
                "revisions": len(list_revisions_xml(saved_to)),
            }

        doc, source_path = load_document(path)
        target_paragraph, target_index = find_paragraph(doc, after_paragraph, anchor_text)
        insert_paragraph_after(target_paragraph, text)
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "inserted_after": target_index,
            "text": text,
            "track_changes_supported": True,
        }

    @server.tool()
    @tool_response("delete_paragraph")
    def delete_paragraph(
        path: str,
        paragraph_index: int,
        output_path: str | None = None,
        track_changes: bool = True,
        author: str = "DOCX MCP",
    ) -> dict[str, Any]:
        """Delete a paragraph by index.

        Args:
            path: Path to the DOCX file.
            paragraph_index: Absolute paragraph index to delete.
            output_path: Optional alternate output path.
            track_changes: Whether to store the deletion as a revision.
            author: Revision author label.

        Returns:
            Summary of the deletion operation.
        """

        if track_changes:
            source_path = resolve_path(path)
            doc, _ = load_document(path)
            root = read_zip_xml(source_path, "word/document.xml")
            if root is None:
                raise ValueError("DOCX package is missing word/document.xml")
            paragraphs = canonical_paragraph_refs(doc, root)
            if paragraph_index < 0 or paragraph_index >= len(paragraphs):
                raise IndexError(f"Paragraph index out of range: {paragraph_index}")
            paragraph = paragraphs[paragraph_index].paragraph
            deleted_text = "".join(paragraph.itertext())
            clear_paragraph_content(paragraph)
            deletion = new_w_element(
                "del",
                {
                    w_attr("id"): next_revision_id(root),
                    w_attr("author"): author,
                    w_attr("date"): dt.datetime.now(dt.timezone.utc).isoformat(),
                },
            )
            deletion.append(new_text_run(deleted_text, deleted=True))
            paragraph.append(deletion)
            parts = {"word/document.xml": serialize_xml(root)}
            settings_part = enable_track_revisions_part(source_path)
            if settings_part is not None:
                parts["word/settings.xml"] = settings_part
            saved_to = save_zip_parts(source_path, output_path, parts)
            return {
                "path": str(source_path),
                "saved_to": str(saved_to),
                "engine": "lxml-ooxml",
                "deleted_paragraph": paragraph_index,
                "text": deleted_text,
                "track_changes_supported": True,
                "revisions": len(list_revisions_xml(saved_to)),
            }

        doc, source_path = load_document(path)
        paragraph, actual_index = find_paragraph(doc, paragraph_index, None)
        deleted_text = paragraph.text
        paragraph._element.getparent().remove(paragraph._element)
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "deleted_paragraph": actual_index,
            "text": deleted_text,
            "track_changes_supported": True,
        }

    @server.tool()
    @tool_response("set_paragraph_format")
    def set_paragraph_format(
        path: str,
        output_path: str | None = None,
        paragraph_index: int | None = None,
        anchor_text: str | None = None,
        alignment: str | None = None,
        keep_with_next: bool | None = None,
        first_line_indent_points: float | None = None,
        left_indent_points: float | None = None,
        right_indent_points: float | None = None,
        space_before_points: float | None = None,
        space_after_points: float | None = None,
    ) -> dict[str, Any]:
        """Apply paragraph-level formatting.

        Args:
            path: Path to the DOCX file.
            output_path: Optional alternate output path.
            paragraph_index: Optional absolute paragraph index.
            anchor_text: Optional paragraph anchor text.
            alignment: Optional public alignment value.
            keep_with_next: Optional keep-with-next flag.
            first_line_indent_points: Optional first-line indent in points.
            left_indent_points: Optional left indent in points.
            right_indent_points: Optional right indent in points.
            space_before_points: Optional spacing before in points.
            space_after_points: Optional spacing after in points.

        Returns:
            Summary of the formatting operation.
        """

        doc, source_path = load_document(path)
        paragraph, actual_index = find_paragraph(doc, paragraph_index, anchor_text)
        paragraph_format = paragraph.paragraph_format
        if alignment is not None:
            paragraph_format.alignment = normalize_mapping_value(alignment, PX_ALIGNMENTS, "alignment")
        if keep_with_next is not None:
            paragraph_format.keep_with_next = keep_with_next
        if first_line_indent_points is not None:
            paragraph_format.first_line_indent = Pt(first_line_indent_points)
        if left_indent_points is not None:
            paragraph_format.left_indent = Pt(left_indent_points)
        if right_indent_points is not None:
            paragraph_format.right_indent = Pt(right_indent_points)
        if space_before_points is not None:
            paragraph_format.space_before = Pt(space_before_points)
        if space_after_points is not None:
            paragraph_format.space_after = Pt(space_after_points)
        saved_to = save_document(doc, source_path, output_path)
        return {"path": str(source_path), "saved_to": str(saved_to), "engine": "python-docx", "paragraph_index": actual_index, "text": paragraph.text}

    @server.tool()
    @tool_response("set_run_format")
    def set_run_format(
        path: str,
        paragraph_index: int,
        run_index: int = 0,
        output_path: str | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: str | None = None,
        font_name: str | None = None,
        font_size_points: float | None = None,
        all_caps: bool | None = None,
    ) -> dict[str, Any]:
        """Apply run-level formatting.

        Args:
            path: Path to the DOCX file.
            paragraph_index: Paragraph containing the run.
            run_index: Run index inside the paragraph.
            output_path: Optional alternate output path.
            bold: Optional bold flag.
            italic: Optional italic flag.
            underline: Optional underline mode.
            font_name: Optional font family.
            font_size_points: Optional font size in points.
            all_caps: Optional all-caps flag.

        Returns:
            Summary of the formatting operation.
        """

        doc, source_path = load_document(path)
        paragraph, actual_index = find_paragraph(doc, paragraph_index, None)
        runs = list(paragraph.runs)
        if run_index < 0 or run_index >= len(runs):
            raise IndexError(f"Run index out of range: {run_index}")
        run = runs[run_index]
        run_data: dict[str, Any] = {
            "font_name": font_name,
            "font_size_points": font_size_points,
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "all_caps": all_caps,
        }
        apply_run_format(run, run_data)
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "paragraph_index": actual_index,
            "run_index": run_index,
            "text": run.text,
        }

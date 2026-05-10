"""Table-focused content tools exposed through FastMCP."""

from __future__ import annotations

import datetime as dt
from typing import Any

from mcp.server.fastmcp import FastMCP

from config import PX_TABLE_ALIGNMENTS, W
from ops.package_io import load_document, read_zip_xml, resolve_path, save_document, save_zip_parts, serialize_xml
from ops.review import (
    append_revision_pair,
    clear_paragraph_content,
    enable_track_revisions_part,
    iter_document_tables_xml,
    list_revisions_xml,
    new_w_element,
    next_revision_id,
    paragraph_text_xml,
)
from ops.structure_ops import (
    find_paragraph,
    insert_table_after,
    iter_paragraphs,
    iter_tables,
    paragraph_to_dict,
    table_to_dict,
    write_cell_value,
)
from ops.text_ops import normalize_mapping_value


def _get_table_cell(table: Any, row_index: int, cell_index: int) -> Any:
    """Return a validated table cell from a ``python-docx`` table.

    Args:
        table: ``python-docx`` table object.
        row_index: Row index inside the table.
        cell_index: Cell index inside the row.

    Returns:
        Target cell object.
    """

    if row_index < 0 or row_index >= len(table.rows):
        raise IndexError(f"Row index out of range: {row_index}")
    row = table.rows[row_index]
    if cell_index < 0 or cell_index >= len(row.cells):
        raise IndexError(f"Cell index out of range: {cell_index}")
    return row.cells[cell_index]


def register_table_tools(server: FastMCP) -> None:
    """Register table-related tools on a FastMCP server.

    Args:
        server: MCP server to populate.
    """

    @server.tool()
    def list_tables(path: str, include_cells: bool = True) -> dict[str, Any]:
        """List document tables with optional cell contents.

        Args:
            path: Path to the DOCX file.
            include_cells: Whether to include cell data for each row.

        Returns:
            Serialized table descriptors.
        """

        doc, resolved = load_document(path)
        tables = iter_tables(doc)
        return {"path": str(resolved), "engine": "python-docx", "tables": [table_to_dict(table, index, include_cells) for index, table in enumerate(tables)]}

    @server.tool()
    def get_table_cell_content(
        path: str,
        table_index: int,
        row_index: int,
        cell_index: int,
        include_runs: bool = True,
    ) -> dict[str, Any]:
        """Return the content of one table cell as paragraph descriptors.

        Args:
            path: Path to the DOCX file.
            table_index: Table index in document order.
            row_index: Row index inside the table.
            cell_index: Cell index inside the row.
            include_runs: Whether to include run-level metadata.

        Returns:
            Structured cell payload with paragraphs and visible text.
        """

        doc, resolved = load_document(path)
        tables = iter_tables(doc)
        if table_index < 0 or table_index >= len(tables):
            raise IndexError(f"Table index out of range: {table_index}")
        cell = _get_table_cell(tables[table_index], row_index, cell_index)
        paragraphs = [paragraph_to_dict(paragraph, index, include_runs) for index, paragraph in enumerate(cell.paragraphs)]
        return {
            "path": str(resolved),
            "engine": "python-docx",
            "table_index": table_index,
            "row_index": row_index,
            "cell_index": cell_index,
            "text": cell.text,
            "paragraphs": paragraphs,
        }

    @server.tool()
    def insert_table(
        path: str,
        data: list[list[str]],
        output_path: str | None = None,
        after_paragraph: int | None = None,
        anchor_text: str | None = None,
        style_name: str | None = None,
        alignment: str | None = None,
        track_changes: bool = True,
        author: str = "DOCX MCP",
    ) -> dict[str, Any]:
        """Insert a table after a paragraph.

        Args:
            path: Path to the DOCX file.
            data: Two-dimensional cell matrix.
            output_path: Optional alternate output path.
            after_paragraph: Optional absolute paragraph index.
            anchor_text: Optional paragraph anchor text.
            style_name: Optional table style to apply after insertion.
            alignment: Optional public table alignment.
            track_changes: Reserved compatibility flag for future tracked table insertion.
            author: Reserved author label for future tracked table insertion.

        Returns:
            Summary of the insertion operation.
        """

        if not data or not data[0]:
            raise ValueError("Table data must contain at least one row and one cell")
        doc, source_path = load_document(path)
        paragraph, actual_index = find_paragraph(doc, after_paragraph, anchor_text)
        table = insert_table_after(paragraph, data)
        if style_name is not None:
            table.style = style_name
        if alignment is not None:
            table.alignment = normalize_mapping_value(alignment, PX_TABLE_ALIGNMENTS, "table alignment")
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "inserted_after_paragraph": actual_index,
            "rows": len(data),
            "columns": max(len(row) for row in data),
            "track_changes_requested": track_changes,
            "track_changes_supported": False,
            "author": author if track_changes else None,
        }

    @server.tool()
    def update_table_cell(
        path: str,
        table_index: int,
        row_index: int,
        cell_index: int,
        text: str,
        output_path: str | None = None,
        track_changes: bool = True,
        author: str = "DOCX MCP",
    ) -> dict[str, Any]:
        """Replace text inside a table cell.

        Args:
            path: Path to the DOCX file.
            table_index: Table index in document order.
            row_index: Row index inside the table.
            cell_index: Cell index inside the row.
            text: Replacement text.
            output_path: Optional alternate output path.
            track_changes: Whether to store the change as a tracked revision.
            author: Revision author label.

        Returns:
            Summary of the cell update operation.
        """

        if track_changes:
            source_path = resolve_path(path)
            root = read_zip_xml(source_path, "word/document.xml")
            if root is None:
                raise ValueError("DOCX package is missing word/document.xml")
            tables = iter_document_tables_xml(root)
            if table_index < 0 or table_index >= len(tables):
                raise IndexError(f"Table index out of range: {table_index}")
            rows = tables[table_index].findall(f"{W}tr")
            if row_index < 0 or row_index >= len(rows):
                raise IndexError(f"Row index out of range: {row_index}")
            cells = rows[row_index].findall(f"{W}tc")
            if cell_index < 0 or cell_index >= len(cells):
                raise IndexError(f"Cell index out of range: {cell_index}")
            paragraph = cells[cell_index].find(f"{W}p")
            if paragraph is None:
                paragraph = new_w_element("p")
                cells[cell_index].append(paragraph)
            old_text = paragraph_text_xml(paragraph)
            clear_paragraph_content(paragraph)
            append_revision_pair(paragraph, old_text, text, next_revision_id(root), author, dt.datetime.now(dt.timezone.utc).isoformat())
            parts = {"word/document.xml": serialize_xml(root)}
            settings_part = enable_track_revisions_part(source_path)
            if settings_part is not None:
                parts["word/settings.xml"] = settings_part
            saved_to = save_zip_parts(source_path, output_path, parts)
            return {
                "path": str(source_path),
                "saved_to": str(saved_to),
                "engine": "lxml-ooxml",
                "table_index": table_index,
                "row_index": row_index,
                "cell_index": cell_index,
                "text": text,
                "track_changes_supported": True,
                "revisions": len(list_revisions_xml(saved_to)),
            }

        doc, source_path = load_document(path)
        tables = iter_tables(doc)
        if table_index < 0 or table_index >= len(tables):
            raise IndexError(f"Table index out of range: {table_index}")
        cell = _get_table_cell(tables[table_index], row_index, cell_index)
        write_cell_value(cell, text)
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "table_index": table_index,
            "row_index": row_index,
            "cell_index": cell_index,
            "text": text,
            "track_changes_supported": True,
        }

    @server.tool()
    def set_table_format(
        path: str,
        table_index: int,
        output_path: str | None = None,
        style_name: str | None = None,
        alignment: str | None = None,
        title: str | None = None,
        description: str | None = None,
        allow_auto_fit: bool | None = None,
        auto_fit_behavior: str | None = None,
        left_indent_points: float | None = None,
    ) -> dict[str, Any]:
        """Apply supported table formatting.

        Args:
            path: Path to the DOCX file.
            table_index: Table index in document order.
            output_path: Optional alternate output path.
            style_name: Optional Word table style.
            alignment: Optional public table alignment.
            title: Unsupported compatibility field.
            description: Unsupported compatibility field.
            allow_auto_fit: Optional ``python-docx`` auto-fit toggle.
            auto_fit_behavior: Unsupported compatibility field.
            left_indent_points: Unsupported compatibility field.

        Returns:
            Summary of the formatting operation.
        """

        doc, source_path = load_document(path)
        tables = iter_tables(doc)
        if table_index < 0 or table_index >= len(tables):
            raise IndexError(f"Table index out of range: {table_index}")
        table = tables[table_index]
        if style_name is not None:
            table.style = style_name
        if alignment is not None:
            table.alignment = normalize_mapping_value(alignment, PX_TABLE_ALIGNMENTS, "table alignment")
        if allow_auto_fit is not None:
            table.autofit = allow_auto_fit
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "table": table_to_dict(table, table_index, include_cells=False),
            "unsupported_options_ignored": {
                "title": title is not None,
                "description": description is not None,
                "auto_fit_behavior": auto_fit_behavior is not None,
                "left_indent_points": left_indent_points is not None,
            },
        }


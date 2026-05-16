"""Batch editing tools exposed through FastMCP."""

from __future__ import annotations

from typing import Any

from docx.shared import Pt
from mcp.server.fastmcp import FastMCP

from config import PX_ALIGNMENTS
from ops.document_ops import iter_paragraphs, iter_tables
from ops.package_io import load_document, save_document
from ops.text_ops import normalize_mapping_value, replace_in_paragraph_plain
from toolsets.response_schema import tool_response


def _get_table_cell(table: Any, row_index: int, cell_index: int) -> Any:
    if row_index < 0 or row_index >= len(table.rows):
        raise IndexError(f"Row index out of range: {row_index}")
    row = table.rows[row_index]
    if cell_index < 0 or cell_index >= len(row.cells):
        raise IndexError(f"Cell index out of range: {cell_index}")
    return row.cells[cell_index]


def _validate_replacements_payload(replacements: Any) -> list[dict[str, str]]:
    if not isinstance(replacements, list) or not replacements:
        raise ValueError("replacements must contain at least one item")
    normalized: list[dict[str, str]] = []
    for index, replacement in enumerate(replacements):
        if not isinstance(replacement, dict):
            raise ValueError(f"Replacement at index {index} must be an object")
        find_text = replacement.get("find_text")
        if not isinstance(find_text, str) or not find_text:
            raise ValueError(f"Replacement at index {index} must include a non-empty find_text")
        normalized.append(
            {
                "find_text": find_text,
                "replace_with": str(replacement.get("replace_with", "")),
            }
        )
    return normalized


def _validate_table_updates_payload(updates: Any) -> list[dict[str, Any]]:
    if not isinstance(updates, list) or not updates:
        raise ValueError("updates must contain at least one item")
    normalized: list[dict[str, Any]] = []
    required_fields = ("table_index", "row_index", "cell_index")
    for index, update in enumerate(updates):
        if not isinstance(update, dict):
            raise ValueError(f"Update at index {index} must be an object")
        missing_fields = [field for field in required_fields if field not in update]
        if missing_fields:
            missing = ", ".join(missing_fields)
            raise ValueError(f"Update at index {index} is missing required fields: {missing}")
        try:
            normalized.append(
                {
                    "table_index": int(update["table_index"]),
                    "row_index": int(update["row_index"]),
                    "cell_index": int(update["cell_index"]),
                    "text": str(update.get("text", "")),
                }
            )
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"Update at index {index} must include integer table_index, row_index and cell_index values"
            ) from exc
    return normalized


def register_batch_tools(server: FastMCP) -> None:
    """Register batch editing helpers."""

    @server.tool()
    @tool_response("batch_replace_text")
    def batch_replace_text(
        path: str,
        replacements: list[Any],
        output_path: str | None = None,
        match_case: bool = False,
        find_whole_words_only: bool = False,
    ) -> dict[str, Any]:
        replacements = _validate_replacements_payload(replacements)
        doc, source_path = load_document(path)
        paragraphs = iter_paragraphs(doc)
        total_replacements = 0
        for replacement in replacements:
            find_text = str(replacement.get("find_text", ""))
            replace_with = str(replacement.get("replace_with", ""))
            for paragraph in paragraphs:
                total_replacements += replace_in_paragraph_plain(
                    paragraph,
                    find_text,
                    replace_with,
                    match_case=match_case,
                    find_whole_words_only=find_whole_words_only,
                )
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "operations": len(replacements),
            "replacements": total_replacements,
        }

    @server.tool()
    @tool_response("batch_set_paragraph_format")
    def batch_set_paragraph_format(
        path: str,
        paragraph_indices: list[int],
        output_path: str | None = None,
        alignment: str | None = None,
        keep_with_next: bool | None = None,
        left_indent_points: float | None = None,
        right_indent_points: float | None = None,
        space_before_points: float | None = None,
        space_after_points: float | None = None,
    ) -> dict[str, Any]:
        if not paragraph_indices:
            raise ValueError("paragraph_indices must contain at least one item")
        doc, source_path = load_document(path)
        paragraphs = iter_paragraphs(doc)
        for paragraph_index in paragraph_indices:
            if paragraph_index < 0 or paragraph_index >= len(paragraphs):
                raise IndexError(f"Paragraph index out of range: {paragraph_index}")
            paragraph_format = paragraphs[paragraph_index].paragraph_format
            if alignment is not None:
                paragraph_format.alignment = normalize_mapping_value(alignment, PX_ALIGNMENTS, "alignment")
            if keep_with_next is not None:
                paragraph_format.keep_with_next = keep_with_next
            if left_indent_points is not None:
                paragraph_format.left_indent = Pt(left_indent_points)
            if right_indent_points is not None:
                paragraph_format.right_indent = Pt(right_indent_points)
            if space_before_points is not None:
                paragraph_format.space_before = Pt(space_before_points)
            if space_after_points is not None:
                paragraph_format.space_after = Pt(space_after_points)
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "paragraph_indices": paragraph_indices,
            "updated": len(paragraph_indices),
        }

    @server.tool()
    @tool_response("batch_update_table_cells")
    def batch_update_table_cells(
        path: str,
        updates: list[Any],
        output_path: str | None = None,
    ) -> dict[str, Any]:
        updates = _validate_table_updates_payload(updates)
        doc, source_path = load_document(path)
        tables = iter_tables(doc)
        for update in updates:
            table_index = update["table_index"]
            row_index = update["row_index"]
            cell_index = update["cell_index"]
            if table_index < 0 or table_index >= len(tables):
                raise IndexError(f"Table index out of range: {table_index}")
            cell = _get_table_cell(tables[table_index], row_index, cell_index)
            cell.text = update["text"]
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "updated": len(updates),
        }

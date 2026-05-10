"""Style-related MCP tools."""

from __future__ import annotations

from typing import Any

from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from mcp.server.fastmcp import FastMCP

from config import PX_ALIGNMENTS, PX_STYLE_TYPES
from ops.document_ops import (
    ensure_style_type_is_paragraph,
    find_paragraph,
    normalize_mapping_value,
    style_to_dict,
)
from ops.package_io import load_document, save_document


def register_style_tools(server: FastMCP) -> None:
    """Register style-related MCP tools.

    Args:
        server: MCP server to populate.
    """

    @server.tool()
    def list_styles(path: str, style_type: str | None = None, include_builtin: bool = True) -> dict[str, Any]:
        """List Word styles available in the document.

        Args:
            path: Path to the DOCX file.
            style_type: Optional public style type filter.
            include_builtin: Whether built-in styles should be included.

        Returns:
            Serialized style descriptors.
        """

        doc, resolved = load_document(path)
        filtered_type = None
        if style_type is not None:
            filtered_type = normalize_mapping_value(style_type, PX_STYLE_TYPES, "style type")
        styles = []
        for style in doc.styles:
            if filtered_type is not None and style.type != filtered_type:
                continue
            if not include_builtin and style.builtin:
                continue
            styles.append(style_to_dict(style))
        return {"path": str(resolved), "engine": "python-docx", "styles": styles}

    @server.tool()
    def create_or_update_style(
        path: str,
        style_name: str,
        style_type: str = "paragraph",
        output_path: str | None = None,
        base_style_name: str | None = None,
        font_name: str | None = None,
        font_size_points: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        all_caps: bool | None = None,
        alignment: str | None = None,
        keep_with_next: bool | None = None,
        space_before_points: float | None = None,
        space_after_points: float | None = None,
    ) -> dict[str, Any]:
        """Create or update a document style.

        Args:
            path: Path to the DOCX file.
            style_name: Style name to create or update.
            style_type: Public style type label.
            output_path: Optional alternate output path.
            base_style_name: Optional base style name.
            font_name: Optional font family.
            font_size_points: Optional font size in points.
            bold: Optional bold flag.
            italic: Optional italic flag.
            all_caps: Optional all-caps flag.
            alignment: Optional paragraph alignment for paragraph styles.
            keep_with_next: Optional keep-with-next flag.
            space_before_points: Optional spacing before in points.
            space_after_points: Optional spacing after in points.

        Returns:
            Summary of the style mutation.
        """

        normalized_type = normalize_mapping_value(style_type, PX_STYLE_TYPES, "style type")
        doc, source_path = load_document(path)
        style = next((candidate for candidate in doc.styles if candidate.name == style_name), None)
        created = style is None
        if created:
            style = doc.styles.add_style(style_name, normalized_type)
        if base_style_name is not None:
            style.base_style = doc.styles[base_style_name]
        if font_name is not None:
            style.font.name = font_name
        if font_size_points is not None:
            style.font.size = Pt(font_size_points)
        if bold is not None:
            style.font.bold = bold
        if italic is not None:
            style.font.italic = italic
        if all_caps is not None:
            style.font.all_caps = all_caps
        if any(value is not None for value in (alignment, keep_with_next, space_before_points, space_after_points)):
            ensure_style_type_is_paragraph(style)
            if alignment is not None:
                style.paragraph_format.alignment = normalize_mapping_value(alignment, PX_ALIGNMENTS, "alignment")
            if keep_with_next is not None:
                style.paragraph_format.keep_with_next = keep_with_next
            if space_before_points is not None:
                style.paragraph_format.space_before = Pt(space_before_points)
            if space_after_points is not None:
                style.paragraph_format.space_after = Pt(space_after_points)
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "created": created,
            "style": style_to_dict(style),
        }

    @server.tool()
    def apply_paragraph_style(
        path: str,
        style_name: str,
        output_path: str | None = None,
        paragraph_index: int | None = None,
        anchor_text: str | None = None,
    ) -> dict[str, Any]:
        """Apply a Word paragraph style to a paragraph.

        Args:
            path: Path to the DOCX file.
            style_name: Existing style name to apply.
            output_path: Optional alternate output path.
            paragraph_index: Optional absolute paragraph index.
            anchor_text: Optional paragraph anchor text.

        Returns:
            Summary of the style application.
        """

        doc, source_path = load_document(path)
        if style_name not in [style.name for style in doc.styles]:
            raise ValueError(f"Style was not found: {style_name}")
        paragraph, actual_index = find_paragraph(doc, paragraph_index, anchor_text)
        paragraph.style = style_name
        saved_to = save_document(doc, source_path, output_path)
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "paragraph_index": actual_index,
            "style_name": style_name,
            "text": paragraph.text,
        }



"""Section and page setup MCP tools."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from ops.document_ops import section_to_dict, update_section_page_setup
from ops.package_io import load_document, save_document


def register_section_tools(server: FastMCP) -> None:
    """Register section and page setup tools.

    Args:
        server: MCP server to populate.
    """

    @server.tool()
    def list_sections(path: str) -> dict[str, Any]:
        """List document sections and page setup information.

        Args:
            path: Path to the DOCX file.

        Returns:
            Serialized section descriptors.
        """

        doc, resolved = load_document(path)
        return {"path": str(resolved), "engine": "python-docx", "sections": [section_to_dict(section, index) for index, section in enumerate(doc.sections)]}

    @server.tool()
    def set_section_page_setup(
        path: str,
        section_index: int,
        output_path: str | None = None,
        orientation: str | None = None,
        paper_size: str | None = None,
        section_start: str | None = None,
        left_margin_points: float | None = None,
        right_margin_points: float | None = None,
        top_margin_points: float | None = None,
        bottom_margin_points: float | None = None,
        different_first_page_header_footer: bool | None = None,
        restart_page_numbering: bool | None = None,
    ) -> dict[str, Any]:
        """Update page setup for one document section.

        Args:
            path: Path to the DOCX file.
            section_index: Absolute section index.
            output_path: Optional alternate output path.
            orientation: Optional public orientation value.
            paper_size: Optional public paper size value.
            section_start: Optional public section break value.
            left_margin_points: Optional left margin in points.
            right_margin_points: Optional right margin in points.
            top_margin_points: Optional top margin in points.
            bottom_margin_points: Optional bottom margin in points.
            different_first_page_header_footer: Optional header/footer flag.
            restart_page_numbering: Unsupported compatibility flag.

        Returns:
            Summary of the section update operation.
        """

        doc, source_path = load_document(path)
        if section_index < 0 or section_index >= len(doc.sections):
            raise IndexError(f"Section index out of range: {section_index}")
        section = doc.sections[section_index]
        unsupported = update_section_page_setup(
            section,
            orientation,
            paper_size,
            section_start,
            left_margin_points,
            right_margin_points,
            top_margin_points,
            bottom_margin_points,
            different_first_page_header_footer,
        )
        saved_to = save_document(doc, source_path, output_path)
        unsupported["restart_page_numbering"] = restart_page_numbering is not None
        return {
            "path": str(source_path),
            "saved_to": str(saved_to),
            "engine": "python-docx",
            "section": section_to_dict(section, section_index),
            "unsupported_options_ignored": unsupported,
        }



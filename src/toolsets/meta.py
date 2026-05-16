"""Metadata and capability-oriented MCP tools."""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, version
from typing import Any

from mcp.server.fastmcp import FastMCP

from config import DEFAULT_DIR, PACKAGE_NAME
from ops.document_ops import iter_paragraphs, iter_tables
from ops.package_io import load_document
from ops.review_ops import list_comments_xml, list_revisions_xml, settings_has_track_revisions
from toolsets.response_schema import tool_response


def package_version() -> str:
    """Return the installed package version when available.

    Returns:
        Distribution version string or ``dev`` fallback.
    """

    try:
        return version(PACKAGE_NAME)
    except PackageNotFoundError:
        return "dev"


def register_meta_tools(server: FastMCP) -> None:
    """Register metadata and discovery tools.

    Args:
        server: MCP server to populate.
    """

    @server.tool()
    @tool_response("server_info")
    def server_info() -> dict[str, Any]:
        """Return server capabilities and runtime metadata.

        Returns:
            Capability overview for IDE and client discovery.
        """

        return {
            "name": "extended-docx-mcp",
            "version": package_version(),
            "engine": "python-docx+lxml",
            "default_dir": str(DEFAULT_DIR),
            "supports": {
                "comments": True,
                "tracked_revisions": True,
                "table_structure_preservation": True,
                "uvx_launcher": True,
                "paragraph_search": True,
                "table_cell_inspection": True,
                "revision_context": True,
            },
            "tool_groups": [
                "meta",
                "content.discovery",
                "content.editing",
                "content.tables",
                "review.comments",
                "review.revisions",
                "styles",
                "sections",
            ],
        }

    @server.tool()
    @tool_response("inspect_document")
    def inspect_document(path: str) -> dict[str, Any]:
        """Return high-level document metadata and review counters.

        Args:
            path: Path to the DOCX file.

        Returns:
            Summary counts and track-revisions metadata.
        """

        doc, resolved = load_document(path)
        return {
            "path": str(resolved),
            "engine": "python-docx+lxml",
            "paragraphs": len(iter_paragraphs(doc)),
            "tables": len(iter_tables(doc)),
            "sections": len(doc.sections),
            "comments": len(list_comments_xml(resolved)),
            "revisions": len(list_revisions_xml(resolved)),
            "has_track_revisions_flag": settings_has_track_revisions(resolved),
        }


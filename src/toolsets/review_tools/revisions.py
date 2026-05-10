"""Tracked-revision tools exposed through FastMCP."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from ops.package_io import read_zip_xml, resolve_path, save_zip_parts, serialize_xml
from ops.review import (
    REVISION_NAMESPACES,
    REVISION_XPATH,
    list_revisions_xml,
    revision_details_xml,
    unwrap_revision,
)


def register_revision_tools(server: FastMCP) -> None:
    """Register tracked-revision tools on a FastMCP server.

    Args:
        server: MCP server to populate.
    """

    @server.tool()
    def list_revisions(path: str) -> dict[str, Any]:
        """Return current tracked insertions and deletions.

        Args:
            path: Path to the DOCX file.

        Returns:
            Serialized revision payload.
        """

        resolved = resolve_path(path)
        return {"path": str(resolved), "engine": "lxml-ooxml", "revisions": list_revisions_xml(resolved)}

    @server.tool()
    def get_revision_details(path: str, revision_index: int, context_paragraphs: int = 1) -> dict[str, Any]:
        """Return one revision together with paragraph-level context.

        Args:
            path: Path to the DOCX file.
            revision_index: Zero-based revision index in ``list_revisions`` order.
            context_paragraphs: Number of neighboring paragraphs to include on each side.

        Returns:
            Detailed revision payload with surrounding paragraph context.
        """

        resolved = resolve_path(path)
        details = revision_details_xml(resolved, revision_index, context_paragraphs)
        return {"path": str(resolved), "engine": "lxml-ooxml", "revision": details}

    @server.tool()
    def accept_all_revisions(path: str, output_path: str | None = None) -> dict[str, Any]:
        """Accept all tracked changes.

        Args:
            path: Path to the DOCX file.
            output_path: Optional alternate output path.

        Returns:
            Summary of the accept operation.
        """

        source_path = resolve_path(path)
        root = read_zip_xml(source_path, "word/document.xml")
        if root is None:
            raise ValueError("DOCX package is missing word/document.xml")
        before = len(list_revisions_xml(source_path))
        for revision in list(root.xpath(REVISION_XPATH, namespaces=REVISION_NAMESPACES)):
            parent = revision.getparent()
            if parent is None:
                continue
            if revision.tag.endswith("ins"):
                unwrap_revision(parent, revision)
            else:
                parent.remove(revision)
        saved_to = save_zip_parts(source_path, output_path, {"word/document.xml": serialize_xml(root)})
        return {"path": str(source_path), "saved_to": str(saved_to), "engine": "lxml-ooxml", "accepted": before, "remaining": len(list_revisions_xml(saved_to))}

    @server.tool()
    def reject_all_revisions(path: str, output_path: str | None = None) -> dict[str, Any]:
        """Reject all tracked changes.

        Args:
            path: Path to the DOCX file.
            output_path: Optional alternate output path.

        Returns:
            Summary of the reject operation.
        """

        source_path = resolve_path(path)
        root = read_zip_xml(source_path, "word/document.xml")
        if root is None:
            raise ValueError("DOCX package is missing word/document.xml")
        before = len(list_revisions_xml(source_path))
        for revision in list(root.xpath(REVISION_XPATH, namespaces=REVISION_NAMESPACES)):
            parent = revision.getparent()
            if parent is None:
                continue
            if revision.tag.endswith("ins"):
                parent.remove(revision)
            else:
                unwrap_revision(parent, revision, convert_deleted_text=True)
        saved_to = save_zip_parts(source_path, output_path, {"word/document.xml": serialize_xml(root)})
        return {"path": str(source_path), "saved_to": str(saved_to), "engine": "lxml-ooxml", "rejected": before, "remaining": len(list_revisions_xml(saved_to))}


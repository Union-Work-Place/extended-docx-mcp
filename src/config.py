"""Project-wide constants and enum mappings for the DOCX MCP server."""

from __future__ import annotations

import os
from pathlib import Path

from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE


SERVER_NAME = "extended-docx-mcp"
"""Human-readable MCP server name used during FastMCP registration."""

PACKAGE_NAME = "extended-docx-mcp-server"
"""Distribution name used for packaging metadata lookups."""

DEFAULT_DIR = Path(os.environ.get("EXTENDED_DOCX_MCP_DEFAULT_DIR", os.getcwd())).resolve()
"""Default working directory used to resolve relative DOCX paths."""

BACKUP_DIRNAME = ".extended-docx-mcp-backups"
"""Folder name used for automatic backups when saving over an existing DOCX file."""

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
"""WordprocessingML namespace URI."""

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
"""Office document relationships namespace URI."""

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
"""Package relationships namespace URI."""

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
"""Package content types namespace URI."""

XML_NS = "http://www.w3.org/XML/1998/namespace"
"""Core XML namespace used for attributes like xml:space."""

COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
"""Content type registered for ``word/comments.xml`` in the DOCX package."""

COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
"""Relationship type used to attach comments to ``word/document.xml``."""

W = f"{{{W_NS}}}"
"""Fully-qualified WordprocessingML namespace prefix for tag construction."""

R = f"{{{R_NS}}}"
"""Fully-qualified office document relationships namespace prefix for tag construction."""

REL = f"{{{REL_NS}}}"
"""Fully-qualified package relationships namespace prefix for tag construction."""

CT = f"{{{CT_NS}}}"
"""Fully-qualified content types namespace prefix for tag construction."""

NSMAP = {"w": W_NS, "r": R_NS}
"""Namespace mapping shared by XPath queries over OOXML parts."""

PX_ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}
"""Supported paragraph alignment values for MCP tools."""

PX_UNDERLINES = {
    "none": False,
    "single": True,
    "double": WD_UNDERLINE.DOUBLE,
}
"""Supported underline values for MCP tools."""

PX_TABLE_ALIGNMENTS = {
    "left": WD_TABLE_ALIGNMENT.LEFT,
    "center": WD_TABLE_ALIGNMENT.CENTER,
    "right": WD_TABLE_ALIGNMENT.RIGHT,
}
"""Supported table alignment values for MCP tools."""

PX_SECTION_STARTS = {
    "continuous": WD_SECTION_START.CONTINUOUS,
    "new_column": WD_SECTION_START.NEW_COLUMN,
    "new_page": WD_SECTION_START.NEW_PAGE,
    "even_page": WD_SECTION_START.EVEN_PAGE,
    "odd_page": WD_SECTION_START.ODD_PAGE,
}
"""Supported section break modes for MCP tools."""

PX_STYLE_TYPES = {
    "paragraph": WD_STYLE_TYPE.PARAGRAPH,
    "character": WD_STYLE_TYPE.CHARACTER,
    "table": WD_STYLE_TYPE.TABLE,
    "list": WD_STYLE_TYPE.LIST,
}
"""Supported Word style categories exposed by the MCP API."""

PX_ORIENTATIONS = {
    "portrait": WD_ORIENT.PORTRAIT,
    "landscape": WD_ORIENT.LANDSCAPE,
}
"""Supported page orientations exposed by the MCP API."""

PAPER_SIZES_INCHES = {
    "a4": (8.27, 11.69),
    "a3": (11.69, 16.54),
    "letter": (8.5, 11),
    "legal": (8.5, 14),
}
"""Known paper sizes mapped to width/height in inches."""

STRUCTURED_BLOCK_TYPES = {
    "paragraph",
    "heading",
    "table",
    "page_break",
    "section_break",
}
"""Supported block types accepted by ``write_docx``."""


"""FastMCP tool registration helpers grouped by responsibility."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from toolsets.content import register_content_tools
from toolsets.meta import register_meta_tools
from toolsets.review import register_review_tools
from toolsets.sections import register_section_tools
from toolsets.styles import register_style_tools


def register_all_tools(server: FastMCP) -> None:
    """Register every toolset on a FastMCP server instance.

    Args:
        server: MCP server instance to populate with tools.
    """

    register_meta_tools(server)
    register_content_tools(server)
    register_review_tools(server)
    register_style_tools(server)
    register_section_tools(server)



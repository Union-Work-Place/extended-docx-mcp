"""Aggregated registration entrypoint for content-related MCP tools."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from toolsets.content_tools.batch import register_batch_tools
from toolsets.content_tools.discovery import register_content_discovery_tools
from toolsets.content_tools.editing import register_content_editing_tools
from toolsets.content_tools.ranges import register_range_tools
from toolsets.content_tools.tables import register_table_tools


def register_content_tools(server: FastMCP) -> None:
    """Register all content-related tools on a FastMCP server.

    Args:
        server: MCP server to populate.
    """

    register_content_discovery_tools(server)
    register_content_editing_tools(server)
    register_range_tools(server)
    register_table_tools(server)
    register_batch_tools(server)

"""Aggregated registration entrypoint for review-related MCP tools."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from toolsets.review_tools.comments import register_comment_tools
from toolsets.review_tools.revisions import register_revision_tools


def register_review_tools(server: FastMCP) -> None:
    """Register all review-related tools on a FastMCP server.

    Args:
        server: MCP server to populate.
    """

    register_comment_tools(server)
    register_revision_tools(server)


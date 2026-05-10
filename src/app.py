"""FastMCP application factory and process entrypoints."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from config import SERVER_NAME
from toolsets import register_all_tools


def create_server() -> FastMCP:
    """Create and configure the DOCX MCP server instance.

    Returns:
        Fully-registered ``FastMCP`` server instance.
    """

    server = FastMCP(SERVER_NAME)
    register_all_tools(server)
    return server


SERVER = create_server()
"""Module-level MCP server instance used by CLI entrypoints."""


def main() -> None:
    """Run the DOCX MCP server over stdio.

    Returns:
        ``None``. The function blocks while the stdio server is running.
    """

    SERVER.run(transport="stdio")



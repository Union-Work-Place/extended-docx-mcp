"""Compatibility entrypoint for the DOCX MCP server."""

from __future__ import annotations

from app import SERVER, create_server, main

__all__ = ["SERVER", "create_server", "main"]


if __name__ == "__main__":
	main()


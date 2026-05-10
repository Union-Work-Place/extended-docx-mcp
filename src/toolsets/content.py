"""Compatibility wrapper for content tool registration.

The actual registration now lives in ``toolsets.content_tools``.
"""

from toolsets.content_tools import register_content_tools

__all__ = ["register_content_tools"]



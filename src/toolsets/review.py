"""Compatibility wrapper for review tool registration.

The actual registration now lives in ``toolsets.review_tools``.
"""

from toolsets.review_tools import register_review_tools

__all__ = ["register_review_tools"]



"""Text search and replacement helpers for DOCX content operations."""

from __future__ import annotations

import re
from typing import Any

from docx.text.paragraph import Paragraph


def normalize_mapping_value(value: str, mapping: dict[str, Any], label: str) -> Any:
    """Resolve a string value through one of the supported enum maps.

    Args:
        value: User-provided string value.
        mapping: Supported string-to-enum mapping.
        label: Human-readable label used in validation errors.

    Returns:
        Matching mapped enum value.
    """

    normalized = value.lower()
    if normalized not in mapping:
        raise ValueError(f"Unsupported {label}: {value}")
    return mapping[normalized]


def enum_name(value: Any, mapping: dict[str, Any]) -> str | None:
    """Convert an enum-like value back to its public MCP string representation.

    Args:
        value: Enum value or primitive to convert.
        mapping: Supported string-to-enum mapping.

    Returns:
        Matching public string or ``str(value)`` fallback.
    """

    if value is None:
        return None
    for key, candidate in mapping.items():
        if value == candidate:
            return key
    return str(value)


def matches_for_text(text: str, find_text: str, match_case: bool, find_whole_words_only: bool) -> list[re.Match[str]]:
    """Find text matches using the same rules as replace/comment tools.

    Args:
        text: Source text to inspect.
        find_text: Search term.
        match_case: Whether matching should be case-sensitive.
        find_whole_words_only: Whether only whole words should match.

    Returns:
        Regex match objects in source order.
    """

    if not find_text:
        raise ValueError("find_text must not be empty")
    flags = 0 if match_case else re.IGNORECASE
    escaped = re.escape(find_text)
    pattern = rf"(?<!\w){escaped}(?!\w)" if find_whole_words_only else escaped
    return list(re.finditer(pattern, text, flags))


def replace_matches_plain(paragraph: Paragraph, matches: list[re.Match[str]], replace_with: str) -> int:
    """Replace already-found matches inside paragraph runs.

    Args:
        paragraph: Paragraph to mutate.
        matches: Precomputed matches for the paragraph text.
        replace_with: Replacement text.

    Returns:
        Number of replacements applied.
    """

    if not matches:
        return 0
    for match in reversed(matches):
        runs = list(paragraph.runs)
        if not runs:
            paragraph.add_run(replace_with)
            continue
        run_ranges: list[tuple[int, int, int]] = []
        offset = 0
        for run_index, run in enumerate(runs):
            length = len(run.text)
            run_ranges.append((run_index, offset, offset + length))
            offset += length
        start, end = match.span()
        touched = [
            (run_index, run_start, run_end)
            for run_index, run_start, run_end in run_ranges
            if run_end > start and run_start < end
        ]
        if not touched:
            continue
        first_index, first_start, _ = touched[0]
        _, last_start, _ = touched[-1]
        first_run = runs[first_index]
        last_run = runs[touched[-1][0]]
        first_run.text = f"{first_run.text[: start - first_start]}{replace_with}{last_run.text[end - last_start :]}"
        for run_index, _, _ in touched[1:]:
            runs[run_index].text = ""
    return len(matches)


def replace_in_paragraph_plain(
    paragraph: Paragraph,
    find_text: str,
    replace_with: str,
    match_case: bool,
    find_whole_words_only: bool,
) -> int:
    """Replace matches in a paragraph while preserving surrounding structure.

    Args:
        paragraph: Paragraph to mutate.
        find_text: Search term.
        replace_with: Replacement text.
        match_case: Whether matching should be case-sensitive.
        find_whole_words_only: Whether only whole words should match.

    Returns:
        Number of replacements applied inside the paragraph.
    """

    matches = matches_for_text(paragraph.text, find_text, match_case, find_whole_words_only)
    return replace_matches_plain(paragraph, matches, replace_with)
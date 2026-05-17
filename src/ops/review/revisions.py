"""Tracked-revision helpers implemented directly against WordprocessingML."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from lxml import etree

from config import W
from ops.package_io import load_document, read_zip_xml, serialize_xml
from ops.review.xml_utils import (
    REVISION_NAMESPACES,
    REVISION_XPATH,
    canonical_paragraph_refs,
    clear_paragraph_content,
    new_text_run,
    new_w_element,
    paragraph_revision_view,
    w_attr,
)
from ops.text_ops import matches_for_text


@dataclass(frozen=True)
class TextRunSegment:
    """A text range and its source run properties."""

    start: int
    end: int
    text: str
    run_properties: etree._Element | None


def run_properties_for_text_node(text_node: etree._Element) -> etree._Element | None:
    """Find the nearest source run properties for a text node."""

    parent = text_node.getparent()
    while parent is not None and parent.tag != f"{W}r":
        parent = parent.getparent()
    if parent is None:
        return None
    return parent.find(f"{W}rPr")


def paragraph_text_segments(paragraph: etree._Element) -> tuple[str, list[TextRunSegment]]:
    """Return paragraph text with source run-property ranges."""

    text_parts: list[str] = []
    segments: list[TextRunSegment] = []
    cursor = 0
    for text_node in paragraph.findall(f".//{W}t"):
        value = text_node.text or ""
        if not value:
            continue
        start = cursor
        cursor += len(value)
        text_parts.append(value)
        segments.append(TextRunSegment(start, cursor, value, run_properties_for_text_node(text_node)))
    return "".join(text_parts), segments


def append_segment_range(
    parent: etree._Element,
    segments: list[TextRunSegment],
    start: int,
    end: int,
    *,
    deleted: bool = False,
) -> None:
    """Append text runs for a source-text range, preserving run properties."""

    for segment in segments:
        if segment.end <= start:
            continue
        if segment.start >= end:
            break
        local_start = max(start, segment.start) - segment.start
        local_end = min(end, segment.end) - segment.start
        parent.append(new_text_run(segment.text[local_start:local_end], deleted=deleted, run_properties=segment.run_properties))


def run_properties_at(segments: list[TextRunSegment], position: int) -> etree._Element | None:
    """Return run properties at a source-text position."""

    for segment in segments:
        if segment.start <= position < segment.end:
            return segment.run_properties
    return segments[-1].run_properties if segments else None


def append_revision_pair(
    paragraph: etree._Element,
    old_text: str,
    new_text: str,
    revision_id: int,
    author: str,
    timestamp: str,
    run_properties: etree._Element | None = None,
) -> None:
    """Append a matched delete/insert revision pair to a paragraph.

    Args:
        paragraph: Paragraph XML element to mutate.
        old_text: Deleted source text.
        new_text: Inserted replacement text.
        revision_id: Base numeric revision id.
        author: Revision author label.
        timestamp: ISO timestamp string.
        run_properties: Optional run properties to apply to generated runs.
    """

    deletion = new_w_element(
        "del",
        {
            w_attr("id"): revision_id,
            w_attr("author"): author,
            w_attr("date"): timestamp,
        },
    )
    deletion.append(new_text_run(old_text, deleted=True, run_properties=run_properties))
    insertion = new_w_element(
        "ins",
        {
            w_attr("id"): revision_id + 1,
            w_attr("author"): author,
            w_attr("date"): timestamp,
        },
    )
    insertion.append(new_text_run(new_text, run_properties=run_properties))
    paragraph.append(deletion)
    paragraph.append(insertion)


def append_revision_pair_from_segments(
    paragraph: etree._Element,
    segments: list[TextRunSegment],
    match_start: int,
    match_end: int,
    new_text: str,
    revision_id: int,
    author: str,
    timestamp: str,
) -> None:
    """Append a tracked replacement while preserving source run properties."""

    deletion = new_w_element(
        "del",
        {
            w_attr("id"): revision_id,
            w_attr("author"): author,
            w_attr("date"): timestamp,
        },
    )
    append_segment_range(deletion, segments, match_start, match_end, deleted=True)
    insertion = new_w_element(
        "ins",
        {
            w_attr("id"): revision_id + 1,
            w_attr("author"): author,
            w_attr("date"): timestamp,
        },
    )
    insertion.append(new_text_run(new_text, run_properties=run_properties_at(segments, match_start)))
    paragraph.append(deletion)
    paragraph.append(insertion)


def next_revision_id(document_root: etree._Element) -> int:
    """Compute the next available tracked-revision id in a document.

    Args:
        document_root: Parsed ``word/document.xml`` root.

    Returns:
        Next numeric revision id.
    """

    ids = []
    for element in document_root.xpath(REVISION_XPATH, namespaces=REVISION_NAMESPACES):
        value = element.get(w_attr("id"))
        if value and value.isdigit():
            ids.append(int(value))
    return max(ids, default=0) + 1


def revision_text(element: etree._Element) -> str:
    """Extract visible text payload from a revision node.

    Args:
        element: ``w:ins`` or ``w:del`` node.

    Returns:
        Revision text payload.
    """

    if element.tag == f"{W}del":
        return "".join(node.text or "" for node in element.findall(f".//{W}delText"))
    return "".join(node.text or "" for node in element.findall(f".//{W}t"))


def revision_to_dict(element: etree._Element, index: int) -> dict[str, Any]:
    """Serialize a revision node into MCP output.

    Args:
        element: ``w:ins`` or ``w:del`` node.
        index: Revision index in document order.

    Returns:
        Serialized revision descriptor.
    """

    return {
        "index": index,
        "id": element.get(w_attr("id")),
        "author": element.get(w_attr("author")),
        "date_time": element.get(w_attr("date")),
        "revision_type": "deletion" if element.tag == f"{W}del" else "insertion",
        "text": revision_text(element),
    }


def list_revisions_xml(path: Any) -> list[dict[str, Any]]:
    """Load and serialize all tracked revisions from a DOCX file.

    Args:
        path: DOCX path to inspect.

    Returns:
        Serialized revision descriptors.
    """

    root = read_zip_xml(path, "word/document.xml")
    if root is None:
        return []
    elements = root.xpath(REVISION_XPATH, namespaces=REVISION_NAMESPACES)
    return [revision_to_dict(element, index) for index, element in enumerate(elements)]


def settings_has_track_revisions(path: Any) -> bool:
    """Check whether ``word/settings.xml`` contains the track revisions flag.

    Args:
        path: DOCX path to inspect.

    Returns:
        ``True`` when ``w:trackRevisions`` is present.
    """

    settings = read_zip_xml(path, "word/settings.xml")
    return bool(settings is not None and settings.find(f"{W}trackRevisions") is not None)


def enable_track_revisions_part(path: Any) -> bytes | None:
    """Ensure ``word/settings.xml`` includes ``w:trackRevisions``.

    Args:
        path: DOCX path to mutate.

    Returns:
        Serialized updated settings part or ``None`` when the part is missing.
    """

    settings = read_zip_xml(path, "word/settings.xml")
    if settings is None:
        return None
    if settings.find(f"{W}trackRevisions") is None:
        settings.append(new_w_element("trackRevisions"))
    return serialize_xml(settings)


def replace_in_paragraph_tracked(
    paragraph: etree._Element,
    find_text: str,
    replace_with: str,
    match_case: bool,
    find_whole_words_only: bool,
    revision_id: int,
    author: str,
    timestamp: str,
) -> tuple[int, int]:
    """Replace paragraph text with OOXML tracked revisions.

    Args:
        paragraph: Paragraph XML element to mutate.
        find_text: Search term.
        replace_with: Replacement text.
        match_case: Whether matching should be case-sensitive.
        find_whole_words_only: Whether only whole words should match.
        revision_id: Starting revision id.
        author: Revision author label.
        timestamp: ISO timestamp string.

    Returns:
        Tuple of replacement count and next available revision id.
    """

    text, segments = paragraph_text_segments(paragraph)
    matches = matches_for_text(text, find_text, match_case, find_whole_words_only)
    if not matches:
        return 0, revision_id
    clear_paragraph_content(paragraph)
    cursor = 0
    for match in matches:
        if match.start() > cursor:
            append_segment_range(paragraph, segments, cursor, match.start())
        append_revision_pair_from_segments(paragraph, segments, match.start(), match.end(), replace_with, revision_id, author, timestamp)
        revision_id += 2
        cursor = match.end()
    if cursor < len(text):
        append_segment_range(paragraph, segments, cursor, len(text))
    return len(matches), revision_id


def convert_del_text_to_text(element: etree._Element) -> None:
    """Convert ``w:delText`` nodes into normal ``w:t`` nodes in-place.

    Args:
        element: Revision subtree to normalize.
    """

    for deleted_text in element.findall(f".//{W}delText"):
        deleted_text.tag = f"{W}t"


def unwrap_revision(parent: etree._Element, revision: etree._Element, convert_deleted_text: bool = False) -> None:
    """Replace a revision wrapper with its children.

    Args:
        parent: Parent node containing the revision.
        revision: ``w:ins`` or ``w:del`` node to unwrap.
        convert_deleted_text: Whether deleted text should be turned into normal text nodes first.
    """

    if convert_deleted_text:
        convert_del_text_to_text(revision)
    index = parent.index(revision)
    for child in list(revision):
        parent.insert(index, child)
        index += 1
    parent.remove(revision)


def revision_details_xml(path: Any, revision_index: int, context_paragraphs: int = 1) -> dict[str, Any]:
    """Return one revision together with paragraph-level context.

    Args:
        path: DOCX path to inspect.
        revision_index: Zero-based revision index in document order.
        context_paragraphs: Number of neighboring paragraphs to include on each side.

    Returns:
        Serialized revision details with paragraph context.
    """

    if revision_index < 0:
        raise ValueError("revision_index must be >= 0")
    if context_paragraphs < 0:
        raise ValueError("context_paragraphs must be >= 0")
    root = read_zip_xml(path, "word/document.xml")
    if root is None:
        raise ValueError("DOCX package is missing word/document.xml")
    doc, _ = load_document(str(path))
    revisions = root.xpath(REVISION_XPATH, namespaces=REVISION_NAMESPACES)
    if revision_index >= len(revisions):
        raise IndexError(f"Revision index out of range: {revision_index}")
    revision = revisions[revision_index]
    paragraph = revision
    while paragraph is not None and paragraph.tag != f"{W}p":
        paragraph = paragraph.getparent()
    paragraph_refs = canonical_paragraph_refs(doc, root)
    paragraph_by_id = {id(ref.paragraph): ref for ref in paragraph_refs}
    paragraph_ref = paragraph_by_id.get(id(paragraph)) if paragraph is not None else None
    paragraph_index = paragraph_ref.index if paragraph_ref is not None else None
    details = revision_to_dict(revision, revision_index)
    if paragraph is None or paragraph_index is None:
        details["paragraph"] = None
        details["context"] = []
        return details
    context_start = max(0, paragraph_index - context_paragraphs)
    context_end = min(len(paragraph_refs), paragraph_index + context_paragraphs + 1)
    details["paragraph"] = {
        "index": paragraph_index,
        "visible_text": paragraph_revision_view(paragraph)["visible_text"],
        "annotated_text": paragraph_revision_view(paragraph)["annotated_text"],
    }
    details["context"] = [
        {
            "index": ref.index,
            "visible_text": paragraph_revision_view(ref.paragraph)["visible_text"],
            "annotated_text": paragraph_revision_view(ref.paragraph)["annotated_text"],
            "contains_revision": ref.index == paragraph_index,
        }
        for ref in paragraph_refs[context_start:context_end]
    ]
    return details

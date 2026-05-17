"""Word comment helpers implemented directly against OOXML package parts."""

from __future__ import annotations

import datetime as dt
from typing import Any

from lxml import etree

from config import COMMENTS_CONTENT_TYPE, COMMENTS_REL_TYPE, CT, REL, W
from ops.package_io import load_document, read_zip_xml
from ops.review.xml_utils import (
    REVISION_NAMESPACES,
    REVISION_XPATH,
    canonical_paragraph_refs,
    clear_paragraph_content,
    new_text_run,
    new_w_element,
    paragraph_text_xml,
    w_attr,
)


def ensure_comments_parts(path: Any) -> tuple[etree._Element, etree._Element, etree._Element]:
    """Load or create the XML parts required for Word comments.

    Args:
        path: DOCX path whose package should be prepared.

    Returns:
        Content types root, relationships root and comments root.
    """

    content_types = read_zip_xml(path, "[Content_Types].xml")
    rels = read_zip_xml(path, "word/_rels/document.xml.rels")
    comments = read_zip_xml(path, "word/comments.xml")
    if content_types is None or rels is None:
        raise ValueError("DOCX package is missing required content type or relationship parts")
    if comments is None:
        comments = etree.Element(f"{W}comments", nsmap={"w": W[1:-1]})

    has_override = any(element.get("PartName") == "/word/comments.xml" for element in content_types.findall(f"{CT}Override"))
    if not has_override:
        override = etree.Element(f"{CT}Override")
        override.set("PartName", "/word/comments.xml")
        override.set("ContentType", COMMENTS_CONTENT_TYPE)
        content_types.append(override)

    has_rel = any(
        element.get("Type") == COMMENTS_REL_TYPE and element.get("Target") == "comments.xml"
        for element in rels.findall(f"{REL}Relationship")
    )
    if not has_rel:
        existing_ids = [element.get("Id", "") for element in rels.findall(f"{REL}Relationship")]
        next_id = 1
        while f"rId{next_id}" in existing_ids:
            next_id += 1
        rel = etree.Element(f"{REL}Relationship")
        rel.set("Id", f"rId{next_id}")
        rel.set("Type", COMMENTS_REL_TYPE)
        rel.set("Target", "comments.xml")
        rels.append(rel)
    return content_types, rels, comments


def next_comment_id(comments_root: etree._Element) -> int:
    """Compute the next available Word comment id.

    Args:
        comments_root: Parsed ``word/comments.xml`` root.

    Returns:
        Next numeric comment id.
    """

    ids = [
        int(comment.get(w_attr("id")))
        for comment in comments_root.findall(f"{W}comment")
        if (comment.get(w_attr("id")) or "").isdigit()
    ]
    return max(ids, default=-1) + 1


def add_comment_node(
    comments_root: etree._Element,
    comment_id: int,
    comment_text: str,
    author: str,
    initials: str,
    parent_id: int | None = None,
) -> etree._Element:
    """Append a comment node to ``word/comments.xml``.

    Args:
        comments_root: Comments XML root.
        comment_id: New comment id.
        comment_text: Comment text body.
        author: Comment author name.
        initials: Author initials.
        parent_id: Optional parent comment id for replies.

    Returns:
        Created comment XML element.
    """

    comment = new_w_element(
        "comment",
        {
            w_attr("id"): comment_id,
            w_attr("author"): author,
            w_attr("initials"): initials,
            w_attr("date"): dt.datetime.now(dt.timezone.utc).isoformat(),
        },
    )
    if parent_id is not None:
        comment.set(w_attr("parentId"), str(parent_id))
    paragraph = new_w_element("p")
    paragraph.append(new_text_run(comment_text))
    comment.append(paragraph)
    comments_root.append(comment)
    return comment


def comment_reference_run(comment_id: int) -> etree._Element:
    """Build the reference run used after a comment range.

    Args:
        comment_id: Linked Word comment id.

    Returns:
        Reference run XML element.
    """

    run = new_w_element("r")
    reference = new_w_element("commentReference", {w_attr("id"): comment_id})
    run.append(reference)
    return run


def anchor_comment_to_paragraph(paragraph: etree._Element, comment_id: int) -> None:
    """Anchor a comment to the entire paragraph range.

    Args:
        paragraph: Target paragraph XML element.
        comment_id: Linked Word comment id.
    """

    start = new_w_element("commentRangeStart", {w_attr("id"): comment_id})
    end = new_w_element("commentRangeEnd", {w_attr("id"): comment_id})
    children = list(paragraph)
    insert_at = 1 if children and children[0].tag == f"{W}pPr" else 0
    paragraph.insert(insert_at, start)
    paragraph.append(end)
    paragraph.append(comment_reference_run(comment_id))


def anchor_comment_to_range(paragraph: etree._Element, comment_id: int, start_offset: int, end_offset: int) -> None:
    """Anchor a comment to a character range inside a paragraph.

    Args:
        paragraph: Target paragraph XML element.
        comment_id: Linked Word comment id.
        start_offset: Inclusive character start offset.
        end_offset: Exclusive character end offset.
    """

    text = paragraph_text_xml(paragraph)
    if start_offset < 0 or end_offset <= start_offset or end_offset > len(text):
        raise ValueError(f"Text range {start_offset}:{end_offset} is outside paragraph length {len(text)}")
    if paragraph.xpath(REVISION_XPATH, namespaces=REVISION_NAMESPACES):
        anchor_comment_to_paragraph(paragraph, comment_id)
        return
    clear_paragraph_content(paragraph)
    if start_offset:
        paragraph.append(new_text_run(text[:start_offset]))
    paragraph.append(new_w_element("commentRangeStart", {w_attr("id"): comment_id}))
    paragraph.append(new_text_run(text[start_offset:end_offset]))
    paragraph.append(new_w_element("commentRangeEnd", {w_attr("id"): comment_id}))
    paragraph.append(comment_reference_run(comment_id))
    if end_offset < len(text):
        paragraph.append(new_text_run(text[end_offset:]))


def comment_to_dict(comment: etree._Element, index: int) -> dict[str, Any]:
    """Serialize a comment node into MCP output.

    Args:
        comment: Comment XML element.
        index: Comment index in comments.xml order.

    Returns:
        Serialized comment descriptor.
    """

    text = "".join(node.text or "" for node in comment.findall(f".//{W}t"))
    parent_id = comment.get(w_attr("parentId"))
    return {
        "index": index,
        "id": int(comment.get(w_attr("id"), "0")),
        "author": comment.get(w_attr("author")),
        "initial": comment.get(w_attr("initials")),
        "date_time": comment.get(w_attr("date")),
        "done": False,
        "parent_id": int(parent_id) if parent_id and parent_id.isdigit() else None,
        "text": text,
    }


def list_comments_xml(path: Any) -> list[dict[str, Any]]:
    """Load and serialize all comments from a DOCX file.

    Args:
        path: DOCX path to inspect.

    Returns:
        Serialized comment descriptors.
    """

    comments_root = read_zip_xml(path, "word/comments.xml")
    if comments_root is None:
        return []
    comments = comments_root.findall(f"{W}comment")
    return [comment_to_dict(comment, index) for index, comment in enumerate(comments)]


def find_text_range_xml(
    doc: Any,
    root: etree._Element,
    target_text: str,
    occurrence_index: int = 0,
    paragraph_index: int | None = None,
    anchor_text: str | None = None,
) -> tuple[etree._Element, int, int, int]:
    """Find the first matching text range within ``word/document.xml``.

    Args:
        doc: Loaded ``python-docx`` document for canonical paragraph ordering.
        root: Parsed ``word/document.xml`` root.
        target_text: Text to locate.
        occurrence_index: Zero-based occurrence to return.
        paragraph_index: Optional paragraph restriction.
        anchor_text: Optional paragraph search anchor.

    Returns:
        Paragraph node, paragraph index, start offset and end offset.
    """

    if occurrence_index < 0:
        raise ValueError("occurrence_index must be >= 0")
    paragraphs = canonical_paragraph_refs(doc, root)
    candidates: list[tuple[int, etree._Element]]
    if paragraph_index is not None:
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            raise IndexError(f"Paragraph index out of range: {paragraph_index}")
        candidates = [(paragraph_index, paragraphs[paragraph_index].paragraph)]
    elif anchor_text:
        candidates = [(ref.index, ref.paragraph) for ref in paragraphs if anchor_text in paragraph_text_xml(ref.paragraph)]
        if not candidates:
            raise ValueError(f"Anchor text was not found in document: {anchor_text}")
    else:
        candidates = [(ref.index, ref.paragraph) for ref in paragraphs]

    seen = 0
    for actual_index, paragraph in candidates:
        text = paragraph_text_xml(paragraph)
        start = 0
        while True:
            offset = text.find(target_text, start)
            if offset < 0:
                break
            if seen == occurrence_index:
                return paragraph, actual_index, offset, offset + len(target_text)
            seen += 1
            start = offset + max(1, len(target_text))
    raise ValueError(f"Target text occurrence was not found: {target_text}")

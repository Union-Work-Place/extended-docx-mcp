"""Low-level OOXML utilities shared by review-oriented operations."""

from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import Any

from lxml import etree

from config import NSMAP, W, XML_NS
from ops.structure_ops import iter_paragraphs


@dataclass(frozen=True)
class CanonicalParagraphRef:
    """Canonical paragraph descriptor shared across python-docx and OOXML paths."""

    index: int
    xml_path: str
    paragraph: etree._Element


def w_attr(name: str) -> str:
    """Build a fully-qualified ``w:`` attribute name.

    Args:
        name: Unqualified attribute name.

    Returns:
        Fully-qualified OOXML attribute name.
    """

    return f"{W}{name}"


def xml_space_attr() -> str:
    """Return the fully-qualified ``xml:space`` attribute name.

    Returns:
        Fully-qualified XML namespace attribute name.
    """

    return f"{{{XML_NS}}}space"


def new_w_element(tag: str, attrs: dict[str, Any] | None = None) -> etree._Element:
    """Create a new WordprocessingML element.

    Args:
        tag: Local WordprocessingML tag name.
        attrs: Optional attribute mapping.

    Returns:
        Newly created XML element.
    """

    element = etree.Element(f"{W}{tag}")
    for key, value in (attrs or {}).items():
        element.set(key, str(value))
    return element


def new_text_run(text: str, deleted: bool = False, run_properties: etree._Element | None = None) -> etree._Element:
    """Create a run containing either normal or deleted text.

    Args:
        text: Visible run text.
        deleted: Whether the run should store deleted text markup.
        run_properties: Optional ``w:rPr`` node to clone into the run.

    Returns:
        Run XML element containing the requested text node.
    """

    run = new_w_element("r")
    if run_properties is not None:
        run.append(copy.deepcopy(run_properties))
    text_element = new_w_element("delText" if deleted else "t")
    if text[:1].isspace() or text[-1:].isspace():
        text_element.set(xml_space_attr(), "preserve")
    text_element.text = text
    run.append(text_element)
    return run


def paragraph_text_xml(paragraph: etree._Element, include_deleted: bool = False) -> str:
    """Read concatenated text from a paragraph XML node.

    Args:
        paragraph: Paragraph XML element.
        include_deleted: Whether to include ``w:delText`` nodes.

    Returns:
        Concatenated plain text.
    """

    tags = [f".//{W}t"]
    if include_deleted:
        tags.append(f".//{W}delText")
    values: list[str] = []
    for tag in tags:
        values.extend(node.text or "" for node in paragraph.findall(tag))
    return "".join(values)


def child_plain_text(element: etree._Element, deleted: bool = False) -> str:
    """Extract visible plain text from a child OOXML element.

    Args:
        element: XML element to inspect.
        deleted: Whether to read deleted text nodes instead of normal text nodes.

    Returns:
        Concatenated plain text.
    """

    tag = "delText" if deleted else "t"
    return "".join(node.text or "" for node in element.findall(f".//{W}{tag}"))


def paragraph_revision_view(paragraph: etree._Element) -> dict[str, str]:
    """Build visible and annotated text views for a paragraph with revisions.

    Args:
        paragraph: Paragraph XML element.

    Returns:
        Mapping with visible, annotated, inserted and deleted text variants.
    """

    visible_parts: list[str] = []
    annotated_parts: list[str] = []
    inserted_parts: list[str] = []
    deleted_parts: list[str] = []

    for child in paragraph:
        if child.tag == f"{W}pPr":
            continue
        if child.tag == f"{W}ins":
            text = child_plain_text(child, deleted=False)
            if text:
                visible_parts.append(text)
                annotated_parts.append(f"{{+{text}+}}")
                inserted_parts.append(text)
            continue
        if child.tag == f"{W}del":
            text = child_plain_text(child, deleted=True)
            if text:
                annotated_parts.append(f"[-{text}-]")
                deleted_parts.append(text)
            continue
        text = child_plain_text(child, deleted=False) if child.tag == f"{W}r" else "".join(node.text or "" for node in child.findall(f".//{W}t"))
        if text:
            visible_parts.append(text)
            annotated_parts.append(text)

    visible_text = "".join(visible_parts)
    return {
        "visible_text": visible_text,
        "annotated_text": "".join(annotated_parts) if annotated_parts else visible_text,
        "inserted_text": "".join(inserted_parts),
        "deleted_text": "".join(deleted_parts),
    }


def iter_document_paragraphs_xml(document_root: etree._Element) -> list[etree._Element]:
    """Return all paragraph nodes from ``word/document.xml``.

    Args:
        document_root: Parsed ``word/document.xml`` root.

    Returns:
        Paragraph XML elements in document order.
    """

    body = document_root.find(f".//{W}body")
    if body is None:
        return []
    return list(body.iter(f"{W}p"))


def canonical_paragraph_refs(doc: Any, document_root: etree._Element | None) -> list[CanonicalParagraphRef]:
    """Return paragraph XML nodes in the canonical ``iter_paragraphs(doc)`` order."""

    if document_root is None:
        return []
    tree = document_root.getroottree()
    xml_by_path = {tree.getpath(paragraph): paragraph for paragraph in iter_document_paragraphs_xml(document_root)}
    refs: list[CanonicalParagraphRef] = []
    for index, paragraph in enumerate(iter_paragraphs(doc)):
        xml_path = paragraph._p.getroottree().getpath(paragraph._p)
        xml_paragraph = xml_by_path.get(xml_path)
        if xml_paragraph is None:
            continue
        refs.append(CanonicalParagraphRef(index=index, xml_path=xml_path, paragraph=xml_paragraph))
    return refs


def canonical_paragraph_xml(doc: Any, document_root: etree._Element | None) -> list[etree._Element]:
    """Return paragraph XML nodes in the canonical ``iter_paragraphs(doc)`` order."""

    return [ref.paragraph for ref in canonical_paragraph_refs(doc, document_root)]


def iter_document_tables_xml(document_root: etree._Element) -> list[etree._Element]:
    """Return all table nodes from ``word/document.xml``.

    Args:
        document_root: Parsed ``word/document.xml`` root.

    Returns:
        Table XML elements in document order.
    """

    body = document_root.find(f".//{W}body")
    if body is None:
        return []
    return list(body.iter(f"{W}tbl"))


def clear_paragraph_content(paragraph: etree._Element) -> None:
    """Remove all visible paragraph children except paragraph properties.

    Args:
        paragraph: Paragraph XML element to clear.
    """

    for child in list(paragraph):
        if child.tag != f"{W}pPr":
            paragraph.remove(child)


REVISION_XPATH = ".//w:ins | .//w:del"
"""XPath expression shared by revision readers and mutators."""

REVISION_NAMESPACES = NSMAP
"""Namespace mapping used by revision XPath helpers."""

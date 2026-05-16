"""Utilities for generating reproducible DOCX test fixtures."""

from __future__ import annotations

import asyncio
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.shared import Pt


def _save(doc: Document, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    return path


def build_simple(path: Path) -> Path:
    doc = Document()
    doc.add_heading("Quarterly Report", level=1)
    doc.add_paragraph("Alpha project budget is 100.")
    doc.add_paragraph("The draft conclusion needs review.")
    doc.add_paragraph("Final summary for regional team.")
    table = doc.add_table(rows=2, cols=3)
    table.style = "Table Grid"
    table.cell(0, 0).text = "Metric"
    table.cell(0, 1).text = "Q1"
    table.cell(0, 2).text = "Q2"
    table.cell(1, 0).text = "Revenue"
    table.cell(1, 1).text = "120"
    table.cell(1, 2).text = "145"
    return _save(doc, path)


def build_with_styles(path: Path) -> Path:
    doc = Document()
    style = doc.styles.add_style("Callout", 1)
    style.font.name = "Arial"
    style.font.size = Pt(14)
    style.font.bold = True
    doc.add_heading("Styled document", level=1)
    paragraph = doc.add_paragraph("Styled callout paragraph.")
    paragraph.style = "Callout"
    doc.add_paragraph("Reference body paragraph.")
    return _save(doc, path)


def build_with_sections(path: Path) -> Path:
    doc = Document()
    doc.add_paragraph("Section one starts here.")
    second = doc.add_section(WD_SECTION_START.NEW_PAGE)
    second.orientation = WD_ORIENT.LANDSCAPE
    doc.add_paragraph("Section two starts here.")
    return _save(doc, path)


def build_with_tables(path: Path) -> Path:
    doc = Document()
    doc.add_paragraph("Table anchor paragraph.")
    first = doc.add_table(rows=2, cols=2)
    first.style = "Table Grid"
    first.cell(0, 0).text = "Name"
    first.cell(0, 1).text = "Value"
    first.cell(1, 0).text = "Budget"
    first.cell(1, 1).text = "100"
    second = doc.add_table(rows=2, cols=2)
    second.cell(0, 0).text = "Region"
    second.cell(0, 1).text = "Status"
    second.cell(1, 0).text = "West"
    second.cell(1, 1).text = "Ready"
    return _save(doc, path)


async def _apply_review_operations(path: Path) -> None:
    from app import create_server

    server = create_server()
    operations = [
        (
            "replace_text",
            {
                "path": str(path),
                "find_text": "draft",
                "replace_with": "approved",
                "track_changes": True,
                "author": "QA",
            },
        ),
        (
            "add_comment_to_matching_text",
            {
                "path": str(path),
                "target_text": "budget",
                "comment_text": "Confirm the source for this budget.",
                "author": "Editor",
                "initials": "ED",
            },
        ),
        (
            "insert_paragraph",
            {
                "path": str(path),
                "after_paragraph": 1,
                "text": "Inserted review note.",
                "track_changes": True,
                "author": "QA",
            },
        ),
    ]
    for name, arguments in operations:
        _, result = await server.call_tool(name, arguments)
        if result["status"] != "ok":
            raise RuntimeError(f"Failed to build review fixture with {name}: {result}")


def build_review(path: Path) -> Path:
    build_simple(path)
    asyncio.run(_apply_review_operations(path))
    return path


def build_complex(path: Path) -> Path:
    doc = Document()
    doc.add_heading("Complex sample", level=1)
    doc.add_paragraph("Alpha project budget is 100.")
    doc.add_paragraph("The draft conclusion needs review.")
    callout = doc.styles.add_style("ComplexCallout", 1)
    callout.font.bold = True
    styled = doc.add_paragraph("Styled paragraph in complex sample.")
    styled.style = "ComplexCallout"
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Risk"
    table.cell(0, 1).text = "Level"
    table.cell(1, 0).text = "Budget"
    table.cell(1, 1).text = "Medium"
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    doc.add_paragraph("Second section paragraph.")
    return _save(doc, path)


def generate_all_fixtures(target_dir: Path) -> None:
    builders = {
        "simple.docx": build_simple,
        "with_styles.docx": build_with_styles,
        "with_sections.docx": build_with_sections,
        "with_tables.docx": build_with_tables,
        "with_review.docx": build_review,
        "complex.docx": build_complex,
    }
    target_dir.mkdir(parents=True, exist_ok=True)
    for filename, builder in builders.items():
        builder(target_dir / filename)

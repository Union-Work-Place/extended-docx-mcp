"""Compatibility facade for legacy imports from ``document_ops``.

The implementation now lives in smaller modules:

- ``text_ops`` for search and replacement helpers
- ``structure_ops`` for paragraph, table, style and section helpers
"""

from ops.structure_ops import (
    apply_paragraph_format,
    apply_run_format,
    ensure_style_type_is_paragraph,
    find_paragraph,
    insert_structured_block_after,
    insert_paragraph_after,
    insert_table_after,
    iter_paragraphs,
    iter_paragraphs_in_parent,
    iter_tables,
    iter_tables_in_parent,
    paragraph_to_dict,
    parent_element,
    populate_paragraph,
    normalize_paragraph_range,
    section_to_dict,
    style_to_dict,
    table_to_dict,
    update_section_page_setup,
    validate_structured_blocks,
    write_cell_value,
    write_paragraph_block,
    write_structured_block,
    write_table_block,
)
from ops.text_ops import (
    enum_name,
    matches_for_text,
    normalize_mapping_value,
    replace_in_paragraph_plain,
    replace_matches_plain,
)

__all__ = [
    "apply_paragraph_format",
    "apply_run_format",
    "ensure_style_type_is_paragraph",
    "enum_name",
    "find_paragraph",
    "insert_structured_block_after",
    "insert_paragraph_after",
    "insert_table_after",
    "iter_paragraphs",
    "iter_paragraphs_in_parent",
    "iter_tables",
    "iter_tables_in_parent",
    "matches_for_text",
    "normalize_mapping_value",
    "paragraph_to_dict",
    "parent_element",
    "populate_paragraph",
    "normalize_paragraph_range",
    "replace_in_paragraph_plain",
    "replace_matches_plain",
    "section_to_dict",
    "style_to_dict",
    "table_to_dict",
    "update_section_page_setup",
    "validate_structured_blocks",
    "write_cell_value",
    "write_paragraph_block",
    "write_structured_block",
    "write_table_block",
]


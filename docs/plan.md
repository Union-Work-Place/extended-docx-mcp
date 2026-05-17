# Refinement Plan Based on MCP Testing Results

## Context

Testing with a real-world document demonstrated that basic reading and editing scenarios generally function correctly; however, two defects require immediate resolution:

1. High Priority: Crashes occurring in sections containing non-integer `w:pgMar` values.
2. Medium Priority: Discrepancies in paragraph indexing between the reading tools and the comment/revision tools.

Outlined below is a refinement plan designed to introduce minimal changes to the current architecture while ensuring comprehensive regression test coverage.

## Refinement Goals

- Eliminate crashes in documents that open successfully in Word but contain non-standard fractional values ​​for section margins.
- Ensure graceful degradation (without raising exceptions) for the `read_docx`, `list_sections`, and `insert_table` functions.
- Standardize paragraph indexing rules across all tools where the user supplies or retrieves a `paragraph_index`.
- Add tests that reproduce both defects and serve as safeguards against future regressions.

## Problem 1: Sections Crash on Fractional Margin Values

### Symptoms

- `list_sections` crashes during section serialization.
- `read_docx` crashes only when `include_sections=true`.
- `insert_table` crashes internally within `python-docx` during the `Document.add_table()` call.

### Confirmed Cause

- The document contains a fractional value for `w:pgMar/@w:left`—for example, `1984.251968503937`.
- `python-docx` expects integer twips and attempts to cast the value using `int(...)`, which results in a `ValueError`.
- In the current codebase, both direct access to `section.left_margin.pt` and indirect access via `document.add_table()` lack proper error handling. ### Affected Locations

- `src/ops/structure_ops.py:503` (`section_to_dict`)
- `src/toolsets/sections.py:23` (`list_sections`)
- `src/toolsets/content_tools/discovery.py:115` (`read_docx`)
- `src/ops/structure_ops.py:382` (`insert_table_after` via `document.add_table()`)

## Plan to Fix Problem 1

### 1. Add Safe Reading of Section Dimensions from OOXML

A single low-level helper for sections needs to be introduced—one that does not rely on directly calling `python-docx`'s margin/page setup properties in cases where the document might contain corrupt numeric values.

What to do:

- Add a helper to the `ops` layer that reads the raw XML of a section and safely parses the numeric attributes within `pgSz` and `pgMar`.
- For section margins, employ the following strategy:
- First, attempt a safe read using `python-docx`. 
- If a `ValueError` or `TypeError` occurs, read the value directly from the XML. 
- If the value is fractional, convert it to points using `float(value)` and a conversion from twips. 
- If the value cannot be parsed at all, return `None` rather than crashing.
- Explicitly distinguish between two cases:
- A valid value obtained from `python-docx`. 
- A fallback value obtained directly from the XML.

Expected Result:

- `section_to_dict()` always returns a serializable structure, even when processing a "dirty" DOCX file.
- `list_sections` and `read_docx(include_sections=true)` no longer crash.

### 2. Rewrite `section_to_dict()` for Safe Serialization

**What to do:**

- Replace direct attribute access—such as `section.left_margin.pt`—with a safe helper function.
- Similarly, validate `page_width`, `page_height`, `right_margin`, `top_margin`, and `bottom_margin` to ensure serialization is consistent across all fields, rather than fixing only `left_margin`.
- Preserve the current response format to avoid breaking existing clients.
- If a specific field cannot be read, return `None` instead of causing the entire tool to crash.

**Optional:**

- If necessary, add a metadata field—such as `warnings` or `source="xml-fallback"`—to the section response, but only if doing so does not complicate the API contract. If it is possible to proceed without this, it is better to keep the format minimal.

### 3. Decouple `insert_table` from the Problematic `_block_width` Read

The current failure of `insert_table` is not caused by the table logic itself, but rather by an internal section-reading operation within the `python-docx` library.

You must choose one of the following two approaches and adhere to it during implementation:

1. **Preferred Approach:** Normalize any invalid margin values ​​within the XML *before* calling `document.add_table()`.
2. **Alternative Approach:** Insert the table directly via OOXML, bypassing the `Document.add_table()` method entirely.

The preferred path is more lightweight and aligns better with the current architecture.

**What to do for the preferred option:**

- Before calling `document.add_table()`, check whether the document's last section contains any invalid margin attributes.
- If invalid attributes are found, carefully normalize *only* the problematic values ​​within the section's XML:
- Read the raw attribute value. 
- If the value is a float-like string, convert it to an integer (in twips) using a clearly defined rounding rule. 
- Do *not* modify any other section parameters.
- Once normalized, proceed with the standard `python-docx` execution path. Why this is better:

- Minimal deviation from the current code;
- No need to manually construct XML tables;
- Lower risk of regressions regarding styles, widths, and the internal `python-docx` model.

### 4. Define a Normalization Rule for Fractional Twips

A rule must be established in advance to ensure that both the code and the tests verify the exact same behavior.

The recommended rule is as follows:

- If a value can be interpreted as a `float`, round it to the nearest integer twip value using `round()`.
- Subsequently, use this integer twip value for all further operations.
- If a value is negative in a context where negative margins are disallowed, do not silently correct it unless absolutely necessary; instead, either leave it as is during serialization (represented as `None`) or validate it separately, specifically within the "mutating path" (i.e., when modifying the document).

This rule is sufficiently simple and predictable to serve as a reliable basis for regression testing.

### 5. Limit the Scope of the Fix

It is crucial to avoid scope creep in this task.

This enhancement *should not* involve:

- Attempting to "fix the entire DOCX file" wholesale upon opening.
- Automatically resaving the document solely for the purpose of reading it.
- Massively rewriting the section/page setup API.

The scope should be limited to the following:

- Safely reading the document.
- Performing localized normalization immediately prior to specific points where `python-docx` would otherwise raise an error.
- Ensuring adequate test coverage for these changes.

## Problem 2: Inconsistent Paragraph Indexing Rules

### Symptoms

- `find_text_occurrences("INTRODUCTION")` returns a specific `paragraph_index`.
- `add_comment_to_matching_text("INTRODUCTION")` places the comment in a *different* paragraph.

### Confirmed Cause

- The reading tools utilize `iter_paragraphs(doc)`, which retrieves a sequence of paragraphs directly from the `python-docx` object model—including paragraphs located *inside* tables.
- The review/comment tools utilize `iter_document_paragraphs_xml(root)` (specifically `body.iter(w:p)`), which yields a *different* sequence of paragraphs.
- In documents containing tables at the very beginning of the file, this discrepancy breaks the consistent transfer of paragraph indices between these different toolsets. ### Affected Locations

- `src/ops/review/xml_utils.py:156`
- `src/ops/review/comments.py:226`
- `src/toolsets/content_tools/discovery.py:64`

## Plan to Fix Issue 2

### 1. Establish a Unified Indexing Model

We need to adopt a single rule and apply it across all user-facing tools.

Recommended Rule:

- The canonical paragraph index must match the index returned by the reading tools via `iter_paragraphs(doc)`;
- The comment tools and revision tools must adapt to this model, rather than the other way around.

Rationale:

- The reading/discovery tools are the ones that first expose the `paragraph_index` to the user;
- The user workflow typically begins with operations such as `find_*`, `extract_text`, `get_paragraph_range`, and `read_docx`;
- Changing the indexing model within the reading tools—which is already visible to users—carries a higher risk.

### 2. Add a Mapping Layer Between `python-docx` Paragraphs and XML Paragraphs

We need a helper utility that establishes a mapping between:

- The paragraph index as yielded by `iter_paragraphs(doc)`;
- The corresponding `w:p` XML element;
- (If necessary) the index within the XML sequence.

What to do:

- Create a unified helper within `ops` that returns paragraph descriptors in canonical document order;
- Use this helper in both the discovery/read path and the review/comment path;
- Eliminate situations where one tool operates according to the `python-docx` order while another operates according to the `body.iter(w:p)` order without proper mapping.

In practical terms, this can be implemented with minimal effort:

- Either by constructing a list of XML elements in the same order as `iter_paragraphs(doc)`;
- Or by mapping based on the underlying XML node of the `Paragraph` objects.

The second option is preferable, provided it can be implemented without relying on fragile logic.

### 3. Migrate comment tools to canonical indices

What needs to change:

- `find_text_range_xml()` must accept and interpret `paragraph_index` using the same coordinate system as `find_text_occurrences()`;
- `add_comment`, `add_comment_to_matching_text`, and `add_comment_to_text_range` must select the target XML paragraph via the canonical mapping;
- If `anchor_text` is used, the search must also be performed against the same canonical sequence of paragraphs.

Expected result:

- An index obtained from `find_text_occurrences` can be passed directly to the comment tools;
- Users will see comments placed in the expected location, even in documents containing tables.

### 4. Check revision tools for the same issue

Even though the bug has been explicitly identified in the comment tools, the revision tools should be checked immediately as well, since they operate on the same XML layer.

What to do:

- Review all review tools that involve `paragraph_index` or paragraph context;
- Verify that `get_revision_details` and related helpers do not rely on a separate, incompatible indexing scheme;
- If they do, migrate them to use the same canonical helper as part of this task. ## Test Plan

### 1. Regression Test for Fractional `w:left` Values

A DOCX fixture needs to be added where a section's `w:pgMar/@w:left` attribute contains a fractional value.

Options for preparing the fixture:

- **Preferred:** Assemble the DOCX from an existing fixture and make targeted XML modifications using the test builder.
- **Acceptable:** Add a pre-built fixture file, provided it is small and stable.

What the tests should verify:

- `list_sections` returns `status == "ok"` and serializes the section without crashing.
- `read_docx(include_sections=true)` returns `status == "ok"`.
- `insert_table` successfully inserts a table into a copy of such a document.
- Furthermore, the resulting document can be successfully re-read using the `list_tables` or `read_docx` tools.

It is best to split this into two separate tests:

- One covering the read/list path.
- One covering the mutating path (`insert_table`).

### 2. Integration Test for Paragraph Index Consistency

A document containing a table positioned *before* the target paragraph is required to reproduce the current defect.

What to verify:

- `find_text_occurrences` returns a match with a specific `paragraph_index`.
- This index is passed to either `add_comment_to_matching_text` or `add_comment_to_text_range`.
- The comment is, in fact, placed within the correct "canonical" paragraph.

Minimum test scenario:

1. A fixture containing a table at the beginning of the document.
2. Unique text located in a standard paragraph following the table.
3. `find_text_occurrences` successfully locates this text.
4. `add_comment_to_matching_text`—using the same target text and/or `paragraph_index`—successfully places a comment.
5. Verification (via `list_comments` and, if necessary, by inspecting the XML anchor) confirms that the comment is correctly bound to the intended paragraph.

### 3. Smoke Testing Existing Scenarios

After applying the fixes, the existing tests must be re-verified for at least the following modules:

- `tests/test_sections.py`
- `tests/test_review.py`
- `tests/test_tables.py`
- `tests/test_content_reading.py`

Then, run the full `pytest` suite to ensure that the unification of paragraph mapping has not introduced any hidden regressions.

## Implementation Plan

### Phase 1: Fixing Sections

- Add safe parsing and fallback reading for page setup values.
- Migrate `section_to_dict()` to use safe serialization.
- Fix `insert_table` by performing local normalization of problematic margin values ​​prior to calling `document.add_table()`.
- Add regression tests for fractional margin values.

Completion Criteria:

- `list_sections`, `read_docx(include_sections=true)`, and `insert_table` no longer fail when run against the problematic document.

### Phase 2: Unifying Paragraph Indices

- Introduce a canonical paragraph mapping helper.
- Migrate the comment tools to use the canonical paragraph order.
- Review the revision tools and, if necessary, migrate them to use the same helper.
- Add an integration test for a document containing tables.

Completion Criteria:

- Indices returned by `find_text_occurrences` can be used within the comment workflow without unexpected issues.

### Phase 3: Final Validation

- Run the full `pytest` suite.
- Repeat the manual verification process using a copy of a real-world document, checking the following functions:
- `read_docx(include_sections=true)`
- `list_sections`
- `insert_table`
- `find_text_occurrences`
- `add_comment_to_matching_text`
- Document any remaining edge cases in the limitations section of the documentation. ## Risks and Solutions

### Risk 1. `python-docx` may fail on more than just `left_margin`

Solution:

- Create a helper function that handles not just a single field, but all numeric section properties that are serialized or read within performance-critical code paths.

### Risk 2. XML normalization prior to table insertion may inadvertently alter the document

Solution:

- Modify only the specific problematic numeric attributes;
- Avoid preemptively altering the document during simple read operations;
- Cover the "mutating path" with a dedicated test case.

### Risk 3. Unifying paragraph indices may impact existing review tests

Solution:

- Migrate all user-facing review tools to use a single helper function in a single pass;
- Avoid leaving a mixed indexing model in place across different tools.

## Completion Criteria

- A test case reproducing a fractional `w:left` value has been added and passes successfully.
- A test case reproducing paragraph index discrepancies caused by tables has been added and passes successfully.
- `list_sections`, `read_docx(include_sections=true)`, and `insert_table` do not fail when run against the problematic document.
- `find_text_occurrences` and `add_comment_to_matching_text` utilize a compatible indexing scheme.
- A full `pytest` run completes successfully without introducing any new regressions.

## Items to Defer Beyond the Scope of This Task

- Extended diagnostics for all non-standard section attributes within API responses;
- Automatic repair of any malformed OOXML values ​​upon file opening;
- A separate public API specifically for "repairing" problematic DOCX files;
- Broader unification of indexing schemes across all internal helper functions that are not exposed externally via the MCP contract.

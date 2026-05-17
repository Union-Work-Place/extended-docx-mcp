# Примеры: рецензирование

## `add_comment`

```json
{
  "name": "add_comment",
  "arguments": {
    "path": "files/sample.docx",
    "paragraph_index": 4,
    "comment_text": "Нужно уточнить формулировку",
    "author": "Editor",
    "initials": "ED"
  }
}
```

## `add_comment_to_text_range`

```json
{
  "name": "add_comment_to_text_range",
  "arguments": {
    "path": "files/sample.docx",
    "anchor_text": "цифрового продвижения",
    "start_offset": 0,
    "end_offset": 22,
    "comment_text": "Добавить источник",
    "author": "Editor",
    "initials": "ED"
  }
}
```

## `add_comment_to_matching_text`

```json
{
  "name": "add_comment_to_matching_text",
  "arguments": {
    "path": "files/sample.docx",
    "target_text": "маржинальность",
    "occurrence_index": 0,
    "comment_text": "Проверить расчёт",
    "author": "QA",
    "initials": "QA"
  }
}
```

## `list_revisions` и `get_revision_details`

```json
{
  "name": "get_revision_details",
  "arguments": {
    "path": "files/sample.docx",
    "revision_index": 0,
    "context_paragraphs": 1
  }
}
```

## `accept_all_revisions` / `reject_all_revisions`

```json
{
  "name": "accept_all_revisions",
  "arguments": {
    "path": "files/sample.docx",
    "output_path": "files/sample-reviewed.docx"
  }
}
```

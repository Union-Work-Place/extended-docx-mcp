# Примеры: таблицы

## `list_tables`

```json
{
  "name": "list_tables",
  "arguments": {
    "path": "files/sample.docx",
    "include_cells": true
  }
}
```

## `get_table_cell_content`

```json
{
  "name": "get_table_cell_content",
  "arguments": {
    "path": "files/sample.docx",
    "table_index": 0,
    "row_index": 1,
    "cell_index": 2
  }
}
```

## `insert_table`

```json
{
  "name": "insert_table",
  "arguments": {
    "path": "files/sample.docx",
    "after_paragraph": 3,
    "alignment": "center",
    "data": [
      ["Показатель", "Q1", "Q2"],
      ["Выручка", "120", "145"]
    ]
  }
}
```

## `update_table_cell`

```json
{
  "name": "update_table_cell",
  "arguments": {
    "path": "files/sample.docx",
    "table_index": 0,
    "row_index": 1,
    "cell_index": 1,
    "text": "150",
    "track_changes": true
  }
}
```

## `set_table_format`

```json
{
  "name": "set_table_format",
  "arguments": {
    "path": "files/sample.docx",
    "table_index": 0,
    "style_name": "Table Grid",
    "alignment": "left",
    "allow_auto_fit": true
  }
}
```

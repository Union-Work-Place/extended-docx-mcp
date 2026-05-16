# Примеры: создание и редактирование

## `write_docx`

```json
{
  "name": "write_docx",
  "arguments": {
    "path": "files/new.docx",
    "mode": "replace",
    "title": "Отчёт",
    "blocks": [
      {"type": "heading", "level": 1, "text": "Отчёт по проекту"},
      {"type": "paragraph", "text": "Краткое описание результата"},
      {"type": "table", "rows": [["Показатель", "Значение"], ["ROI", "18%"]]}
    ]
  }
}
```

## `replace_text`

```json
{
  "name": "replace_text",
  "arguments": {
    "path": "files/sample.docx",
    "find_text": "нейросеть",
    "replace_with": "нейронная сеть",
    "track_changes": true,
    "author": "DOCX MCP"
  }
}
```

В режиме `track_changes=true` правка записывается через OOXML revisions.

## `insert_paragraph`

```json
{
  "name": "insert_paragraph",
  "arguments": {
    "path": "files/sample.docx",
    "anchor_text": "Раздел 2",
    "text": "Добавленный поясняющий абзац",
    "track_changes": false
  }
}
```

## `delete_paragraph`

```json
{
  "name": "delete_paragraph",
  "arguments": {
    "path": "files/sample.docx",
    "paragraph_index": 8,
    "track_changes": true,
    "author": "Reviewer"
  }
}
```

## Форматирование

Для одиночных точечных правок доступны `set_paragraph_format`, `set_run_format` и `apply_paragraph_style`. Для пакетных и диапазонных правок используйте инструменты из этапа развития сервера.

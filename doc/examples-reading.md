# Примеры: чтение и поиск

## `read_docx`

```json
{
  "name": "read_docx",
  "arguments": {
    "path": "files/sample.docx",
    "start_paragraph": 0,
    "paragraph_count": 5,
    "include_runs": true,
    "include_tables": true
  }
}
```

Используйте для быстрой структурной сводки по документу, когда нужно увидеть абзацы, таблицы, секции и revision-aware текст.

## `extract_text`

```json
{
  "name": "extract_text",
  "arguments": {
    "path": "files/sample.docx",
    "start_paragraph": 10,
    "count": 3
  }
}
```

Подходит для чтения диапазона абзацев без полного структурного дампа.

## `find_text_occurrences`

```json
{
  "name": "find_text_occurrences",
  "arguments": {
    "path": "files/sample.docx",
    "target_text": "нейросеть",
    "find_whole_words_only": true,
    "max_results": 20
  }
}
```

Инструмент возвращает индексы абзацев, смещения и короткий preview — удобно перед заменой текста или постановкой комментария.

## `find_paragraphs`

```json
{
  "name": "find_paragraphs",
  "arguments": {
    "path": "files/sample.docx",
    "search_text": "Итоги квартала",
    "style_name": "Heading 2",
    "max_results": 10
  }
}
```

Можно искать только по тексту, только по стилю или комбинировать оба фильтра.

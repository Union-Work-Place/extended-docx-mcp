# extended-docx-mcp

MCP-сервер для чтения и редактирования `.docx` на базе `python-docx` и `lxml`.
Проект запускает локальный `stdio`-сервер и даёт инструменты для работы со
структурой документа, таблицами, стилями, комментариями и tracked revisions.

## Что умеет сервер

### Метаданные и обзор документа
- `server_info` — вернуть сведения о сервере и доступных группах инструментов
- `inspect_document` — собрать краткую сводку по документу: абзацы, таблицы, секции, комментарии и правки
- `read_docx` и `extract_text` — прочитать структуру документа или диапазон абзацев
- `find_text_occurrences` и `find_paragraphs` — искать текст и абзацы по содержимому и стилю

### Редактирование содержимого
- `write_docx` — создать новый документ из структурных блоков
- `replace_text` — заменить текст в обычном режиме или через tracked revisions
- `insert_paragraph` и `delete_paragraph` — вставлять и удалять абзацы
- `set_paragraph_format` и `set_run_format` — менять форматирование абзацев и runs

### Таблицы, стили и секции
- `list_tables`, `get_table_cell_content`, `insert_table`, `update_table_cell`, `set_table_format`
- `list_styles`, `create_or_update_style`, `apply_paragraph_style`
- `list_sections`, `set_section_page_setup`

### Рецензирование
- `list_comments`, `add_comment`, `add_comment_to_text_range`, `add_comment_to_matching_text`, `add_comment_reply`
- `list_revisions`, `get_revision_details`, `accept_all_revisions`, `reject_all_revisions`

## Структура проекта

- `src/app.py` — фабрика FastMCP-сервера и основной запуск
- `src/server.py` — совместимая точка входа для `python -m server`
- `src/toolsets/` — регистрация MCP-инструментов по группам
- `src/ops/` — операции над DOCX и OOXML
- `scripts/setup.ps1` — локальная подготовка Windows-окружения

## Быстрый старт

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e .
.\.venv\Scripts\python.exe -m server
```

Альтернатива для Windows:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\setup.ps1
```

После установки доступен консольный entrypoint:

```powershell
.\.venv\Scripts\extended-docx-mcp.exe
```

## Подключение в MCP-клиент

Пример `mcp.json` для локального запуска:

```json
{
  "servers": {
    "docx-local": {
      "type": "stdio",
      "command": "${workspaceFolder}/.venv/Scripts/python.exe",
      "args": ["-m", "server"],
      "env": {
        "EXTENDED_DOCX_MCP_DEFAULT_DIR": "${workspaceFolder}",
        "PYTHONPATH": "${workspaceFolder}/src",
        "PYTHONUNBUFFERED": "1"
      }
    }
  }
}
```

## Поддерживаемые блоки `write_docx`

- `paragraph`
- `heading`
- `table`
- `page_break`
- `section_break`

## Примеры запросов

- Покажи первые 20 абзацев файла `files/sample.docx`
- Прочитай `files/sample.docx` как структурную модель без полного текста
- Найди абзацы со стилем `Heading 2`
- Замени термин `нейросеть` на `нейронная сеть` в режиме рецензирования
- Покажи содержимое ячейки таблицы 1, строка 0, столбец 2
- Добавь комментарий к фразе `цифрового продвижения`
- Покажи детали правки 3 с одним соседним абзацем контекста

## План работ

Актуальный план вынесен в [`doc/plan.md`](doc/plan.md).

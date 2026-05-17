# Примеры конфигурации MCP-клиентов

## Базовые переменные окружения

- `EXTENDED_DOCX_MCP_DEFAULT_DIR` — директория, от которой разрешаются относительные пути.
- `PYTHONPATH=${workspaceFolder}/src` — нужен при запуске из исходников без установки wheel.
- `PYTHONUNBUFFERED=1` — полезен для стабильного `stdio`.

## Claude Desktop / совместимые клиенты

```json
{
  "servers": {
    "extended-docx-mcp": {
      "type": "stdio",
      "command": "${workspaceFolder}/.venv/bin/python",
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

## Windows-конфигурация

```json
{
  "servers": {
    "extended-docx-mcp": {
      "type": "stdio",
      "command": "${workspaceFolder}\\\\.venv\\\\Scripts\\\\python.exe",
      "args": ["-m", "server"],
      "env": {
        "EXTENDED_DOCX_MCP_DEFAULT_DIR": "${workspaceFolder}",
        "PYTHONPATH": "${workspaceFolder}\\\\src",
        "PYTHONUNBUFFERED": "1"
      }
    }
  }
}
```

## Запуск через entrypoint

После `pip install -e .` можно запускать не `python -m server`, а консольный скрипт:

```json
{
  "servers": {
    "extended-docx-mcp": {
      "type": "stdio",
      "command": "${workspaceFolder}/.venv/bin/extended-docx-mcp",
      "env": {
        "EXTENDED_DOCX_MCP_DEFAULT_DIR": "${workspaceFolder}",
        "PYTHONUNBUFFERED": "1"
      }
    }
  }
}
```

## Проверка конфигурации

1. Убедитесь, что команда запуска существует в виртуальном окружении.
2. Запустите сервер локально и проверьте, что процесс стартует без ошибок.
3. Проверьте `server_info` — он должен вернуть имя сервера, версию и список групп инструментов.

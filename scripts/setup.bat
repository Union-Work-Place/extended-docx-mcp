@echo off
setlocal

set PROJECT_ROOT=%~dp0..
pushd "%PROJECT_ROOT%"

if not exist ".venv" (
    py -3.12 -m venv .venv
)

call ".venv\Scripts\python.exe" -m pip install --upgrade pip
call ".venv\Scripts\python.exe" -m pip install -e .

echo DOCX MCP server environment is ready.
popd

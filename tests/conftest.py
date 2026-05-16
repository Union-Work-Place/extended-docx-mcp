from __future__ import annotations

import shutil
import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
TESTS = ROOT / "tests"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
if str(TESTS) not in sys.path:
    sys.path.insert(0, str(TESTS))

from app import create_server
from fixture_builder import generate_all_fixtures


@pytest.fixture(scope="session")
def fixtures_dir() -> Path:
    target = ROOT / "tests" / "fixtures"
    generate_all_fixtures(target)
    return target


@pytest.fixture()
def copy_fixture(tmp_path: Path, fixtures_dir: Path):
    def _copy(name: str) -> Path:
        source = fixtures_dir / name
        target = tmp_path / name
        shutil.copy2(source, target)
        return target

    return _copy


@pytest.fixture()
def server():
    return create_server()


@pytest.fixture()
def invoke_tool(server):
    async def _invoke(name: str, **arguments: Any) -> dict[str, Any]:
        _, result = await server.call_tool(name, arguments)
        return result

    return _invoke

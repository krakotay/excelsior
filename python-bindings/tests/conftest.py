from __future__ import annotations

import os
import shutil

import pytest


@pytest.fixture(scope="session", autouse=True)
def ensure_libreoffice_available() -> None:
    if shutil.which("libreoffice") is None and shutil.which("soffice") is None:
        pytest.skip("libreoffice/soffice is required for integration tests", allow_module_level=True)


@pytest.fixture(scope="session", autouse=True)
def set_test_environment() -> None:
    # Isolate any runtime caches under /tmp for CI/sandbox friendliness.
    os.environ.setdefault("HOME", os.environ.get("HOME", "/tmp"))

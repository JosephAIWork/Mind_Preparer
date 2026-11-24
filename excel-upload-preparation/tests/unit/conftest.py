"""Every unit test writes into a folder of its own, never into the user's real data folder
(%LOCALAPPDATA%\\MindReady): no fake incident in feedback\\incidents.jsonl, no test timing in
perf\\timings.jsonl. `feedback._path()`, `perf.log_path()` and `excel_com` read
`SETTINGS.data_folder` at call time, so pointing it at tmp_path is enough for them. Modules that
copy the folder when imported (app/ai_key.py KEY_FILE, app/mind_client.py _ROOT, app/mind_mcp.py
TOKEN_FILE) are left as they are: no unit test writes through them."""
from __future__ import annotations

import pytest

from app.settings import SETTINGS


@pytest.fixture(autouse=True)
def _own_data_folder(tmp_path):
    # Settings is a frozen dataclass and every module holds the same SETTINGS object: the field is
    # set on that object for the test and put back after it.
    real = SETTINGS.data_folder
    object.__setattr__(SETTINGS, "data_folder", tmp_path / "MindReady")
    try:
        yield
    finally:
        object.__setattr__(SETTINGS, "data_folder", real)

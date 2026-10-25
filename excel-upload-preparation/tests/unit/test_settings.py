"""settings.yaml: the shipped file means exactly the old built-in values; mistakes stop the app with a
message; the sandbox guard and the API pages follow it (app/settings.py)."""
from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app import mind_client, settings
from app.settings import Settings, load
from app.web import server


def write(tmp_path: Path, text: str) -> Path:
    path = tmp_path / "settings.yaml"
    path.write_text(text, encoding="utf-8")
    return path


def test_the_shipped_file_means_the_old_built_in_values():
    s = load(settings.SETTINGS_FILE)
    assert s == Settings()
    assert s.mind_main_url == "https://shared.milliman-mind.com/main"
    assert s.mind_host_suffix == "milliman-mind.com"


def test_no_file_or_empty_values_mean_the_defaults(tmp_path):
    assert load(tmp_path / "missing.yaml") == Settings()
    assert load(write(tmp_path, "port:\nmind:\n  sandbox_folder: ''\n")) == Settings()


def test_values_are_read(tmp_path):
    s = load(write(tmp_path, "port: 8700\napi_pages: true\ndata_folder: local_data\nmind:\n  site: https://other.example.com/\n  project_prefix: test_mr\n"))
    assert s.port == 8700 and s.api_pages is True
    assert s.data_folder == settings.PACKAGE_ROOT / "local_data"
    assert s.mind_main_url == "https://other.example.com/main" and s.mind_host_suffix == "example.com"
    assert s.mind_project_prefix == "test_mr"


@pytest.mark.parametrize("text, words", [
    ("prot: 8700\n", "unknown setting 'prot'"),
    ("port: eighty\n", "whole number"),
    ("api_pages: yes please\n", "true or false"),
    ("mind:\n  project_prefix: z\n", "at least 3"),
    ("port: [1\n", "not valid YAML"),
])
def test_mistakes_stop_with_a_message(tmp_path, text, words):
    with pytest.raises(ValueError, match=words.replace("(", r"\(")):
        load(write(tmp_path, text))


def test_the_app_deletes_only_projects_with_the_prefix():
    assert mind_client.delete_project(None, "Client model 2026") == "refused"  # refused before any page is touched
    assert mind_client.PROJECT_PREFIX == settings.SETTINGS.mind_project_prefix


def test_the_api_pages_are_off_by_default():
    if settings.SETTINGS.api_pages:
        pytest.skip("api_pages is switched on in this settings.yaml")
    client = TestClient(server.app)
    assert server.app.openapi_url is None and server.app.docs_url is None
    for path in ("/openapi.json", "/docs", "/redoc"):
        res = client.get(path)
        assert '"openapi"' not in res.text and "swagger" not in res.text.lower(), path

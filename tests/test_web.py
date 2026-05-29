from __future__ import annotations

import shutil
import time
from io import BytesIO

from fastapi.testclient import TestClient
from openpyxl import Workbook

from web import app as web_app
from xlsx_ai_translate.workbook import TranslationStats


def test_index_page_renders(tmp_path):
    web_app.reset_for_tests(tmp_path / "data")

    with TestClient(web_app.app) as client:
        response = client.get("/")

    assert response.status_code == 200
    assert "Start translation" in response.text


def test_job_upload_runs_background_translation_and_downloads(tmp_path, monkeypatch):
    web_app.reset_for_tests(tmp_path / "data")

    def fake_translate_workbook(input_path, output_path, **kwargs):
        progress = kwargs.get("progress")
        if progress is not None:
            progress(1, 2)
            progress(2, 2)
        shutil.copyfile(input_path, output_path)
        return TranslationStats(
            sheets_seen=1,
            sheets_translated=1,
            string_cells_found=1,
            unique_strings=1,
            sheet_titles_found=1,
            sheet_titles_translated=1,
        )

    monkeypatch.setattr(web_app, "translate_workbook", fake_translate_workbook)

    with TestClient(web_app.app) as client:
        response = client.post(
            "/jobs",
            data={
                "target_language": "en",
                "source_language": "auto",
                "model": "openai/gpt-4o-mini",
                "translate_sheet_names": "true",
            },
            files={
                "file": (
                    "document.xlsx",
                    _workbook_bytes(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
            follow_redirects=False,
        )

        assert response.status_code == 303
        job_url = response.headers["location"]
        job_id = job_url.rsplit("/", 1)[-1]

        status = _wait_for_done(client, job_id)
        assert status["status"] == "done"
        assert status["progress_percent"] == 100
        assert status["stats"]["sheet_titles_translated"] == 1

        download = client.get(f"/jobs/{job_id}/download")
        assert download.status_code == 200
        assert download.content.startswith(b"PK")


def test_rejects_non_xlsx_upload(tmp_path):
    web_app.reset_for_tests(tmp_path / "data")

    with TestClient(web_app.app) as client:
        response = client.post(
            "/jobs",
            data={"target_language": "en"},
            files={"file": ("document.txt", b"hello", "text/plain")},
        )

    assert response.status_code == 400


def _workbook_bytes() -> BytesIO:
    workbook = Workbook()
    workbook.active["A1"] = "Bonjour"
    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return stream


def _wait_for_done(client: TestClient, job_id: str) -> dict[str, object]:
    deadline = time.time() + 3
    last_status: dict[str, object] = {}
    while time.time() < deadline:
        response = client.get(f"/jobs/{job_id}/status")
        response.raise_for_status()
        last_status = response.json()
        if last_status["status"] in {"done", "error"}:
            return last_status
        time.sleep(0.05)
    return last_status

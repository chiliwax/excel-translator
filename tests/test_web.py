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
    assert "Inspect workbook sheets" in response.text
    assert "Start translation" in response.text


def test_inspect_upload_returns_sheet_names(tmp_path):
    web_app.reset_for_tests(tmp_path / "data")

    with TestClient(web_app.app) as client:
        response = client.post(
            "/inspect",
            files={
                "file": (
                    "document.xlsx",
                    _workbook_bytes(["Keep", "Skip"]),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["filename"] == "document.xlsx"
    assert payload["sheets"] == ["Keep", "Skip"]
    assert len(payload["upload_id"]) == 32


def test_job_upload_runs_background_translation_and_downloads(tmp_path, monkeypatch):
    web_app.reset_for_tests(tmp_path / "data")
    captured_kwargs = {}

    def fake_translate_workbook(input_path, output_path, **kwargs):
        captured_kwargs.update(kwargs)
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
        inspect = client.post(
            "/inspect",
            files={
                "file": (
                    "document.xlsx",
                    _workbook_bytes(["Keep", "Skip"]),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
        )
        assert inspect.status_code == 200

        response = client.post(
            "/jobs",
            data={
                "upload_id": inspect.json()["upload_id"],
                "target_language": "en",
                "source_language": "auto",
                "model": "openai/gpt-4o-mini",
                "translate_sheet_names": "true",
                "exclude_sheet_names": "Skip",
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
        assert captured_kwargs["exclude_sheets"] == ["Skip"]

        download = client.get(f"/jobs/{job_id}/download")
        assert download.status_code == 200
        assert download.content.startswith(b"PK")


def test_rejects_non_xlsx_upload(tmp_path):
    web_app.reset_for_tests(tmp_path / "data")

    with TestClient(web_app.app) as client:
        response = client.post(
            "/inspect",
            files={"file": ("document.txt", b"hello", "text/plain")},
        )

    assert response.status_code == 400


def test_download_button_is_hidden_until_ready(tmp_path):
    web_app.reset_for_tests(tmp_path / "data")
    job = web_app.JobState(
        id="a" * 32,
        status="queued",
        input_filename="document.xlsx",
        source_language="auto",
        target_language="en",
        model="openai/gpt-4o-mini",
        exclude_sheets=[],
        translate_sheet_names=True,
        batch_size=50,
        max_batch_chars=12_000,
        concurrency=8,
        rpm=500,
        tpm=200_000,
        request_timeout=120,
    )
    web_app._set_job(job)
    web_app._write_job_file(job)

    with TestClient(web_app.app) as client:
        page = client.get(f"/jobs/{job.id}")
        css = client.get("/static/app.css")

    assert 'id="download"' in page.text
    assert "hidden" in page.text
    assert "[hidden]" in css.text


def _workbook_bytes(sheet_names: list[str] | None = None) -> BytesIO:
    workbook = Workbook()
    names = sheet_names or ["Sheet"]
    workbook.active.title = names[0]
    workbook.active["A1"] = "Bonjour"
    for name in names[1:]:
        workbook.create_sheet(name)["A1"] = "Bonjour"
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

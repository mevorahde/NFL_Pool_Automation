import os

import pandas as pd
from dotenv import dotenv_values

import pool


def test_creates_current_pacific_year_workbook_and_updates_only_file_path(
    tmp_path, monkeypatch
):
    template = tmp_path / "Family Football Pool Template.xlsx"
    template.write_bytes(b"clean-template")
    env_path = tmp_path / ".env"
    env_path.write_text(
        "SMTP_SERVER=smtp.example.test\nfile_path=old-workbook.xlsx\n",
        encoding="utf-8",
    )
    monkeypatch.delenv("file_path", raising=False)

    target = pool.ensure_current_year_workbook(
        template_path=template,
        env_path=env_path,
        now=pd.Timestamp("2027-01-01T07:30:00Z"),
    )

    assert target == tmp_path / "Family Football Pool 2026.xlsx"
    assert target.read_bytes() == b"clean-template"
    values = dotenv_values(env_path)
    assert values["file_path"] == str(target.resolve())
    assert values["SMTP_SERVER"] == "smtp.example.test"
    assert os.environ["file_path"] == str(target.resolve())


def test_existing_year_workbook_is_never_overwritten(tmp_path, monkeypatch):
    template = tmp_path / "Family Football Pool Template.xlsx"
    template.write_bytes(b"new-template")
    target = tmp_path / "Family Football Pool 2026.xlsx"
    target.write_bytes(b"existing-season-data")
    env_path = tmp_path / ".env"
    env_path.write_text("file_path=old-workbook.xlsx\n", encoding="utf-8")
    monkeypatch.delenv("file_path", raising=False)

    result = pool.ensure_current_year_workbook(
        template_path=template,
        env_path=env_path,
        now=pd.Timestamp("2026-08-09T12:00:00Z"),
    )

    assert result == target
    assert target.read_bytes() == b"existing-season-data"
    assert dotenv_values(env_path)["file_path"] == str(target.resolve())


def test_production_main_ensures_the_current_year_workbook():
    assert "ensure_current_year_workbook" in pool.main.__code__.co_names

from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest
import pytz
from bs4 import BeautifulSoup

import pool


PACIFIC = pytz.timezone("America/Los_Angeles")
FIXTURE_DIR = Path(__file__).parent / "mock_html"


def schedule(*kickoffs):
    return pd.DataFrame(
        {
            "Team1": [f"AWAY-{number}" for number in range(len(kickoffs))],
            "Team2": [f"HOME-{number}" for number in range(len(kickoffs))],
            "UTC_DateTime": list(kickoffs),
        },
        index=[100 + number for number in range(len(kickoffs))],
    )


def fixture_schedule(filename):
    soup = BeautifulSoup((FIXTURE_DIR / filename).read_text(encoding="utf-8"), "html.parser")
    parsed = [pool.parse_game_card(card) for card in soup.find_all("div", class_="event-card")]
    return pd.DataFrame(
        {
            "Team1": [game[0] for game in parsed],
            "Team2": [game[2] for game in parsed],
            "UTC_DateTime": [game[6] for game in parsed],
        }
    )


def mapping(frame):
    return dict(zip(frame["Team1"], frame["Excel_Row"]))


def test_normal_schedule_uses_header_offset_and_preserves_scraped_order():
    full_schedule = schedule(
        "2026-09-13T20:25:00Z",
        "2026-09-10T00:20:00Z",
        "2026-09-14T00:15:00Z",
    )
    original = full_schedule.copy(deep=True)

    rows = pool.assign_excel_rows(full_schedule)

    assert rows.tolist() == [2, 3, 4]
    assert rows.index.tolist() == full_schedule.index.tolist()
    pd.testing.assert_frame_equal(full_schedule, original)


def test_started_games_do_not_shift_remaining_excel_rows():
    full_schedule = schedule(
        "2026-09-11T00:15:00Z",
        "2026-09-13T17:00:00Z",
        "2026-09-14T00:20:00Z",
    )

    remaining, _ = pool.prepare_schedule_for_excel(
        full_schedule,
        now=PACIFIC.localize(datetime(2026, 9, 11, 12, 0)),
    )

    assert mapping(remaining) == {"AWAY-1": 3, "AWAY-2": 4}


def test_matchups_keep_rows_across_repeated_runs_at_different_times():
    full_schedule = schedule(
        "2026-09-11T00:15:00Z",
        "2026-09-13T17:00:00Z",
        "2026-09-14T00:20:00Z",
    )
    before, _ = pool.prepare_schedule_for_excel(
        full_schedule,
        now=PACIFIC.localize(datetime(2026, 9, 10, 9, 0)),
    )
    after, _ = pool.prepare_schedule_for_excel(
        full_schedule,
        now=PACIFIC.localize(datetime(2026, 9, 13, 12, 0)),
    )

    assert mapping(before)["AWAY-2"] == 4
    assert mapping(after) == {"AWAY-2": 4}


def test_friday_game_consumes_its_full_schedule_row():
    full_schedule = schedule(
        "2026-11-27T01:15:00Z",  # Thursday Pacific
        "2026-11-28T01:00:00Z",  # Friday Pacific
        "2026-11-29T21:25:00Z",  # Sunday Pacific
    )

    assert pool.assign_excel_rows(full_schedule).tolist() == [2, 3, 4]


@pytest.mark.parametrize(
    "filename",
    ["thanksgiving.html", "saturday_tripleheader.html"],
)
def test_fixture_schedules_keep_every_game_in_fixture_order(filename):
    full_schedule = fixture_schedule(filename)

    assert pool.assign_excel_rows(full_schedule).tolist() == [2, 3, 4]


@pytest.mark.parametrize("bad_timestamp", [None, "not-a-kickoff"])
def test_invalid_or_missing_kickoff_timestamps_fail_clearly(bad_timestamp):
    full_schedule = schedule("2026-09-11T00:15:00Z", bad_timestamp)

    with pytest.raises(ValueError, match=r"UTC_DateTime.*101"):
        pool.assign_excel_rows(full_schedule)


def test_production_schedule_preparation_calls_row_assignment(monkeypatch):
    full_schedule = schedule("2026-09-13T17:00:00Z", "2026-09-14T00:20:00Z")
    original = full_schedule.copy(deep=True)
    calls = []

    def tested_assignment(frame, header_rows=1):
        calls.append((frame.copy(deep=True), header_rows))
        return pd.Series([7, 8], index=frame.index, name="Excel_Row")

    monkeypatch.setattr(pool, "assign_excel_rows", tested_assignment)

    prepared, _ = pool.prepare_schedule_for_excel(
        full_schedule,
        now=PACIFIC.localize(datetime(2026, 9, 10, 9, 0)),
    )

    assert len(calls) == 1
    assert calls[0][1] == 1
    assert prepared["Excel_Row"].tolist() == [7, 8]
    pd.testing.assert_frame_equal(full_schedule, original)
    assert "prepare_schedule_for_excel" in pool.main.__code__.co_names

from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest
from bs4 import BeautifulSoup

import pool
from pool import (
    apply_team_abbreviations,
    assign_excel_rows,
    get_week_number,
    parse_game_card,
)


TEST_DIR = Path(__file__).parent
MOCK_HTML_DIR = TEST_DIR / "mock_html"
EXPECTED_SCHEDULE_PATH = TEST_DIR / "test_runs" / "test_schedule.xlsx"
GAME_COLUMNS = [
    "Team1",
    "Spread",
    "Team2",
    "Team1_Abbr",
    "Team2_Abbr",
    "Home_Team",
    "UTC_DateTime",
    "Favorite_Side",
]

OFFICIAL_TEAM_ABBREVIATIONS = {
    "ARIZONA CARDINALS": "ARI",
    "ATLANTA FALCONS": "ATL",
    "BALTIMORE RAVENS": "BAL",
    "BUFFALO BILLS": "BUF",
    "CAROLINA PANTHERS": "CAR",
    "CHICAGO BEARS": "CHI",
    "CINCINNATI BENGALS": "CIN",
    "CLEVELAND BROWNS": "CLE",
    "DALLAS COWBOYS": "DAL",
    "DENVER BRONCOS": "DEN",
    "DETROIT LIONS": "DET",
    "GREEN BAY PACKERS": "GB",
    "HOUSTON TEXANS": "HOU",
    "INDIANAPOLIS COLTS": "IND",
    "JACKSONVILLE JAGUARS": "JAC",
    "KANSAS CITY CHIEFS": "KC",
    "LAS VEGAS RAIDERS": "LV",
    "LOS ANGELES CHARGERS": "LAC",
    "LOS ANGELES RAMS": "LAR",
    "MIAMI DOLPHINS": "MIA",
    "MINNESOTA VIKINGS": "MIN",
    "NEW ENGLAND PATRIOTS": "NE",
    "NEW ORLEANS SAINTS": "NO",
    "NEW YORK GIANTS": "NYG",
    "NEW YORK JETS": "NYJ",
    "PHILADELPHIA EAGLES": "PHI",
    "PITTSBURGH STEELERS": "PIT",
    "SAN FRANCISCO 49ERS": "SF",
    "SEATTLE SEAHAWKS": "SEA",
    "TAMPA BAY BUCCANEERS": "TB",
    "TENNESSEE TITANS": "TEN",
    "WASHINGTON COMMANDERS": "WAS",
}


@pytest.fixture(autouse=True)
def prevent_error_email(monkeypatch):
    monkeypatch.setattr(pool, "send_error_email", lambda *args, **kwargs: None)


def load_mock_html(filename):
    return BeautifulSoup(
        (MOCK_HTML_DIR / filename).read_text(encoding="utf-8"),
        "html.parser",
    )


def load_fixture_games(filename):
    soup = load_mock_html(filename)
    rows = [parse_game_card(card) for card in soup.find_all("div", class_="event-card")]
    assert all(row is not None for row in rows), f"Failed to parse {filename}"
    return pd.DataFrame(rows, columns=GAME_COLUMNS)


def load_complete_edge_schedule():
    schedule = pd.concat(
        [
            load_fixture_games("thanksgiving.html"),
            load_fixture_games("black_friday.html"),
            load_fixture_games("saturday_tripleheader.html"),
        ],
        ignore_index=True,
    )
    schedule = apply_team_abbreviations(schedule)
    schedule["Excel_Row"] = assign_excel_rows(schedule)
    return schedule


def load_corrected_expected_schedule():
    expected = pd.read_excel(EXPECTED_SCHEDULE_PATH).rename(
        columns={"Excel_Row": "Legacy_Excel_Row"}
    )
    # The tracked workbook's team/day order is canonical. Its stored row values
    # predate the row-2 header offset and Friday inclusion, so derive corrected
    # expectations from the complete seven-game fixture order.
    expected["Excel_Row"] = range(2, len(expected) + 2)
    return expected


def assert_rows_match_tracked_schedule(game_day):
    actual = load_complete_edge_schedule()
    actual = actual[
        actual["UTC_DateTime"].map(lambda kickoff: kickoff.strftime("%A")) == game_day
    ]
    expected = load_corrected_expected_schedule()
    expected = expected[expected["Expected_Weekday"] == game_day]

    actual_rows = list(
        zip(actual["Team1_Abbr"], actual["Team2_Abbr"], actual["Excel_Row"])
    )
    expected_rows = list(
        zip(expected["Team1"], expected["Team2"], expected["Excel_Row"])
    )
    assert actual_rows == expected_rows


@pytest.mark.parametrize(
    "filename,expected_week",
    [
        ("thanksgiving.html", 12),
        ("friday_game.html", 13),
        ("black_friday.html", 13),
        ("christmas_tuesday.html", 16),
        ("christmas_wednesday.html", 16),
        ("saturday_tripleheader.html", 15),
    ],
)
def test_week_extraction(filename, expected_week):
    soup = load_mock_html(filename)
    week = get_week_number(soup)
    assert int(week) == expected_week


def test_row_assignment_thanksgiving():
    games = load_fixture_games("thanksgiving.html")
    weekdays = games["UTC_DateTime"].map(lambda kickoff: kickoff.strftime("%A"))
    assert weekdays.tolist() == ["Thursday", "Thursday", "Thursday"]


def test_row_assignment_saturday_tripleheader():
    games = load_fixture_games("saturday_tripleheader.html")
    weekdays = games["UTC_DateTime"].map(lambda kickoff: kickoff.strftime("%A"))
    assert weekdays.tolist() == ["Saturday", "Saturday", "Saturday"]


def test_abbreviation_mapping():
    df = load_fixture_games("christmas_tuesday.html")
    df = apply_team_abbreviations(df)
    assert df[["Team1_Abbr", "Team2_Abbr"]].iloc[0].to_dict() == {
        "Team1_Abbr": "KC",
        "Team2_Abbr": "DEN",
    }


def test_all_official_full_team_names_map_explicitly_and_preserve_valid_abbreviations():
    names = list(OFFICIAL_TEAM_ABBREVIATIONS)
    abbreviations = list(OFFICIAL_TEAM_ABBREVIATIONS.values())
    df = pd.DataFrame(
        {
            "Team1": [f"  {name.lower()}  " for name in names],
            "Team2": names,
            "Team1_Abbr": [f" {abbr.lower()} " for abbr in abbreviations],
            "Team2_Abbr": abbreviations,
        }
    )

    result = apply_team_abbreviations(df)

    assert result["Team1"].tolist() == names
    assert result["Team2"].tolist() == names
    assert result["Team1_Abbr"].tolist() == abbreviations
    assert result["Team2_Abbr"].tolist() == abbreviations


def test_supported_nicknames_still_map_after_case_and_whitespace_normalization():
    df = pd.DataFrame(
        {
            "Team1": ["  chiefs  "],
            "Team2": [" broncos "],
            "Team1_Abbr": [None],
            "Team2_Abbr": [None],
        }
    )

    result = apply_team_abbreviations(df)

    assert result[["Team1", "Team2"]].iloc[0].to_dict() == {
        "Team1": "CHIEFS",
        "Team2": "BRONCOS",
    }
    assert result[["Team1_Abbr", "Team2_Abbr"]].iloc[0].to_dict() == {
        "Team1_Abbr": "KC",
        "Team2_Abbr": "DEN",
    }


def test_misspelled_or_unknown_full_names_remain_unmapped(caplog):
    df = pd.DataFrame(
        {
            "Team1": ["  dalls cowboys  "],
            "Team2": [" unknown cowboys "],
            "Team1_Abbr": ["DAL"],
            "Team2_Abbr": ["DAL"],
        }
    )

    result = apply_team_abbreviations(df)

    assert result[["Team1_Abbr", "Team2_Abbr"]].isna().all().all()
    assert "DALLS COWBOYS" in caplog.text
    assert "UNKNOWN COWBOYS" in caplog.text


def test_black_friday_single_game():
    games = load_fixture_games("black_friday.html")
    weekdays = games["UTC_DateTime"].map(lambda kickoff: kickoff.strftime("%A"))
    assert weekdays.tolist() == ["Friday"]


def test_excel_row_assignment_thanksgiving():
    assert_rows_match_tracked_schedule("Thursday")


def test_excel_row_assignment_saturday_tripleheader():
    assert_rows_match_tracked_schedule("Saturday")


@pytest.mark.parametrize(
    "filename",
    [
        "thanksgiving.html",
        "friday_game.html",
        "black_friday.html",
        "saturday_tripleheader.html",
        "christmas_tuesday.html",
        "christmas_wednesday.html",
    ],
)
def test_datetime_extraction(filename):
    games = load_fixture_games(filename)
    assert games["UTC_DateTime"].map(lambda value: isinstance(value, datetime)).all()

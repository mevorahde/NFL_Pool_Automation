import pytest
import pandas as pd
import os
import logging
from datetime import datetime
from bs4 import BeautifulSoup
from pool import (
    parse_game_card,
    get_week_number,
    apply_team_abbreviations,
    assign_excel_rows
)

def load_mock_html(filename):
    with open(f"tests/mock_html/{filename}", "r", encoding="utf-8") as f:
        return BeautifulSoup(f.read(), "html.parser")

@pytest.mark.parametrize("filename,expected_week", [
    ("thanksgiving.html", 12),
    ("friday_game.html", 13),
    ("black_friday.html", 13),
    ("christmas_tuesday.html", 16),
    ("christmas_wednesday.html", 16),
    ("saturday_tripleheader.html", 15),
])
def test_week_extraction(filename, expected_week):
    soup = load_mock_html(filename)
    week = get_week_number(soup)
    assert int(week) == expected_week

def test_row_assignment_thanksgiving():
    soup = load_mock_html("thanksgiving.html")
    cards = soup.find_all("div", class_="event-card")
    rows = [row for card in cards if (row := parse_game_card(card)) is not None]
    assert sum("Thursday" in row[-1] for row in rows) == 3

def test_row_assignment_saturday_tripleheader():
    soup = load_mock_html("saturday_tripleheader.html")
    cards = soup.find_all("div", class_="event-card")
    rows = [row for card in cards if (row := parse_game_card(card)) is not None]
    assert sum("Saturday" in row[-1] for row in rows) == 3

def test_abbreviation_mapping():
    soup = load_mock_html("christmas_tuesday.html")
    cards = soup.find_all("div", class_="event-card")
    rows = [row for card in cards if (row := parse_game_card(card)) is not None]
    df = pd.DataFrame(rows, columns=[
        "Team1", "Spread", "Team2", "Team1_Abbr", "Team2_Abbr",
        "Home_Team", "UTC_DateTime", "Favorite_Side"
    ])
    df = apply_team_abbreviations(df)
    assert all(df["Team1_Abbr"].notnull()) and all(df["Team2_Abbr"].notnull())

def test_black_friday_single_game():
    soup = load_mock_html("black_friday.html")
    cards = soup.find_all("div", class_="event-card")
    rows = [row for card in cards if (row := parse_game_card(card)) is not None]
    friday_games = [row for row in rows if "Friday" in row[-1]]
    assert len(friday_games) == 1

def test_excel_row_assignment_thanksgiving():
    soup = load_mock_html("thanksgiving.html")
    cards = soup.find_all("div", class_="event-card")
    rows = []
    for card in cards:
        row = parse_game_card(card)
        assert row is not None, "parse_game_card returned None"
        assert isinstance(row[6], datetime), f"Invalid datetime: {row[6]}"
        rows.append(row)

    df = pd.DataFrame(rows, columns=[
        "Team1", "Spread", "Team2", "Team1_Abbr", "Team2_Abbr",
        "Home_Team", "UTC_DateTime", "Favorite_Side"
    ])
    df = apply_team_abbreviations(df)
    df["Excel_Row"] = assign_excel_rows(df)

    assert os.path.exists("test_runs/test_schedule.xlsx"), "Missing test_schedule.xlsx"
    expected_df = pd.read_excel("test_runs/test_schedule.xlsx")

    for _, row in df.iterrows():
        match = expected_df[
            (expected_df["Team1"] == row.Team1_Abbr) &
            (expected_df["Team2"] == row.Team2_Abbr)
        ]
        assert not match.empty, f"No expected row for {row.Team1_Abbr} vs {row.Team2_Abbr}"
        expected_row = match.iloc[0]["Excel_Row"]
        assert row.Excel_Row == expected_row, (
            f"Row mismatch for {row.Team1_Abbr} vs {row.Team2_Abbr}: "
            f"expected {expected_row}, got {row.Excel_Row}"
        )

def test_excel_row_assignment_saturday_tripleheader():
    soup = load_mock_html("saturday_tripleheader.html")
    cards = soup.find_all("div", class_="event-card")
    rows = []
    for card in cards:
        row = parse_game_card(card)
        assert row is not None, "parse_game_card returned None"
        assert isinstance(row[6], datetime), f"Invalid datetime: {row[6]}"
        rows.append(row)

    df = pd.DataFrame(rows, columns=[
        "Team1", "Spread", "Team2", "Team1_Abbr", "Team2_Abbr",
        "Home_Team", "UTC_DateTime", "Favorite_Side"
    ])
    df = apply_team_abbreviations(df)
    df["Excel_Row"] = assign_excel_rows(df)

    assert os.path.exists("test_runs/test_schedule.xlsx"), "Missing test_schedule.xlsx"
    expected_df = pd.read_excel("test_runs/test_schedule.xlsx")

    for _, row in df.iterrows():
        match = expected_df[
            (expected_df["Team1"] == row.Team1_Abbr) &
            (expected_df["Team2"] == row.Team2_Abbr)
        ]
        assert not match.empty, f"No expected row for {row.Team1_Abbr} vs {row.Team2_Abbr}"
        expected_row = match.iloc[0]["Excel_Row"]
        assert row.Excel_Row == expected_row, (
            f"Row mismatch for {row.Team1_Abbr} vs {row.Team2_Abbr}: "
            f"expected {expected_row}, got {row.Excel_Row}"
        )

@pytest.mark.parametrize("filename", [
    "thanksgiving.html",
    "friday_game.html",
    "black_friday.html",
    "saturday_tripleheader.html",
    "christmas_tuesday.html",
    "christmas_wednesday.html",
])
def test_datetime_extraction(filename):
    soup = load_mock_html(filename)
    cards = soup.find_all("div", class_="event-card")
    for card in cards:
        row = parse_game_card(card)
        assert row is not None, f"parse_game_card returned None in {filename}"
        assert isinstance(row[6], datetime), f"Invalid datetime in {filename}: {row[6]}"
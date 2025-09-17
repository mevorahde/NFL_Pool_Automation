import pytest
from bs4 import BeautifulSoup
from pool import parse_game_card
from datetime import datetime

def load_mock_html(filename):
    with open(f"tests/mock_html/{filename}", "r", encoding="utf-8") as f:
        return BeautifulSoup(f.read(), "html.parser")

@pytest.mark.parametrize("filename", [
    "thanksgiving.html",
    "friday_game.html",
    "black_friday.html",
    "saturday_tripleheader.html",
    "christmas_tuesday.html",
    "christmas_wednesday.html",
])
def test_mock_html_structure(filename):
    soup = load_mock_html(filename)
    cards = soup.find_all("div", class_="event-card")
    assert cards, f"No event cards found in {filename}"

    for card in cards:
        row = parse_game_card(card)
        assert row is not None, f"parse_game_card returned None in {filename}"
        assert isinstance(row[6], datetime), f"Invalid datetime in {filename}: {row[6]}"
        assert all(row[i] for i in [0, 2, 3, 4]), f"Incomplete team info in {filename}: {row}"
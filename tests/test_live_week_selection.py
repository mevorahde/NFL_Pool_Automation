from bs4 import BeautifulSoup
import pandas as pd

import pool


BASE_URL = "https://www.scoresandodds.com/nfl"


def schedule_page(active_week, active_endpoint, games, next_week=None):
    menu_items = [
        f"""
        <li class="menu-item active">
          <span data-endpoint="{active_endpoint}">{active_week}</span>
        </li>
        """
    ]
    if next_week is not None:
        next_label, next_endpoint = next_week
        menu_items.append(
            f"""
            <li class="menu-item">
              <span data-endpoint="{next_endpoint}">{next_label}</span>
            </li>
            """
        )

    cards = []
    for away, away_abbr, home, home_abbr, kickoff in games:
        cards.append(
            f"""
            <div class="event-card">
              <span data-value="{kickoff}"></span>
              <table>
                <tr data-side="away">
                  <td><span class="team-name"><a data-abbr="{away_abbr}"><span>{away}</span></a></span></td>
                </tr>
                <tr data-side="home">
                  <td><span class="team-name"><a data-abbr="{home_abbr}"><span>{home}</span></a></span></td>
                </tr>
                <tr><td data-field="current-spread" data-side="home"><span class="data-value">-2.5</span></td></tr>
              </table>
            </div>
            """
        )

    return BeautifulSoup(
        f"""
        <div class="filters-week-picker">
          <div class="selector week-picker-week">
            <ul>{''.join(menu_items)}</ul>
          </div>
        </div>
        {''.join(cards)}
        """,
        "html.parser",
    )


def test_expired_hof_slate_advances_to_preseason_week_one(monkeypatch):
    hof_page = schedule_page(
        "HOF Game",
        "/nfl?week=2026-pre-0",
        [("Panthers", "CAR", "Cardinals", "ARI", "2026-08-07T00:00:00+00:00")],
        next_week=("Pre 1", "/nfl?week=2026-pre-1"),
    )
    preseason_page = schedule_page(
        "Pre 1",
        "/nfl?week=2026-pre-1",
        [("Packers", "GB", "Steelers", "PIT", "2026-08-13T23:00:00+00:00")],
    )
    requested_urls = []

    def fake_get_webpage(url, headers=None):
        requested_urls.append(url)
        return {
            BASE_URL: hof_page,
            f"{BASE_URL}?week=2026-pre-1": preseason_page,
        }[url]

    monkeypatch.setattr(pool, "get_webpage", fake_get_webpage)

    games, week = pool.scrape_nfl_data(now=pd.Timestamp("2026-08-08T12:00:00Z"))

    assert week == "Pre 1"
    assert len(games) == 1
    assert games.iloc[0][["Team1_Abbr", "Team2_Abbr"]].tolist() == ["GB", "PIT"]
    assert requested_urls == [BASE_URL, f"{BASE_URL}?week=2026-pre-1"]


def test_current_base_slate_does_not_make_an_extra_request(monkeypatch):
    preseason_page = schedule_page(
        "Pre 1",
        "/nfl?week=2026-pre-1",
        [("Packers", "GB", "Steelers", "PIT", "2026-08-13T23:00:00+00:00")],
        next_week=("Pre 2", "/nfl?week=2026-pre-2"),
    )
    requested_urls = []

    def fake_get_webpage(url, headers=None):
        requested_urls.append(url)
        return preseason_page

    monkeypatch.setattr(pool, "get_webpage", fake_get_webpage)

    games, week = pool.scrape_nfl_data(now=pd.Timestamp("2026-08-08T12:00:00Z"))

    assert week == "Pre 1"
    assert len(games) == 1
    assert requested_urls == [BASE_URL]


def test_expired_slate_without_a_next_week_aborts_before_excel(monkeypatch):
    hof_page = schedule_page(
        "HOF Game",
        "/nfl?week=2026-pre-0",
        [("Panthers", "CAR", "Cardinals", "ARI", "2026-08-07T00:00:00+00:00")],
    )
    monkeypatch.setattr(pool, "get_webpage", lambda url, headers=None: hof_page)
    monkeypatch.setattr(pool, "send_error_email", lambda *args, **kwargs: None)

    games, week = pool.scrape_nfl_data(now=pd.Timestamp("2026-08-08T12:00:00Z"))

    assert games is None
    assert week == "HOF Game"

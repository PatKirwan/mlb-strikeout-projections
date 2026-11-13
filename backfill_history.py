"""
backfill_history.py — one-time catch-up tool for MLB K Tracker

Fills history.csv for a date range you missed running run_today.py on.
For each pitcher who actually started a game in the range, this:
  1. Computes what your model WOULD have projected, using only game
     logs from before that start (no lookahead).
  2. Pulls the pitcher's ACTUAL strikeouts from that start.
  3. Leaves vegas_ou / my_bet blank — those were never captured
     automatically and can't be reconstructed after the fact.

Run with:  py backfill_history.py
Optional:  py backfill_history.py --start 2026-04-21 --end 2026-07-14

Safe to re-run: it skips (date, pitcher) pairs already in history.csv.
"""

import argparse
import csv
import datetime
import json
import time
import urllib.parse
import urllib.request
from pathlib import Path

MLB_BASE = "https://statsapi.mlb.com/api/v1"
HISTORY_FILE = Path(__file__).parent / "history.csv"
HIST_FIELDS = ['date', 'pitcher', 'team', 'opp', 'proj_ks', 'actual_ks',
               'vegas_ou', 'conf', 'my_bet', 'model_ok', 'odds']

# ── CLI args ────────────────────────────────────────────────────
parser = argparse.ArgumentParser()
parser.add_argument("--start", default="2026-04-21", help="First date to backfill (YYYY-MM-DD)")
parser.add_argument("--end", default=None, help="Last date to backfill (YYYY-MM-DD). Default: yesterday")
args = parser.parse_args()

START_DATE = args.start
END_DATE = args.end or (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")

print(f"⚾ Backfilling {START_DATE} → {END_DATE}")

# ── HTTP helper (same pattern as run_today.py) ──────────────────
def safe_get(url, retries=3, delay=2):
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                return json.loads(resp.read())
        except Exception:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                return None

# ── Step 1: pull every game + probable/starting pitcher in range ─
def get_games_in_range(start, end):
    """One schedule call for the whole range. Returns list of
    (date, pitcher_name, pitcher_id, team_abbrev, opp_abbrev)."""
    url = f"{MLB_BASE}/schedule?sportId=1&startDate={start}&endDate={end}&hydrate=probablePitcher,team"
    data = safe_get(url)
    if not data:
        return []

    out = []
    for date_entry in data.get("dates", []):
        game_date = date_entry.get("date", "")
        for game in date_entry.get("games", []):
            # Skip postponed/cancelled games — no real start happened
            if game.get("status", {}).get("detailedState", "") in ("Postponed", "Cancelled", "Suspended"):
                continue
            home = game.get("teams", {}).get("home", {})
            away = game.get("teams", {}).get("away", {})
            home_abbrev = home.get("team", {}).get("abbreviation", "")
            away_abbrev = away.get("team", {}).get("abbreviation", "")

            home_p = home.get("probablePitcher")
            if home_p and home_p.get("id"):
                out.append((game_date, home_p.get("fullName", ""), home_p["id"], home_abbrev, away_abbrev))
            away_p = away.get("probablePitcher")
            if away_p and away_p.get("id"):
                out.append((game_date, away_p.get("fullName", ""), away_p["id"], away_abbrev, home_abbrev))
    return out

# ── Step 2: full-season game log per pitcher (cached, one call each) ─
_gamelog_cache = {}

def get_full_gamelog(player_id, season):
    if player_id in _gamelog_cache:
        return _gamelog_cache[player_id]
    url = f"{MLB_BASE}/people/{player_id}/stats?stats=gameLog&group=pitching&season={season}&sportId=1"
    data = safe_get(url)
    starts = []
    if data:
        stats_list = data.get("stats", [])
        if stats_list:
            for split in stats_list[0].get("splits", []):
                stat = split.get("stat", {})
                ip_str = stat.get("inningsPitched", "0")
                ks = stat.get("strikeOuts", 0)
                date = split.get("date", "")
                try:
                    parts = str(ip_str).split(".")
                    ip = int(parts[0]) + (int(parts[1]) / 3 if len(parts) > 1 else 0)
                except Exception:
                    ip = 0
                if ip >= 2.0:
                    starts.append((date, int(ks), round(ip, 1)))
    starts.sort(key=lambda s: s[0])
    _gamelog_cache[player_id] = starts
    time.sleep(0.15)  # be polite to the API
    return starts

# ── Step 3: team K% (same as run_today.py — current full-season,
#            so this carries the same minor lookahead the live
#            script always has; impact is capped at ±5% in the model) ─
def get_team_kpct(season):
    url = f"{MLB_BASE}/teams/stats?season={season}&group=hitting&stats=season&sportId=1"
    data = safe_get(url)
    if not data:
        return {}
    teams_url = f"{MLB_BASE}/teams?season={season}&sportId=1"
    teams_data = safe_get(teams_url)
    id_to_abbrev = {}
    if teams_data:
        for t in teams_data.get("teams", []):
            id_to_abbrev[t["id"]] = t.get("abbreviation", "")
    result = {}
    for split in data.get("stats", [{}])[0].get("splits", []):
        team_id = split.get("team", {}).get("id")
        stat = split.get("stat", {})
        ks = stat.get("strikeOuts")
        pa = stat.get("plateAppearances")
        abbrev = id_to_abbrev.get(team_id, "")
        if abbrev and ks and pa:
            try:
                result[abbrev] = round(int(ks) / int(pa) * 100, 1)
            except Exception:
                pass
    return result

# ── Same projection model as run_today.py ────────────────────────
def project_ks(starts, ip_ceil=5.5, opp_kpct=None):
    if not starts:
        return None
    real = [s for s in starts if s[2] >= 2.0]
    if not real:
        return None
    k9_list = [round(s[1] / s[2] * 9, 2) for s in real if s[2] > 0]
    if not k9_list:
        return None
    weight_map = {1: 0.40, 2: 0.30, 3: 0.20, 4: 0.07, 5: 0.03}
    weighted_k9, total_w = 0.0, 0.0
    for i, k9 in enumerate(reversed(k9_list)):
        rank = i + 1
        w = weight_map.get(rank, 0.02)
        weighted_k9 += k9 * w
        total_w += w
    weighted_k9 /= total_w
    raw = weighted_k9 * ip_ceil / 9
    if opp_kpct and opp_kpct > 0:
        modifier = max(-0.05, min(0.05, (opp_kpct - 22.0) / 22.0))
        raw *= (1 + modifier)
    raw = round(raw, 1)
    avg_ks = sum(s[1] for s in real) / len(real)
    return min(raw, round(avg_ks * 2.0, 1))

# ── Load existing history to avoid duplicates ────────────────────
def load_existing():
    if not HISTORY_FILE.exists():
        return [], set()
    with open(HISTORY_FILE, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    keys = {(r.get("date", ""), r.get("pitcher", "").strip()) for r in rows}
    return rows, keys

def save_all(rows):
    for r in rows:
        for f2 in HIST_FIELDS:
            r.setdefault(f2, "")
    with open(HISTORY_FILE, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=HIST_FIELDS)
        w.writeheader()
        w.writerows(rows)

# ── Main ──────────────────────────────────────────────────────────
def main():
    season = datetime.date.fromisoformat(START_DATE).year
    existing_rows, existing_keys = load_existing()
    print(f"📜 Existing history: {len(existing_rows)} rows")

    print("🌐 Fetching schedule + probable pitchers for range (1 call)...")
    games = get_games_in_range(START_DATE, END_DATE)
    print(f"   ✅ Found {len(games)} pitcher-starts in range")

    print("🌐 Fetching team K% (current season)...")
    team_kpct = get_team_kpct(season)

    new_rows = []
    skipped = 0
    for i, (date, name, pid, team, opp) in enumerate(games):
        key = (date, name.strip())
        if key in existing_keys:
            skipped += 1
            continue

        full_log = get_full_gamelog(pid, season)

        # Actual result: the entry matching this exact date
        actual_entry = next((s for s in full_log if s[0] == date), None)
        if not actual_entry:
            continue  # pitcher didn't actually pitch that day (scratched, etc.)
        actual_ks = actual_entry[1]

        # Projection input: only starts strictly before this date
        prior_starts = [s for s in full_log if s[0] < date][-5:]
        opp_kpct = team_kpct.get(opp)
        proj = project_ks(prior_starts, opp_kpct=opp_kpct) if prior_starts else None

        if len(prior_starts) >= 3 and proj:
            conf = "HIGH"
        elif len(prior_starts) >= 1 and proj:
            conf = "MEDIUM"
        else:
            conf = "LOW"

        new_rows.append({
            "date": date, "pitcher": name, "team": team, "opp": opp,
            "proj_ks": proj if proj is not None else "",
            "actual_ks": actual_ks,
            "vegas_ou": "", "conf": conf, "my_bet": "", "model_ok": "", "odds": "",
        })
        existing_keys.add(key)

        if (i + 1) % 25 == 0:
            print(f"   ...processed {i + 1}/{len(games)}")

    all_rows = existing_rows + new_rows
    all_rows.sort(key=lambda r: (r["date"], r["pitcher"]))
    save_all(all_rows)

    print(f"✅ Added {len(new_rows)} rows | Skipped {skipped} already-logged | Total now {len(all_rows)}")
    print(f"💾 Saved to {HISTORY_FILE}")

if __name__ == "__main__":
    main()

"""
MLB K Projections — run_today.py v2
Fully automated: pulls pitcher game logs + team K% from MLB Stats API.
You only need to fill in: Vegas O/U, My Take, and your real bets.

Run with: py run_today.py
"""

import os, re, json, smtplib, datetime, math, time, argparse
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from dotenv import load_dotenv

load_dotenv()

# ── Parse --date argument ─────────────────────────────────────
_parser = argparse.ArgumentParser(add_help=False)
_parser.add_argument("--date", default=None,
    help="Run for a past date (YYYY-MM-DD). Builds slate only — no sync, no email.")
_parser.add_argument("--resync", action="store_true",
    help="Re-read all master Excel files in Output folder and rebuild history/bets from scratch.")
_args, _ = _parser.parse_known_args()
RESYNC_MODE = _args.resync

if _args.date:
    try:
        _run_date = datetime.date.fromisoformat(_args.date)
    except ValueError:
        print(f"❌ Invalid date format: {_args.date}. Use YYYY-MM-DD.")
        raise SystemExit(1)
    BACKFILL_MODE = True
else:
    _run_date = datetime.date.today()
    BACKFILL_MODE = False

EMAIL_ADDRESS  = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER    = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT      = int(os.getenv("SMTP_PORT", "587"))
HISTORY_FILE   = Path(__file__).parent / "history.csv"
BETS_FILE      = Path(__file__).parent / "bets.csv"
OUTPUT_DIR     = Path(__file__).parent / "Output"
OUTPUT_DIR.mkdir(exist_ok=True)

TODAY      = _run_date.strftime("%Y-%m-%d")
YESTERDAY  = (_run_date - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
DATE_LABEL = _run_date.strftime("%B %d, %Y").upper()
SEASON     = _run_date.year

def _normalize_date(d):
    """Normalize any date format to YYYY-MM-DD."""
    if not d: return ''
    import datetime as _dt
    d = str(d).strip()[:10]
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y'):
        try: return _dt.datetime.strptime(d, fmt).strftime('%Y-%m-%d')
        except: pass
    return d



# ── Design tokens ─────────────────────────────────────────────
BG     = "0B1320"; SURF1  = "111D2E"; SURF2  = "172234"; HEADER = "0D1A2A"
GREEN  = "00D26A"; RED    = "FF4757"; AMBER  = "FFA502"; TEAL   = "00C8C8"
YELLOW = "FFD32A"; WHITE  = "FFFFFF"; MUTED  = "7F8C9A"; DIM    = "3D5166"

# ── Helpers ───────────────────────────────────────────────────
def _font(bold=False, color=WHITE, size=9, italic=False):
    from openpyxl.styles import Font
    return Font(name="Calibri", bold=bold, color=color, size=size, italic=italic)

def _fill(color):
    from openpyxl.styles import PatternFill
    return PatternFill("solid", fgColor=color)

def _align(h="left", v="center"):
    from openpyxl.styles import Alignment
    return Alignment(horizontal=h, vertical=v)

def sc(ws, r, c, val="", bold=False, color=WHITE, bg=None,
       size=9, h="left", italic=False, fmt=None):
    from openpyxl.styles import Font, PatternFill, Alignment
    cl = ws.cell(row=r, column=c)
    if cl.__class__.__name__ == "MergedCell": return
    cl.value = val
    cl.font      = Font(name="Calibri", bold=bold, color=color, size=size, italic=italic)
    cl.alignment = Alignment(horizontal=h, vertical="center")
    if bg:  cl.fill = PatternFill("solid", fgColor=bg)
    if fmt: cl.number_format = fmt

def rb(r): return SURF1 if r % 2 == 0 else SURF2

def safe_get(url, retries=3, delay=2):
    """HTTP GET with retry logic."""
    import urllib.request, urllib.error
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                return json.loads(resp.read())
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                return None

# ── MLB Stats API ─────────────────────────────────────────────
MLB_BASE = "https://statsapi.mlb.com/api/v1"

def get_team_kpct():
    """
    Returns dict: {team_abbrev: K%} for current season.
    Uses MLB Stats API team hitting stats.
    """
    url = f"{MLB_BASE}/teams/stats?season={SEASON}&group=hitting&stats=season&sportId=1"
    data = safe_get(url)
    if not data: return {}

    # Also need team ID -> abbreviation mapping
    teams_url = f"{MLB_BASE}/teams?season={SEASON}&sportId=1"
    teams_data = safe_get(teams_url)
    id_to_abbrev = {}
    if teams_data:
        for t in teams_data.get("teams", []):
            id_to_abbrev[t["id"]] = t.get("abbreviation", "")

    result = {}
    for split in data.get("stats", [{}])[0].get("splits", []):
        team_id  = split.get("team", {}).get("id")
        stat     = split.get("stat", {})
        ks       = stat.get("strikeOuts")
        pa       = stat.get("plateAppearances")
        abbrev   = id_to_abbrev.get(team_id, "")
        if abbrev and ks and pa:
            try: result[abbrev] = round(int(ks) / int(pa) * 100, 1)
            except: pass
    return result

def search_player_id(name):
    """Find MLB player ID by name."""
    # Try search endpoint
    encoded = urllib.parse.quote(name)
    url = f"{MLB_BASE}/people/search?names={encoded}&sportId=1"
    data = safe_get(url)
    if data:
        for p in data.get("people", []):
            pos = p.get("primaryPosition", {}).get("abbreviation", "")
            if pos in ("P", "SP", "RP"):
                return p["id"]
    # Try sports people search
    url2 = f"{MLB_BASE}/sports/1/players?season={SEASON}&gameType=R"
    data2 = safe_get(url2)
    if data2:
        name_lower = name.lower()
        for p in data2.get("people", []):
            if p.get("fullName", "").lower() == name_lower:
                pos = p.get("primaryPosition", {}).get("abbreviation", "")
                if pos in ("P", "SP", "RP"):
                    return p["id"]
    return None

def get_pitcher_gamelogs(player_id, num_starts=5):
    """
    Returns list of last N starts: [(date, ks, ip), ...]
    Filters out outings < 2 IP (relief appearances).
    """
    import urllib.parse
    url = f"{MLB_BASE}/people/{player_id}/stats?stats=gameLog&group=pitching&season={SEASON}&sportId=1"
    data = safe_get(url)
    if not data: return []
    stats_list = data.get("stats", [])
    if not stats_list: return []

    starts = []
    for split in stats_list[0].get("splits", []):
        stat = split.get("stat", {})
        ip_str = stat.get("inningsPitched", "0")
        ks     = stat.get("strikeOuts", 0)
        date   = split.get("date", "")
        try:
            ip_parts = str(ip_str).split(".")
            ip = int(ip_parts[0]) + (int(ip_parts[1]) / 3 if len(ip_parts) > 1 else 0)
        except:
            ip = 0
        if ip >= 2.0:  # Only count real starts
            starts.append((date, int(ks), round(ip, 1)))

    return starts[-num_starts:]  # Most recent N starts

def get_probable_pitchers_today():
    """
    Returns today's probable pitchers from MLB Stats API schedule.
    Format: [(pitcher_name, pitcher_id, team_abbrev, opp_abbrev, game_time), ...]
    """
    import urllib.parse
    url = f"{MLB_BASE}/schedule?sportId=1&date={TODAY}&hydrate=probablePitcher,team"
    data = safe_get(url)
    if not data: return []

    pitchers = []
    for date_entry in data.get("dates", []):
        for game in date_entry.get("games", []):
            game_time_utc = game.get("gameDate", "")
            # Convert to local time label
            try:
                from datetime import timezone
                dt = datetime.datetime.fromisoformat(game_time_utc.replace("Z", "+00:00"))
                # Convert to CT (UTC-5 or UTC-6)
                ct_offset = -5  # CDT
                dt_ct = dt + datetime.timedelta(hours=ct_offset)
                time_label = dt_ct.strftime("%-I:%M %p")
            except:
                time_label = "TBD"

            home = game.get("teams", {}).get("home", {})
            away = game.get("teams", {}).get("away", {})
            home_abbrev = home.get("team", {}).get("abbreviation", "")
            away_abbrev = away.get("team", {}).get("abbreviation", "")

            # Home pitcher
            home_p = home.get("probablePitcher")
            if home_p:
                pitchers.append((
                    home_p.get("fullName", ""),
                    home_p.get("id"),
                    home_abbrev,
                    away_abbrev,
                    time_label
                ))
            # Away pitcher
            away_p = away.get("probablePitcher")
            if away_p:
                pitchers.append((
                    away_p.get("fullName", ""),
                    away_p.get("id"),
                    away_abbrev,
                    home_abbrev,
                    time_label
                ))

    return pitchers

# ── Projection engine ─────────────────────────────────────────
def project_ks(starts, ip_ceil=5.5, opp_kpct=None):
    """
    Projection model:
    1. Weighted K/9 from recent starts (L1=40%, L2=30%, L3=20%, L4=7%, L5=3%)
    2. Multiply by IP ceiling / 9 to get raw Ks
    3. Apply OPP K% modifier: league avg ~22%, scale ±15% max
    4. Cap at 2x historical average (anti-inflation)

    Args:
        starts: list of (date, ks, ip) tuples, oldest first
        ip_ceil: expected IP for this start
        opp_kpct: opponent team K% this season (e.g. 24.5 means 24.5%)
    """
    if not starts: return None
    real = [s for s in starts if s[2] >= 2.0]
    if not real: return None

    # K/9 for each real start
    k9_list = [round(s[1] / s[2] * 9, 2) for s in real if s[2] > 0]
    if not k9_list: return None

    # Recency weights — last start weighted most
    weight_map = {1: 0.40, 2: 0.30, 3: 0.20, 4: 0.07, 5: 0.03}
    n = len(k9_list)
    # k9_list[-1] = most recent, assign highest weight
    weighted_k9 = 0.0
    total_w = 0.0
    for i, k9 in enumerate(reversed(k9_list)):
        rank = i + 1  # 1=most recent
        w = weight_map.get(rank, 0.02)
        weighted_k9 += k9 * w
        total_w += w
    weighted_k9 = weighted_k9 / total_w  # normalize

    # Base projection
    raw = weighted_k9 * ip_ceil / 9

    # OPP K% modifier — league avg ~22%
    if opp_kpct and opp_kpct > 0:
        league_avg_kpct = 22.0
        modifier = (opp_kpct - league_avg_kpct) / league_avg_kpct
        modifier = max(-0.05, min(0.05, modifier))  # cap at ±5%
        raw = raw * (1 + modifier)

    raw = round(raw, 1)

    # Cap at 2x historical average
    avg_ks = sum(s[1] for s in real) / len(real)
    raw = min(raw, round(avg_ks * 2.0, 1))

    return raw

# ── Load / Save CSVs ──────────────────────────────────────────
def load_history():
    import csv
    if not HISTORY_FILE.exists(): return []
    with open(HISTORY_FILE, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r['date'] = _normalize_date(r.get('date',''))
    return rows

def save_history(rows):
    import csv
    if not rows: return
    HIST_FIELDS = ['date','pitcher','team','opp','proj_ks','actual_ks','vegas_ou','conf','my_bet','model_ok','odds']
    for r in rows:
        for f2 in HIST_FIELDS:
            r.setdefault(f2,'')
    with open(HISTORY_FILE, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=HIST_FIELDS)
        w.writeheader(); w.writerows(rows)

def load_bets():
    import csv
    if not BETS_FILE.exists(): return []
    with open(BETS_FILE, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r['date'] = _normalize_date(r.get('date',''))
    return rows

def save_bets(rows):
    import csv
    if not rows: return
    BETS_FIELDS = ['date','pitcher','team','opp','proj_ks','vegas','my_take','odds','actual_ks','correct','units','outcome','pl']
    for r in rows:
        for f2 in BETS_FIELDS:
            r.setdefault(f2,'')
    with open(BETS_FILE, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=BETS_FIELDS)
        w.writeheader(); w.writerows(rows)

# ── Sync yesterday's results from master Excel ────────────────

def scan_output_files(history, bets):
    """
    Silently scans every master Excel in Output folder BEFORE every run.
    Pulls Vegas O/U, My Take, Actual Ks, Proj Ks, and model grades.
    Excel always wins. Filters out placeholder values (N/A, —, -).
    Ensures any input you make in an Excel file is captured on the next run.
    """
    import glob, re as _re
    from openpyxl import load_workbook

    _junk = {None, "", "N/A", "—", "-", "--", "n/a", "None", "none"}

    def _is_real_vegas(v):
        if v in _junk: return False
        try: float(v); return True
        except: return False

    def _compute_mok(proj, actual, vegas):
        try:
            pj=float(proj); ac=float(actual); vg=float(vegas)
            mc="OVER" if pj>vg else "UNDER" if pj<vg else "PUSH"
            ao="OVER" if ac>vg else "UNDER" if ac<vg else "PUSH"
            if mc in ("OVER","UNDER") and ao in ("OVER","UNDER"):
                return "✅" if mc==ao else "❌"
            return "—"
        except: return ""

    files = sorted(glob.glob(str(OUTPUT_DIR / "MLB_Ks_Master_*.xlsx")))
    files = [f for f in files if "_BACKFILL" not in f]
    if not files:
        return history, bets

    existing_hist = {(r["date"], r["pitcher"]): r for r in history}
    existing_bets = {(b["date"], b["pitcher"]): b for b in bets}
    h_changed = 0; b_changed = 0

    for filepath in files:
        fname = Path(filepath).name
        m = _re.search(r"(\d{4}-\d{2}-\d{2})", fname)
        if not m: continue
        try:
            wb = load_workbook(filepath, data_only=True)
        except Exception as e:
            print(f"   ⚠️  Could not open {fname}: {e}")
            continue

        # ── Pull from History Log ──────────────────────────────
        hs = next((s for s in wb.sheetnames if "History" in s), None)
        if hs:
            ws = wb[hs]
            for r in range(4, 1000):
                p = ws.cell(r, 2).value
                if not p: break
                date   = _normalize_date(ws.cell(r, 1).value)
                proj   = ws.cell(r, 5).value
                actual = ws.cell(r, 6).value
                vegas  = ws.cell(r, 7).value
                conf   = ws.cell(r, 13).value
                my_bet = ws.cell(r, 14).value
                odds   = ws.cell(r, 16).value  # ODDS col (new)
                mok_xl = ws.cell(r, 12).value
                if not date: continue
                name = str(p).strip()

                # Clean my_bet — strip notes like "(bet)"
                mb = None
                if my_bet and str(my_bet).strip() not in _junk:
                    mb = str(my_bet).strip().split("(")[0].strip()
                    if mb in _junk: mb = None

                key = (date, name)
                if key in existing_hist:
                    rec = existing_hist[key]
                    changed = False
                    # Pull proj_ks if missing from CSV
                    if proj not in _junk and not rec.get("proj_ks"):
                        try: rec["proj_ks"] = str(float(proj)); changed = True
                        except: pass
                    # Actual Ks — Excel wins
                    if actual not in _junk:
                        if str(actual) != str(rec.get("actual_ks", "")): changed = True
                        rec["actual_ks"] = str(actual)
                    # Vegas — numeric only, Excel wins
                    if _is_real_vegas(vegas):
                        if str(vegas) != str(rec.get("vegas_ou", "")): changed = True
                        rec["vegas_ou"] = str(vegas)
                    # My Bet
                    if mb and mb != rec.get("my_bet", ""):
                        rec["my_bet"] = mb; changed = True
                    # Conf
                    if conf and str(conf).strip().upper() not in _junk:
                        rec["conf"] = str(conf).strip().upper()
                    if changed:
                        # Recompute model_ok
                        mok = _compute_mok(rec.get("proj_ks",""), rec.get("actual_ks",""), rec.get("vegas_ou",""))
                        # Fall back to Excel-calculated value if we can't compute
                        if mok:
                            rec["model_ok"] = mok
                        elif mok_xl and str(mok_xl) in ("✅","❌","—"):
                            rec["model_ok"] = str(mok_xl)
                        h_changed += 1
                else:
                    # Always add — track every slate appearance for My Take/Vegas
                    new_rec = {
                            "date":      date,
                            "pitcher":   name,
                            "team":      str(ws.cell(r,3).value or ""),
                            "opp":       str(ws.cell(r,4).value or ""),
                            "proj_ks":   str(proj) if proj not in _junk else "",
                            "actual_ks": str(actual) if actual not in _junk else "",
                            "vegas_ou":  str(vegas) if _is_real_vegas(vegas) else "",
                            "odds":      str(odds).strip() if odds not in (None,'','None') else "",
                            "conf":      str(conf or "").strip().upper(),
                            "my_bet":    mb or "",
                            "model_ok":  "",
                        }
                    mok = _compute_mok(new_rec["proj_ks"], new_rec["actual_ks"], new_rec["vegas_ou"])
                    new_rec["model_ok"] = mok if mok else (str(mok_xl) if mok_xl and str(mok_xl) in ("✅","❌","—") else "")
                    existing_hist[key] = new_rec
                    history.append(new_rec)
                    h_changed += 1

        # ── Pull from My Bets ──────────────────────────────────
        bs = next((s for s in wb.sheetnames if "Bets" in s and "My" in s), None)
        if bs:
            ws2 = wb[bs]
            for r in range(3, 500):
                p = ws2.cell(r, 2).value
                if not p or str(p) in ("PITCHER","TOTALS","ROI","WIN RATE","DATE"): continue
                dv = ws2.cell(r, 1).value
                ds = _normalize_date(dv) if dv else ""
                vegas   = ws2.cell(r, 6).value
                my_take = ws2.cell(r, 7).value
                if not ds or my_take in _junk: continue
                if not _is_real_vegas(vegas): continue
                key = (ds, str(p).strip())
                actual  = ws2.cell(r, 9).value
                correct = ws2.cell(r, 11).value
                units   = ws2.cell(r, 12).value
                odds    = ws2.cell(r, 8).value
                pl      = ws2.cell(r, 14).value
                if key in existing_bets:
                    rec = existing_bets[key]
                    changed = False
                    if actual not in _junk and str(actual) != str(rec.get("actual_ks","")):
                        rec["actual_ks"] = str(actual)
                        rec["correct"]   = str(correct or "")
                        rec["pl"]        = str(pl or "")
                        changed = True
                    if str(my_take) not in _junk and str(my_take) != str(rec.get("my_take","")):
                        rec["my_take"] = str(my_take); changed = True
                    if changed: b_changed += 1
                else:
                    try: u = abs(float(units or 0))
                    except: u = 0
                    existing_bets[key] = {
                        "date":      ds,
                        "pitcher":   str(p).strip(),
                        "team":      str(ws2.cell(r,3).value or ""),
                        "opp":       str(ws2.cell(r,4).value or ""),
                        "proj_ks":   ws2.cell(r,5).value or "",
                        "vegas":     str(vegas),
                        "my_take":   str(my_take),
                        "odds":      odds or -110,
                        "actual_ks": str(actual) if actual not in _junk else "",
                        "correct":   str(correct or ""),
                        "units":     u,
                        "outcome":   "",
                        "pl":        str(pl or ""),
                    }
                    b_changed += 1

    if h_changed or b_changed:
        save_history(list(existing_hist.values()))
        save_bets(list(existing_bets.values()))
        print(f"   📂 Pre-scan: {h_changed} history rows updated, {b_changed} bet rows updated")

    return list(existing_hist.values()), list(existing_bets.values())


def sync_yesterday(history, bets):
    """
    Reads yesterday's master Excel from Output folder.
    Pulls actual Ks from History Log and any new bets from My Bets.
    """
    from openpyxl import load_workbook
    import csv

    yesterday_file = OUTPUT_DIR / f"MLB_Ks_Master_{YESTERDAY}.xlsx"
    if not yesterday_file.exists():
        print(f"   ℹ️  No yesterday file found — skipping sync")
        return history, bets

    print(f"🔄 Syncing from {yesterday_file.name}...")
    wb = load_workbook(yesterday_file, data_only=True)

    # ── Sync actuals from History Log ──────────────────────────
    hist_sheet_name = next((s for s in wb.sheetnames if "History" in s), None)
    if hist_sheet_name:
        ws3 = wb[hist_sheet_name]
        existing = {(r["date"], r["pitcher"]): r for r in history}
        updated = 0
        for r in range(4, 500):
            p     = ws3.cell(r,2).value
            if not p: break
            date  = _normalize_date(ws3.cell(r,1).value)
            actual= ws3.cell(r,6).value
            vegas = ws3.cell(r,7).value
            my_bet= ws3.cell(r,14).value
            key   = (date, str(p).strip())

            if key in existing:
                rec = existing[key]
                changed = False
                if actual not in (None,""):
                    if str(actual) != str(rec.get("actual_ks","")): changed = True
                    rec["actual_ks"] = actual
                if vegas not in (None,""):
                    if str(vegas) != str(rec.get("vegas_ou","")): changed = True
                    rec["vegas_ou"] = vegas
                if my_bet not in (None,"") and not rec.get("my_bet"):
                    rec["my_bet"] = my_bet
                if changed:
                    # Recompute model_ok
                    try:
                        proj_v = float(rec.get("proj_ks") or 0)
                        act_v  = float(rec["actual_ks"])
                        veg_v  = float(rec["vegas_ou"])
                        if proj_v and veg_v:
                            mcall = "OVER" if proj_v>veg_v else "UNDER" if proj_v<veg_v else "PUSH"
                            aou   = "OVER" if act_v>veg_v  else "UNDER" if act_v<veg_v  else "PUSH"
                            rec["model_ok"] = "✅" if mcall==aou else "❌" if mcall in ("OVER","UNDER") and aou in ("OVER","UNDER") else "—"
                    except: pass
                    updated += 1
            else:
                if actual not in (None,"") and date:
                    new_rec = {
                        "date":date, "pitcher":str(p).strip(),
                        "team":str(ws3.cell(r,3).value or ""),
                        "opp":str(ws3.cell(r,4).value or ""),
                        "proj_ks":ws3.cell(r,5).value or "",
                        "actual_ks":actual,
                        "vegas_ou":vegas or "",
                        "conf":str(ws3.cell(r,13).value or ""),
                        "my_bet":str(my_bet or ""),
                        "model_ok":"",
                    }
                    try:
                        proj_v=float(new_rec["proj_ks"] or 0)
                        act_v =float(actual)
                        veg_v =float(vegas or 0)
                        if proj_v and veg_v:
                            mcall="OVER" if proj_v>veg_v else "UNDER" if proj_v<veg_v else "PUSH"
                            aou  ="OVER" if act_v>veg_v  else "UNDER" if act_v<veg_v  else "PUSH"
                            new_rec["model_ok"]="✅" if mcall==aou else "❌" if mcall in ("OVER","UNDER") and aou in ("OVER","UNDER") else "—"
                    except: pass
                    history.append(new_rec)
                    existing[key] = new_rec
                    updated += 1

        if updated:
            save_history(list(existing.values()))
            print(f"   ✅ Synced {updated} results")
        else:
            print(f"   ℹ️  History already up to date")

    # ── Sync bets from My Bets tab ─────────────────────────────
    bets_sheet_name = next((s for s in wb.sheetnames if "Bets" in s), None)
    if bets_sheet_name:
        ws2 = wb[bets_sheet_name]
        existing_bets = {(b["date"], b["pitcher"]): b for b in bets}
        new_bets = []
        added = 0
        for r in range(3, 200):
            p = ws2.cell(r,2).value
            if not p or str(p) in ("PITCHER","TOTALS","ROI","WIN RATE"): continue
            date_v  = ws2.cell(r,1).value
            date_s  = _normalize_date(date_v) if date_v else ""
            vegas   = ws2.cell(r,6).value
            my_take = ws2.cell(r,7).value
            if not my_take or not vegas: continue
            key = (date_s, str(p).strip())
            actual  = ws2.cell(r,9).value
            correct = ws2.cell(r,11).value
            units   = ws2.cell(r,12).value
            odds    = ws2.cell(r,8).value
            pl      = ws2.cell(r,14).value

            if key in existing_bets:
                rec = existing_bets[key]
                if actual not in (None,"") and str(rec.get("actual_ks","")) in ("","None"):
                    rec["actual_ks"] = actual
                    rec["correct"]   = correct or ""
                    rec["pl"]        = pl or ""
                    added += 1
                new_bets.append(rec)
            else:
                new_bets.append({
                    "date":date_s, "pitcher":str(p).strip(),
                    "team":str(ws2.cell(r,3).value or ""),
                    "opp":str(ws2.cell(r,4).value or ""),
                    "proj_ks":ws2.cell(r,5).value or "",
                    "vegas":vegas, "my_take":str(my_take).strip(),
                    "odds":odds or -110,
                    "actual_ks":actual or "", "correct":str(correct or ""),
                    "units":units or "", "pl":pl or "",
                })
                added += 1

        if new_bets:
            seen = set(); deduped = []
            for b in new_bets:
                k=(b["date"],b["pitcher"])
                if k not in seen: seen.add(k); deduped.append(b)
            save_bets(deduped); bets=deduped
            if added: print(f"   ✅ Synced {added} bet rows")

    return history, bets

# ── Build master Excel ────────────────────────────────────────

def _bucket_ev(signal_label, history, per_units=10.0):
    """
    Compute EV per $per_units for a given signal bucket.
    Uses actual logged odds from history where available, else -110.
    Returns (ev, n_starts, n_with_odds, low_sample)
    """
    if signal_label is None:
        rows = [r for r in history if r.get('model_ok') in ('✅','❌')]
    else:
        rows = [r for r in history
                if r.get('model_ok') in ('✅','❌')
                and r.get('model_signal','') == signal_label]
    n = len(rows)
    if n == 0:
        return None, 0, 0, True
    wins   = sum(1 for r in rows if r.get('model_ok') == '✅')
    win_rt = wins / n
    # Collect odds for EV calc
    odds_vals = []
    for r in rows:
        try: odds_vals.append(float(r.get('odds') or -110))
        except: odds_vals.append(-110.0)
    avg_odds = sum(odds_vals) / len(odds_vals)
    # Win/loss amounts per $per_units
    if avg_odds > 0:
        win_amt  = per_units * avg_odds / 100
        risk_amt = per_units
    else:
        win_amt  = per_units * 100 / abs(avg_odds)
        risk_amt = per_units
    ev = round(win_rt * win_amt - (1 - win_rt) * risk_amt, 2)
    n_with_odds = sum(1 for r in rows if r.get('odds'))
    low_sample  = n < 30
    return ev, n, n_with_odds, low_sample


def _odds_display(odds, per=10.0):
    """Format odds as 'X to win $10' or '$10 to win X'."""
    try:
        o = float(odds)
        if o > 0:
            win = round(per * o / 100, 2)
            return f"${per:.0f} to win ${win:.2f}"
        else:
            risk = round(per * 100 / abs(o), 2)
            return f"${risk:.2f} to win ${per:.0f}"
    except:
        return "—"


def _signal_for_edge(proj, vegas):
    """Return signal bucket label for a proj/vegas pair."""
    if proj is None or vegas is None:
        return "—"
    e = round(proj - vegas, 2)
    if abs(e) <= 0.5: return "PUSH"
    if e > 0:
        if e <= 1.0: return "OVER 0.6-1.0"
        if e <= 1.5: return "OVER 1.1-1.5"
        if e <= 2.0: return "OVER 1.6-2.0"
        if e <= 3.0: return "OVER 2.1-3.0"
        return "OVER 3.1+"
    else:
        ae = abs(e)
        if ae <= 1.0: return "UNDER 0.6-1.0"
        if ae <= 1.5: return "UNDER 1.1-1.5"
        if ae <= 2.0: return "UNDER 1.6-2.0"
        if ae <= 3.0: return "UNDER 2.1-3.0"
        return "UNDER 3.1+"

def build_master(pitchers, history, bets):
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    from collections import defaultdict

    wb = Workbook()
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    # ────────────────────────────────────────────────────────────
    # SHEET 1: Today's Slate
    # Cols: PITCHER|TEAM|OPP|PROJ Ks|AVG Ks|IP CEIL|VEGAS O/U|EDGE|OPP K%|PLATOON|CONF|MY TAKE|L1|L2|L3|L4|L5
    # ────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "📅 Today's Slate"
    NCOLS = 20
    for i, w in enumerate([22,7,7,9,8,8,10,7,14,8,9,9,11,8,12,5,5,5,5,5], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells(f"A1:T1")
    sc(ws,1,1,f"⚾  MLB STRIKEOUT PROJECTIONS  —  {DATE_LABEL}",
       bold=True,color=GREEN,bg=BG,size=12,h="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(f"A2:T2")
    sc(ws,2,1,"✏️  Yellow = your inputs  |  PROJ = weighted K/9 (recency + OPP K%) × IP ceiling  |  EDGE auto-calculates when you type Vegas",
       italic=True,color=MUTED,bg=SURF1,size=8,h="center")
    ws.row_dimensions[2].height = 13

    hdrs = ["PITCHER","TEAM","OPP","PROJ Ks","AVG Ks","IP CEIL","VEGAS O/U","EDGE","SIGNAL",
            "OPP K%","PLATOON","CONF","MY TAKE","ODDS","EV","L1","L2","L3","L4","L5"]
    for i,h in enumerate(hdrs,1):
        sc(ws,3,i,h,bold=True,color=MUTED,bg=HEADER,size=8,h="center")
    ws.row_dimensions[3].height = 17

    conf_color = {"HIGH":GREEN,"MEDIUM":AMBER,"LOW":RED}

    for ri, p in enumerate(pitchers, 4):
        bg = rb(ri)
        proj   = p.get("proj")
        avg    = p.get("avg_ks")
        vegas  = p.get("vegas")
        edge   = round(float(proj)-float(vegas),1) if proj and vegas else None
        ec     = GREEN if isinstance(edge,float) and edge>0 else RED if isinstance(edge,float) and edge<0 else MUTED
        opp_k  = p.get("opp_kpct")
        opp_k_str = f"{opp_k:.1f}%" if opp_k else "—"
        last5  = p.get("starts",[])
        conf   = p.get("conf","MEDIUM")

        sc(ws,ri,1,p["name"],      color=WHITE, bg=bg,h="left")
        sc(ws,ri,2,p["team"],      color=MUTED, bg=bg,h="center")
        sc(ws,ri,3,p["opp"],       color=MUTED, bg=bg,h="center")
        # Data flags
        n_starts = len([s for s in p.get("starts",[]) if s[2] >= 2.0])
        proj_flag = ""
        if proj is None or n_starts == 0:
            proj_display = "DEBUT"
            proj_color   = AMBER
            conf = "DEBUT"
        elif n_starts == 1:
            proj_display = f"{proj}*"   # only 1 start — low sample
            proj_color   = WHITE
        else:
            proj_display = str(proj)
            proj_color   = WHITE
        # Vegas flag
        vegas_display = vegas if vegas else "N/A"
        vegas_color   = YELLOW if vegas else MUTED

        sc(ws,ri,4,proj_display if proj is None or n_starts<=1 else proj,
           color=proj_color, bg=bg,bold=True,h="center")
        sc(ws,ri,5,avg if avg else "—", color=TEAL,bg=bg,h="center")
        sc(ws,ri,6,p.get("ip_ceil",5.5), color=DIM,bg=bg,h="center")
        sc(ws,ri,7,vegas_display, color=vegas_color,bold=bool(vegas),h="center")
        # EDGE: live formula so it auto-updates as you type Vegas O/U
        edge_cell = ws.cell(row=ri, column=8)
        edge_cell.value = f'=IF(AND(D{ri}<>"",G{ri}<>""),D{ri}-G{ri},"")'
        edge_cell.font = _font(bold=True, color=TEAL, size=9)
        edge_cell.alignment = _align(h="center")
        edge_cell.fill = _fill(bg)
        edge_cell.number_format = "0.0" 
        opp_k_color = MUTED if opp_k else DIM
        # SIGNAL col 9: live bucket label
        _sc = ws.cell(row=ri, column=9)
        _sc.value = f'=IF(OR(D{ri}="",G{ri}="",G{ri}="N/A"),"—",IF(ABS(D{ri}-G{ri})<=0.5,"PUSH",IF(D{ri}-G{ri}>0,IF(D{ri}-G{ri}<=1.0,"OVER 0.6-1.0",IF(D{ri}-G{ri}<=1.5,"OVER 1.1-1.5",IF(D{ri}-G{ri}<=2.0,"OVER 1.6-2.0",IF(D{ri}-G{ri}<=3.0,"OVER 2.1-3.0","OVER 3.1+")))),IF(G{ri}-D{ri}<=1.0,"UNDER 0.6-1.0",IF(G{ri}-D{ri}<=1.5,"UNDER 1.1-1.5",IF(G{ri}-D{ri}<=2.0,"UNDER 1.6-2.0",IF(G{ri}-D{ri}<=3.0,"UNDER 2.1-3.0","UNDER 3.1+")))))))'
        _sc.font = _font(bold=True, color=TEAL, size=8)
        _sc.alignment = _align(h="center")
        _sc.fill = _fill(bg)
        sc(ws,ri,10,opp_k_str, color=opp_k_color,bg=bg,h="center",size=8)
        sc(ws,ri,11,"—",           color=DIM,   bg=bg,h="center",size=8)
        conf_display_color = AMBER if conf == "DEBUT" else conf_color.get(conf, WHITE)
        sc(ws,ri,12,conf, color=conf_display_color,bg=bg,bold=True,h="center")
        sc(ws,ri,13,"",            color=YELLOW,bg=bg,h="center")   # MY TAKE
        sc(ws,ri,14,"",            color=YELLOW,bg=bg,h="center")   # ODDS (user input)

        # EV: auto-calc from bucket hit rate × odds, per $10
        # Uses _bucket_ev helper; asterisk if <30 starts in bucket
        signal_now = _signal_for_edge(proj if proj and not isinstance(proj,str) else None,
                                       p.get("vegas"))
        ev_val, ev_n, _, ev_low = _bucket_ev(signal_now, history) if signal_now not in ("—","PUSH","") else (None,0,0,True)
        if ev_val is not None:
            # EV formula cell: =IF(AND(N{ri}<>"",G{ri}<>""), EV_from_odds_and_hitrate, "—")
            # Since hit rate comes from Python, embed it directly
            # Format: show per-$10 in odds-display style
            ev_str = f"+${ev_val:.2f}" if ev_val >= 0 else f"-${abs(ev_val):.2f}"
            if ev_low: ev_str += "*"
            ev_color = GREEN if ev_val > 0 else RED if ev_val < 0 else MUTED
        else:
            ev_str = "—"
            ev_color = MUTED
        sc(ws,ri,15,ev_str, color=ev_color,bold=ev_val is not None and not ev_low,bg=bg,h="center",size=8)

        ks_list = [s[1] for s in reversed(last5[-5:])]  # most recent first
        for li, k in enumerate(ks_list, 16):
            sc(ws,ri,li,k,color=TEAL,bg=bg,h="center",size=8)
        for li in range(16+len(ks_list), 21):
            sc(ws,ri,li,"—",color=DIM,bg=bg,h="center",size=8)
        ws.row_dimensions[ri].height = 15

    fr = 4 + len(pitchers)
    ws.merge_cells(f"A{fr}:T{fr}")
    sc(ws,fr,1,"OPP K% = opponent team strikeout rate this season  |  AVG Ks = avg per qualifying start (≥2 IP)  |  L1 = most recent start",
       italic=True,color=DIM,bg=BG,size=8,h="center")
    ws.row_dimensions[fr].height = 13
    # EV key box — explains * notation
    key_r = fr + 1
    ws.merge_cells(f"A{key_r}:O{key_r}")
    sc(ws,key_r,1,
       "📌  EV KEY:  Shows expected value per $10 bet using this model\'s historical win rate for that edge bucket  |  * = fewer than 30 graded starts in this bucket — treat as estimate",
       italic=True,color=AMBER,bg=HEADER,size=8,h="left")
    ws.row_dimensions[key_r].height=13
    ws.freeze_panes = "A4"

    # ────────────────────────────────────────────────────────────
    # SHEET 2: My Bets
    # Cols: DATE|PITCHER|TEAM|OPP|PROJ Ks|VEGAS O/U|MY TAKE|ODDS|ACTUAL Ks|RESULT|CORRECT?|BET UNITS|WIN AMT|TRUE P/L|RUNNING P/L
    # ────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("📈 My Bets")
    for i, w in enumerate([12,22,7,7,9,9,10,8,10,9,10,11,10,10,12], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    ws2.merge_cells("A1:O1")
    sc(ws2,1,1,"📈  MY BETS  —  All-time personal bet log",bold=True,color=GREEN,bg=BG,size=11,h="center")
    ws2.row_dimensions[1].height = 24

    for i,h in enumerate(["DATE","PITCHER","TEAM","OPP","PROJ Ks","VEGAS O/U","MY TAKE","ODDS",
                           "ACTUAL Ks","RESULT","CORRECT?","BET UNITS","WIN AMT","TRUE P/L","RUNNING P/L"],1):
        sc(ws2,2,i,h,bold=True,color=MUTED,bg=HEADER,size=8,h="center")
    ws2.row_dimensions[2].height = 17

    data_row = 3
    for b in bets:
        bg  = rb(data_row)
        ri  = data_row
        try: odds  = float(b.get("odds") or -110)
        except: odds = -110
        try: units = abs(float(b.get("units") or 0))
        except: units = 0

        # Static data columns
        sc(ws2,ri,1,  b.get("date",""),      color=MUTED,  bg=bg, h="center")
        sc(ws2,ri,2,  b.get("pitcher",""),   color=WHITE,  bg=bg, h="left")
        sc(ws2,ri,3,  b.get("team",""),      color=MUTED,  bg=bg, h="center")
        sc(ws2,ri,4,  b.get("opp",""),       color=MUTED,  bg=bg, h="center")
        try: sc(ws2,ri,5, float(b.get("proj_ks") or 0), color=WHITE, bg=bg, h="center")
        except: sc(ws2,ri,5, b.get("proj_ks",""), color=WHITE, bg=bg, h="center")
        try: sc(ws2,ri,6, float(b.get("vegas") or 0), color=WHITE, bg=bg, h="center")
        except: sc(ws2,ri,6, b.get("vegas",""), color=WHITE, bg=bg, h="center")
        sc(ws2,ri,7,  b.get("my_take",""),   color=YELLOW, bg=bg, h="center")
        sc(ws2,ri,8,  odds,                  color=YELLOW, bg=bg, h="center")
        sc(ws2,ri,9,  b.get("actual_ks",""), color=YELLOW, bg=bg, h="center")
        sc(ws2,ri,12, units or "",           color=WHITE,  bg=bg, h="center")

        # Live formulas for calculated columns — same as blank rows
        # Col 10: RESULT = actual vs vegas
        # Col 11: CORRECT? = MY TAKE matches RESULT
        # Col 13: WIN AMT = units * odds payout
        # Col 14: TRUE P/L = win or lose
        # Col 15: RUNNING P/L = cumulative sum
        for ci, formula, fcol, fbold in [
            (10, f'=IF(AND(I{ri}<>"",F{ri}<>""),IF(I{ri}>F{ri},"OVER",IF(I{ri}<F{ri},"UNDER","PUSH")),"")', WHITE,  False),
            (11, f'=IF(AND(G{ri}<>"",J{ri}<>""),IF(OR(AND(UPPER(G{ri})="OVER",J{ri}="OVER"),AND(UPPER(G{ri})="UNDER",J{ri}="UNDER")),"✅","❌"),"")', WHITE, True),
            (13, f'=IF(AND(L{ri}<>"",H{ri}<>""),IF(H{ri}>0,ABS(L{ri})*H{ri}/100,ABS(L{ri})*100/ABS(H{ri})),"")', TEAL,  False),
            (14, f'=IF(K{ri}="✅",M{ri},IF(K{ri}="❌",-ABS(L{ri}),\"\"))', GREEN,  True),
            (15, f'=IF(N{ri}="","",SUM($N$3:N{ri}))', TEAL,  True),
        ]:
            sc(ws2, ri, ci, formula, color=fcol, bg=bg, h="center", bold=fbold)
        ws2.row_dimensions[ri].height = 15
        data_row += 1

    # Blank formula rows
    for ri in range(data_row, data_row+50):
        bg=rb(ri)
        formulas={
            10:f'=IF(AND(I{ri}<>"",F{ri}<>""),IF(I{ri}>F{ri},"OVER",IF(I{ri}<F{ri},"UNDER","PUSH")),"")',
            11:f'=IF(AND(G{ri}<>"",J{ri}<>""),IF(OR(AND(UPPER(G{ri})="OVER",J{ri}="OVER"),AND(UPPER(G{ri})="UNDER",J{ri}="UNDER")),"✅","❌"),"")',
            13:f'=IF(AND(L{ri}<>"",H{ri}<>""),IF(H{ri}>0,ABS(L{ri})*H{ri}/100,ABS(L{ri})*100/ABS(H{ri})),"")',
            14:f'=IF(K{ri}="✅",M{ri},IF(K{ri}="❌",-ABS(L{ri}),""))',
            15:f'=IF(N{ri}="","",SUM($N$3:N{ri}))',
        }
        for ci in range(1,16):
            if ci in formulas:
                col=TEAL if ci in [13,15] else GREEN if ci==14 else WHITE
                sc(ws2,ri,ci,formulas[ci],color=col,bg=bg,bold=ci in [11,14,15],h="center")
            else:
                sc(ws2,ri,ci,"",bg=bg)
        ws2.row_dimensions[ri].height=15

    # Totals
    tr=data_row+50
    ws2.merge_cells(f"A{tr}:I{tr}")
    sc(ws2,tr,1,"TOTALS",bold=True,color=AMBER,bg=HEADER,h="right")
    sc(ws2,tr,12,f"=SUM(L3:L{tr-1})",bold=True,color=WHITE,bg=HEADER,h="center")
    sc(ws2,tr,13,f"=SUM(M3:M{tr-1})",bold=True,color=TEAL,bg=HEADER,h="center")
    sc(ws2,tr,14,f"=SUM(N3:N{tr-1})",bold=True,color=GREEN,bg=HEADER,h="center")
    roi_r=tr+1
    ws2.merge_cells(f"A{roi_r}:I{roi_r}")
    sc(ws2,roi_r,1,"ROI",bold=True,color=AMBER,bg=HEADER,h="right")
    sc(ws2,roi_r,14,f'=IFERROR(N{tr}/SUMIF(L3:L{tr-1},">0"),"—")',
       bold=True,color=TEAL,bg=HEADER,h="center",fmt="0.0%")
    wr_r=tr+2
    ws2.merge_cells(f"A{wr_r}:I{wr_r}")
    sc(ws2,wr_r,1,"WIN RATE",bold=True,color=AMBER,bg=HEADER,h="right")
    sc(ws2,wr_r,14,f'=IFERROR(COUNTIF(K3:K{tr-1},"✅")/SUMPRODUCT((K3:K{tr-1}="✅")+(K3:K{tr-1}="❌")),"—")',
       bold=True,color=GREEN,bg=HEADER,h="center",fmt="0.0%")
    ws2.freeze_panes="A3"

    # ────────────────────────────────────────────────────────────
    # SHEET 3: History Log
    # ────────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("📜 History Log")
    for i,w in enumerate([12,22,7,7,9,10,10,8,10,12,11,10,9,14,14,8,11],1):
        ws3.column_dimensions[get_column_letter(i)].width=w

    ws3.merge_cells("A1:Q1")
    sc(ws3,1,1,"📜  MASTER HISTORY LOG  —  Auto-updated each morning from your email",
       bold=True,color=GREEN,bg=BG,size=11,h="center")
    ws3.row_dimensions[1].height=24
    ws3.merge_cells("A2:Q2")
    sc(ws3,2,1,"After games: fill ACTUAL Ks + VEGAS O/U only — all grades calculate automatically",
       italic=True,color=MUTED,bg=SURF1,size=8,h="center")
    ws3.row_dimensions[2].height=13

    for i,h in enumerate(["DATE","PITCHER","TEAM","OPP","PROJ Ks","ACTUAL Ks","VEGAS O/U",
                           "ERROR","WITHIN 0.5?","MODEL CALL","ACTUAL O/U","MODEL ✓?","CONF","MY BET","MODEL SIGNAL","ODDS","EV"],1):
        sc(ws3,3,i,h,bold=True,color=MUTED,bg=HEADER,size=8,h="center")
    ws3.row_dimensions[3].height=17

    log_row=4
    for h in history:
        bg=rb(log_row)
        proj   =_safe_float(h.get("proj_ks"))
        actual =_safe_float(h.get("actual_ks"))
        vegas  =_safe_float(h.get("vegas_ou"))
        conf   =str(h.get("conf","")).upper()
        my_bet =h.get("my_bet","")
        error  =round(actual-proj,1) if actual is not None and proj else None
        within =("✅" if abs(error)<=0.5 else "❌") if error is not None else ""
        # MODEL CALL uses same ±0.5 neutral zone as MODEL SIGNAL
        mcall  =("PUSH" if abs(proj-vegas)<=0.5 else "OVER" if proj>vegas else "UNDER") if proj and vegas else ""
        aou    =("OVER" if actual>vegas else "UNDER" if actual<vegas else "PUSH") if actual is not None and vegas else ""
        # MODEL ✓?: grade OVER/UNDER calls; show actual result for PUSH (informational)
        if mcall in ("OVER","UNDER") and aou in ("OVER","UNDER"):
            mok = "✅" if mcall==aou else "❌"
        elif mcall == "PUSH" and aou in ("OVER","UNDER"):
            mok = aou  # show what actually happened (OVER/UNDER)
        else:
            mok = "—"
        err_c  =GREEN if error and error>0 else RED if error and error<0 else MUTED
        mok_c  = GREEN if mok=="✅" else RED if mok=="❌" else TEAL if mok=="OVER" else AMBER if mok=="UNDER" else MUTED
        conf_c =GREEN if conf=="HIGH" else AMBER if conf=="MEDIUM" else RED if conf=="LOW" else MUTED

        sc(ws3,log_row,1,h.get("date",""),  color=MUTED,bg=bg,h="center")
        sc(ws3,log_row,2,h.get("pitcher",""),color=WHITE,bg=bg,h="left")
        sc(ws3,log_row,3,h.get("team",""),  color=MUTED,bg=bg,h="center")
        sc(ws3,log_row,4,h.get("opp",""),   color=MUTED,bg=bg,h="center")
        sc(ws3,log_row,5,proj,              color=WHITE,bg=bg,h="center")
        sc(ws3,log_row,6,actual,            color=YELLOW,bg=bg,h="center")
        sc(ws3,log_row,7,vegas,             color=YELLOW,bg=bg,h="center")
        # Compute signal bucket from proj/vegas
        def _sig(p, v):
            if p is None or v is None: return "—"
            e = round(p - v, 2)
            if abs(e) <= 0.5: return "PUSH"
            if e > 0:
                if e <= 1.0: return "OVER 0.6-1.0"
                if e <= 1.5: return "OVER 1.1-1.5"
                if e <= 2.0: return "OVER 1.6-2.0"
                if e <= 3.0: return "OVER 2.1-3.0"
                return "OVER 3.1+"
            else:
                ae = abs(e)
                if ae <= 1.0: return "UNDER 0.6-1.0"
                if ae <= 1.5: return "UNDER 1.1-1.5"
                if ae <= 2.0: return "UNDER 1.6-2.0"
                if ae <= 3.0: return "UNDER 2.1-3.0"
                return "UNDER 3.1+"
        signal = _sig(proj, vegas)
        h['model_signal'] = signal  # store for EV lookup
        mok_c  = GREEN if mok=="✅" else RED if mok=="❌" else MUTED
        sig_c  = GREEN if "OVER" in signal else RED if "UNDER" in signal else MUTED
        # Live Excel formulas — recalculate when user types actual/vegas
        r = log_row
        for fc, formula, fcol, fbold in [
            (8,  f'=IF(AND(F{r}<>"",E{r}<>""),F{r}-E{r},"")',                                           AMBER, False),
            (9,  f'=IF(H{r}="","",IF(ABS(H{r})<=0.5,"✅","❌"))',                                       WHITE, False),
            (10, f'=IF(AND(E{r}<>"",G{r}<>""),IF(ABS(E{r}-G{r})<=0.5,"PUSH",IF(E{r}>G{r},"OVER","UNDER")),"N/A")', WHITE, False),
            (11, f'=IF(AND(F{r}<>"",G{r}<>""),IF(F{r}>G{r},"OVER",IF(F{r}<G{r},"UNDER","PUSH")),"")',   WHITE, False),
            (12, f'=IF(AND(J{r}<>"",K{r}<>"",J{r}<>"PUSH",J{r}<>"N/A"),IF(J{r}=K{r},"✅",IF(K{r}="PUSH","—","❌")),"—")', WHITE, True),
            (15, f'=IF(OR(E{r}="",G{r}="",G{r}="N/A"),"—",IF(ABS(E{r}-G{r})<=0.5,"PUSH",IF(E{r}-G{r}>0,IF(E{r}-G{r}<=1.0,"OVER 0.6-1.0",IF(E{r}-G{r}<=1.5,"OVER 1.1-1.5",IF(E{r}-G{r}<=2.0,"OVER 1.6-2.0",IF(E{r}-G{r}<=3.0,"OVER 2.1-3.0","OVER 3.1+")))),IF(G{r}-E{r}<=1.0,"UNDER 0.6-1.0",IF(G{r}-E{r}<=1.5,"UNDER 1.1-1.5",IF(G{r}-E{r}<=2.0,"UNDER 1.6-2.0",IF(G{r}-E{r}<=3.0,"UNDER 2.1-3.0","UNDER 3.1+")))))))', TEAL, True),
        ]:
            cl = ws3.cell(r, fc)
            cl.value = formula
            cl.font = _font(bold=fbold, color=fcol, size=9)
            cl.alignment = _align(h="center")
            cl.fill = _fill(bg)
        # Store signal for EV (computed from Python value for this build)
        h['model_signal'] = signal
        sc(ws3,log_row,13, conf,   color=conf_c, bold=True, bg=bg, h="center")
        # ODDS (col16) — from history dict (populated by daily scan from slate)
        odds_v = h.get("odds","")
        try: odds_f = float(odds_v)
        except: odds_f = None
        odds_str = _odds_display(odds_f) if odds_f else ""
        sc(ws3,log_row,16, odds_str, color=TEAL if odds_str else MUTED, bg=bg, h="center", size=8)
        # EV (col17) — computed from signal bucket hit rate × logged odds (default -110 if none)
        ev_val, ev_n, _, ev_low = _bucket_ev(signal, history) if signal not in ("—","PUSH","") else (None,0,0,True)
        if ev_val is not None:
            # Use logged odds if available, else -110 default
            eff_odds = odds_f if odds_f is not None else -110.0
            if eff_odds > 0:
                win_amt_ev = 10.0 * eff_odds / 100
                risk_ev    = 10.0
            else:
                win_amt_ev = 10.0 * 100 / abs(eff_odds)
                risk_ev    = 10.0
            # Recalculate EV with effective odds
            b_rows = [r2 for r2 in history if r2.get("model_ok") in ("✅","❌") and r2.get("model_signal","") == signal]
            win_rt = sum(1 for r2 in b_rows if r2.get("model_ok")=="✅") / len(b_rows) if b_rows else 0
            ev_computed = round(win_rt * win_amt_ev - (1-win_rt) * risk_ev, 2)
            ev_str = f"+${ev_computed:.2f}" if ev_computed >= 0 else f"-${abs(ev_computed):.2f}"
            if ev_low: ev_str += "*"
            ev_c = GREEN if ev_computed > 0 else RED if ev_computed < 0 else MUTED
            no_odds_note = "" if odds_f else "†"  # † = using -110 default
            ev_str += no_odds_note
        else:
            ev_str = "—"; ev_c = MUTED
        sc(ws3,log_row,17, ev_str, color=ev_c, bold=True if ev_val is not None and not ev_low else False, bg=bg, h="center", size=8)
        sc(ws3,log_row,14,my_bet,           color=TEAL if my_bet else MUTED,bg=bg,h="center")
        ws3.row_dimensions[log_row].height=15
        log_row+=1

    # Today's pitchers — only add if not already in history from scan
    already_in_history = {(r.get('date',''), r.get('pitcher','')) for r in history}
    for p in pitchers:
        if (TODAY, p['name']) in already_in_history:
            continue  # scan already wrote this row with real Vegas/odds data
        bg=rb(log_row); r=log_row
        conf=p.get("conf","MEDIUM")
        conf_c=conf_color.get(conf,WHITE)
        sc(ws3,r,1,TODAY,   color=MUTED,bg=bg,h="center")
        sc(ws3,r,2,p["name"],color=WHITE,bg=bg,h="left")
        sc(ws3,r,3,p["team"],color=MUTED,bg=bg,h="center")
        sc(ws3,r,4,p["opp"], color=MUTED,bg=bg,h="center")
        sc(ws3,r,5,p.get("proj"),color=WHITE,bg=bg,h="center")
        sc(ws3,r,6,"",      color=YELLOW,bg=bg,h="center")
        sc(ws3,r,7,p.get("vegas") or "N/A",color=YELLOW if p.get("vegas") else MUTED,bg=bg,h="center")
        for fc,formula in [
            (8, f'=IF(AND(E{r}<>"",F{r}<>""),F{r}-E{r},"")'),
            (9, f'=IF(H{r}="","",IF(ABS(H{r})<=0.5,"✅","❌"))'),
            (10,f'=IF(AND(E{r}<>"",G{r}<>""),IF(E{r}>G{r},"OVER",IF(E{r}<G{r},"UNDER","PUSH")),"")'),
            (11,f'=IF(AND(F{r}<>"",G{r}<>""),IF(F{r}>G{r},"OVER",IF(F{r}<G{r},"UNDER","PUSH")),"")'),
            (12,f'=IF(AND(J{r}<>"",K{r}<>"",J{r}<>"PUSH"),IF(J{r}=K{r},"✅","❌"),"—")'),
            (15,f'=IF(OR(E{r}="",G{r}="",G{r}="N/A"),"—",IF(ABS(E{r}-G{r})<=0.5,"PUSH",IF(E{r}-G{r}>0,IF(E{r}-G{r}<=1.0,"OVER 0.6-1.0",IF(E{r}-G{r}<=1.5,"OVER 1.1-1.5",IF(E{r}-G{r}<=2.0,"OVER 1.6-2.0",IF(E{r}-G{r}<=3.0,"OVER 2.1-3.0","OVER 3.1+")))),IF(G{r}-E{r}<=1.0,"UNDER 0.6-1.0",IF(G{r}-E{r}<=1.5,"UNDER 1.1-1.5",IF(G{r}-E{r}<=2.0,"UNDER 1.6-2.0",IF(G{r}-E{r}<=3.0,"UNDER 2.1-3.0","UNDER 3.1+")))))))'),
        ]:
            ws3.cell(r,fc).value=formula
            ws3.cell(r,fc).font=_font(color=MUTED,size=9)
            ws3.cell(r,fc).alignment=_align(h="center")
            ws3.cell(r,fc).fill=_fill(bg)
        sc(ws3,r,13,conf,color=conf_c,bold=True,bg=bg,h="center")
        sc(ws3,r,14,"",  color=MUTED,bg=bg,h="center")
        sc(ws3,r,16,"",  color=YELLOW,bg=bg,h="center")  # ODDS blank for today
        sc(ws3,r,17,"",  color=MUTED,bg=bg,h="center")   # EV blank until odds entered
        ws3.row_dimensions[r].height=15
        log_row+=1
    ws3.freeze_panes="A4"

    # ────────────────────────────────────────────────────────────
    # SHEET 4: Accuracy
    # ────────────────────────────────────────────────────────────
    ws4 = wb.create_sheet("🎯 Accuracy")
    for i,w in enumerate([18,10,14,14,14,14,12,12],1):
        ws4.column_dimensions[get_column_letter(i)].width=w

    ws4.merge_cells("A1:H1")
    sc(ws4,1,1,"🎯  MODEL ACCURACY  &  BETTING RECORD",bold=True,color=GREEN,bg=BG,size=12,h="center")
    ws4.row_dimensions[1].height=28

    # Stat cards
    cards=[
        ("All-Time Starts", "=COUNTA('📜 History Log'!B4:B500)",WHITE,1,None),
        ("Model vs O/U",    "=IFERROR(COUNTIF('📜 History Log'!L4:L500,\"✅\")/SUMPRODUCT(('📜 History Log'!L4:L500=\"✅\")+('📜 History Log'!L4:L500=\"❌\")),\"—\")",TEAL,3,"0.0%"),
        ("MY BETS W%",      "=IFERROR(COUNTIF('📈 My Bets'!K3:K200,\"✅\")/SUMPRODUCT(('📈 My Bets'!K3:K200=\"✅\")+('📈 My Bets'!K3:K200=\"❌\")),\"—\")",GREEN,5,"0.0%"),
        ("Total P/L",       "=IFERROR(SUM('📈 My Bets'!N3:N500),0)",YELLOW,7,"+0.00;-0.00;0.00"),
    ]
    ws4.row_dimensions[2].height=14; ws4.row_dimensions[3].height=26
    for label,formula,color,col,fmt in cards:
        ws4.merge_cells(start_row=2,start_column=col,end_row=2,end_column=col+1)
        sc(ws4,2,col,label,color=MUTED,bg=SURF1,size=8,h="center")
        ws4.merge_cells(start_row=3,start_column=col,end_row=3,end_column=col+1)
        sc(ws4,3,col,formula,bold=True,color=color,bg=SURF2,size=13,h="center",fmt=fmt)

    ws4.row_dimensions[4].height=15
    sc(ws4,4,1,"Bets Placed",color=MUTED,bg=SURF1,size=8,h="center")
    sc(ws4,4,2,"=SUMPRODUCT(('📈 My Bets'!K3:K200=\"✅\")+('📈 My Bets'!K3:K200=\"❌\"))",color=WHITE,bg=SURF1,size=9,h="center")
    sc(ws4,4,3,"ROI",color=MUTED,bg=SURF1,size=8,h="center")
    sc(ws4,4,4,'=IFERROR(SUM(\'📈 My Bets\'!N3:N500)/SUMIF(\'📈 My Bets\'!L3:L500,">0"),"—")',
       bold=True,color=TEAL,bg=SURF1,size=9,h="center",fmt="0.0%")
    ws4.row_dimensions[5].height=6

    # Confidence tier
    LOG="'📜 History Log'"
    ws4.merge_cells("A6:H6")
    sc(ws4,6,1,"🔵  MODEL ACCURACY BY CONFIDENCE TIER",bold=True,color=AMBER,bg=HEADER,size=9,h="left")
    ws4.row_dimensions[6].height=18
    for i,h in enumerate(["TIER","STARTS","MODEL vs O/U","AVG ERROR","MY BETS W%","MY BETS P/L","",""],1):
        sc(ws4,7,i,h,bold=True,color=MUTED,bg=HEADER,size=8,h="center")
    ws4.row_dimensions[7].height=15

    for ti,(tier,tc) in enumerate([("HIGH",GREEN),("MEDIUM",AMBER),("LOW",RED)],8):
        bg=rb(ti)
        sc(ws4,ti,1,tier,bold=True,color=tc,bg=bg,h="center")
        sc(ws4,ti,2,f'=COUNTIF({LOG}!M4:M1000,"{tier}")',color=WHITE,bg=bg,h="center")
        sc(ws4,ti,3,f'=IFERROR(COUNTIFS({LOG}!L4:L500,"✅",{LOG}!M4:M1000,"{tier}")/SUMPRODUCT(({LOG}!L4:L500="✅")*({LOG}!M4:M500="{tier}")+({LOG}!L4:L500="❌")*({LOG}!M4:M500="{tier}")),"—")',
           color=TEAL,bg=bg,h="center",fmt="0.0%")
        sc(ws4,ti,4,f'=IFERROR(AVERAGEIFS({LOG}!H4:H500,{LOG}!M4:M500,"{tier}",{LOG}!H4:H500,"<>"),"—")',
           color=AMBER,bg=bg,h="center",fmt="0.0")
        sc(ws4,ti,5,"—",color=MUTED,bg=bg,h="center")
        sc(ws4,ti,6,"—",color=MUTED,bg=bg,h="center")
        ws4.row_dimensions[ti].height=15
    ws4.row_dimensions[11].height=6

    # Monthly breakdown
    ws4.merge_cells("A12:H12")
    sc(ws4,12,1,"📅  MONTHLY BREAKDOWN",bold=True,color=AMBER,bg=HEADER,size=9,h="left")
    ws4.row_dimensions[12].height=18
    for i,h in enumerate(["MONTH","STARTS","GRADED","MODEL vs O/U","AVG ERROR","MODEL P/L","MY BETS W%","MY BETS P/L"],1):
        sc(ws4,13,i,h,bold=True,color=MUTED,bg=HEADER,size=8,h="center")
    ws4.row_dimensions[13].height=15

    # Monthly breakdown — live COUNTIFS so it updates as data comes in
    from collections import defaultdict
    months = sorted(set(
        str(r.get('date',''))[:7] for r in history
        if r.get('date') and len(str(r.get('date',''))) >= 7
    ))
    mr = 14
    BETS = "'📈 My Bets'"
    for month in months:
        if not month or month == 'da': continue
        bg = rb(mr)
        # Live Excel formulas — update automatically
        starts_f  = f"=COUNTIF({LOG}!A4:A1000,\"{month}*\")"
        graded_f  = f"=SUMPRODUCT(({LOG}!A4:A1000>=\"{month}-01\")*({LOG}!A4:A1000<=\"{month}-31\")*ISNUMBER(MATCH({LOG}!L4:L1000,{{\"✅\",\"❌\"}},0)))"
        correct_f = f"=SUMPRODUCT(({LOG}!A4:A1000>=\"{month}-01\")*({LOG}!A4:A1000<=\"{month}-31\")*({LOG}!L4:L1000=\"✅\"))"
        pct_f     = f"=IFERROR({correct_f[1:]}/{graded_f[1:]},\"—\")"
        err_f     = f"=IFERROR(AVERAGEIFS({LOG}!H4:H1000,{LOG}!A4:A1000,\">={month}-01\",{LOG}!A4:A1000,\"<={month}-31\",{LOG}!H4:H1000,\"<>\"),\"—\")"
        wins_f    = correct_f[1:]
        loss_f    = f"({graded_f[1:]}-{correct_f[1:]})"
        pl_f      = f"=IFERROR({wins_f}*18.18-{loss_f}*20,\"—\")"
        # My Take W% and P/L for this month from My Bets tab
        mt_wrate = (f'=IFERROR(COUNTIFS({BETS}!A3:A500,">={month}-01",{BETS}!A3:A500,"<={month}-31",{BETS}!K3:K500,"✅")/' +
                    f'SUMPRODUCT(({BETS}!A3:A500>="{month}-01")*({BETS}!A3:A500<="{month}-31")*(({BETS}!K3:K500="✅")+({BETS}!K3:K500="❌"))),"—")')
        mt_pl    = (f'=IFERROR(SUMPRODUCT(({BETS}!A3:A500>="{month}-01")*({BETS}!A3:A500<="{month}-31")*' +
                    f'IF(ISNUMBER({BETS}!N3:N500),{BETS}!N3:N500,0)),"—")')
        sc(ws4,mr,1,month,    color=WHITE, bg=bg,h="center",size=9)
        sc(ws4,mr,2,starts_f, color=WHITE, bg=bg,h="center",size=9)
        sc(ws4,mr,3,graded_f, color=MUTED, bg=bg,h="center",size=9)
        sc(ws4,mr,4,pct_f,    color=TEAL,  bold=True,bg=bg,h="center",size=9,fmt="0.0%")
        sc(ws4,mr,5,err_f,    color=AMBER, bg=bg,h="center",size=9,fmt="0.0")
        sc(ws4,mr,6,pl_f,     color=GREEN, bold=True,bg=bg,h="center",size=9,fmt="+$0.00;-$0.00;$0.00")
        sc(ws4,mr,7,mt_wrate, color=TEAL,  bold=True,bg=bg,h="center",size=9,fmt="0.0%")
        sc(ws4,mr,8,mt_pl,    color=GREEN, bold=True,bg=bg,h="center",size=9,fmt="+$0.00;-$0.00;$0.00")
        ws4.row_dimensions[mr].height=15
        mr+=1
    mr+=2

    # ── RECENT BREAKDOWN ─────────────────────────────────────
    import datetime as _dt
    _today_dt = _dt.date.fromisoformat(TODAY)
    _yesterday = _today_dt - _dt.timedelta(days=1)

    ws4.merge_cells(f"A{mr}:H{mr}")
    sc(ws4,mr,1,"📆  RECENT BREAKDOWN — Live formulas, update as results come in",
       bold=True,color=AMBER,bg=HEADER,size=9,h="left")
    ws4.row_dimensions[mr].height=18
    for i,h in enumerate(["PERIOD","STARTS","GRADED","WIN RATE","AVG ERROR","MODEL P/L","MY BETS W%","MY BETS P/L"],1):
        sc(ws4,mr+1,i,h,bold=True,color=MUTED,bg=HEADER,size=8,h="center")
    ws4.row_dimensions[mr+1].height=15

    def _recent_row(row, label, start_offset, end_offset):
        """Use TODAY()-N so dates stay current without rebuilding."""
        bg = rb(row)
        fd = f'TEXT(TODAY()-{start_offset},"YYYY-MM-DD")'
        td = f'TEXT(TODAY()-{end_offset},"YYYY-MM-DD")'
        starts_f  = f'=COUNTIFS({LOG}!A4:A1000,">="&{fd},{LOG}!A4:A1000,"<="&{td})'
        graded_f  = f'=SUMPRODUCT(({LOG}!A4:A1000>={fd})*({LOG}!A4:A1000<={td})*(({LOG}!L4:L1000="✅")+({LOG}!L4:L1000="❌")))'
        correct_i = f'SUMPRODUCT(({LOG}!A4:A1000>={fd})*({LOG}!A4:A1000<={td})*({LOG}!L4:L1000="✅"))'
        pct_f     = f'=IFERROR({correct_i}/{graded_f[1:]},"—")'
        err_f     = f'=IFERROR(AVERAGEIFS({LOG}!H4:H1000,{LOG}!A4:A1000,">="&{fd},{LOG}!A4:A1000,"<="&{td},{LOG}!H4:H1000,"<>"),"—")'
        pl_f      = f'=IFERROR({correct_i}*18.18-({graded_f[1:]}-{correct_i})*20,"—")'
        mt_w_f    = f'=IFERROR(COUNTIFS({BETS}!A3:A500,">="&{fd},{BETS}!A3:A500,"<="&{td},{BETS}!K3:K500,"✅")/SUMPRODUCT(({BETS}!A3:A500>={fd})*({BETS}!A3:A500<={td})*(({BETS}!K3:K500="✅")+({BETS}!K3:K500="❌"))),"—")'
        mt_pl_f   = f'=IFERROR(SUMPRODUCT(({BETS}!A3:A500>={fd})*({BETS}!A3:A500<={td})*IF(ISNUMBER({BETS}!N3:N500),{BETS}!N3:N500,0)),"—")'
        sc(ws4,row,1,label,   bold=True,color=WHITE,bg=bg,h="left")
        sc(ws4,row,2,starts_f,color=WHITE,bg=bg,h="center")
        sc(ws4,row,3,graded_f,color=MUTED,bg=bg,h="center")
        sc(ws4,row,4,pct_f,   color=TEAL,bold=True,bg=bg,h="center",fmt="0.0%")
        sc(ws4,row,5,err_f,   color=AMBER,bg=bg,h="center",fmt="0.0")
        sc(ws4,row,6,pl_f,    color=GREEN,bold=True,bg=bg,h="center",fmt="+$0.00;-$0.00;$0.00")
        sc(ws4,row,7,mt_w_f,  color=TEAL,bold=True,bg=bg,h="center",fmt="0.0%")
        sc(ws4,row,8,mt_pl_f, color=GREEN,bold=True,bg=bg,h="center",fmt="+$0.00;-$0.00;$0.00")
        ws4.row_dimensions[row].height=15

    _periods = [
        ("Yesterday",          1,  1),
        ("Last 3 Days",        3,  1),
        ("Last 7 Days",        7,  1),
        ("Week Prior (8-14d)", 14, 8),
        ("Last 14 Days",       14, 1),
    ]
    for _i, (_lbl, _start_off, _end_off) in enumerate(_periods):
        _recent_row(mr+2+_i, _lbl, _start_off, _end_off)

    sc(ws4,mr+7,1,"* Periods end yesterday  |  Win Rate only counts graded starts (have both Proj + Vegas)  |  My Bets P/L from My Bets tab",
       italic=True,color=DIM,bg=BG,size=8,h="left")
    ws4.row_dimensions[mr+7].height=13
    mr+=9


    # Edge sections
    PROJ=f"{LOG}!E4:E500"; VEG=f"{LOG}!G4:G500"; MOK=f"{LOG}!L4:L500"
    GRADED=f'(({MOK}="✅")+({MOK}="❌"))'; CORRECT=f'({MOK}="✅")'

    def edge_row(row, label, signal_label, lc):
        """Each bucket: counts, win rate, avg odds, EV per $10, profitable flag."""
        bg  = rb(row)
        SIG = f"{LOG}!O4:O1000"
        MOK = f"{LOG}!L4:L1000"
        cnt_all = f'COUNTIF({SIG},"{signal_label}")'
        cnt_ok  = f'COUNTIFS({SIG},"{signal_label}",{MOK},"✅")'
        s  = f"={cnt_all}"
        p  = f'=IFERROR({cnt_ok}/{cnt_all},"—")'
        b  = f'=IFERROR(IF({cnt_ok}/{cnt_all}>0.57,"✅ Bet it","—"),"—")'
        # EV and avg odds — computed from Python for reliability
        ev_val, ev_n, ev_odds_n, ev_low = _bucket_ev(signal_label, history)
        if ev_val is not None:
            ev_str = f"+${ev_val:.2f}" if ev_val >= 0 else f"-${abs(ev_val):.2f}"
            if ev_low: ev_str += "*"
            ev_c = GREEN if ev_val > 0 else RED if ev_val < 0 else MUTED
            b_rows = [r for r in history if r.get("model_ok") in ("✅","❌") and r.get("model_signal","") == signal_label]
            o_list = []
            for r in b_rows:
                try: o_list.append(float(r.get("odds") or -110))
                except: o_list.append(-110.0)
            avg_o = sum(o_list)/len(o_list) if o_list else -110.0
            avg_odds_str = _odds_display(avg_o)
        else:
            ev_str="—"; ev_c=MUTED; avg_odds_str="—"
        sc(ws4,row,1,label,        bold=True,color=lc,bg=bg,h="left")
        sc(ws4,row,2,s,            color=WHITE,bg=bg,h="center")
        sc(ws4,row,3,p,            color=TEAL,bold=True,bg=bg,h="center",fmt="0.0%")
        sc(ws4,row,4,avg_odds_str, color=MUTED,bg=bg,h="center",size=8)
        # Running profit formula — wins×$9.09 - losses×$10 (flat $10 at -110)
        rp_wins = f'COUNTIFS({SIG},"{signal_label}",{MOK},"✅")'
        rp_loss = f'COUNTIFS({SIG},"{signal_label}",{MOK},"❌")'
        rp_f    = f'=IFERROR({rp_wins}*9.09-{rp_loss}*10,"—")'
        sc(ws4,row,5,rp_f, color=GREEN,bold=True,bg=bg,h="center",fmt="+$0.00;-$0.00;$0.00")
        sc(ws4,row,6,b,            color=GREEN,bold=True,bg=bg,h="center")
        ws4.row_dimensions[row].height=15

    ws4.merge_cells(f"A{mr}:H{mr}")
    sc(ws4,mr,1,"📈  MODEL ACCURACY BY EDGE — OVERS",bold=True,color=GREEN,bg=HEADER,size=9,h="left")
    ws4.row_dimensions[mr].height=18
    for i,h in enumerate(["EDGE RANGE","STARTS","WIN RATE","AVG ODDS","RUNNING P/L","✅ BET IT?","",""],1):
        sc(ws4,mr+1,i,h,bold=True,color=MUTED,bg=HEADER,size=8,h="center")
    ws4.row_dimensions[mr+1].height=15
    edge_row(mr+2,"0.6 – 1.0","OVER 0.6-1.0",AMBER)
    edge_row(mr+3,"1.1 – 1.5","OVER 1.1-1.5",AMBER)
    edge_row(mr+4,"1.6 – 2.0","OVER 1.6-2.0",GREEN)
    edge_row(mr+5,"2.1 – 3.0","OVER 2.1-3.0",GREEN)
    edge_row(mr+6,"3.1+","OVER 3.1+",GREEN)
    ws4.row_dimensions[mr+7].height=6
    ur=mr+8
    ws4.merge_cells(f"A{ur}:H{ur}")
    sc(ws4,ur,1,"📉  MODEL ACCURACY BY EDGE — UNDERS",bold=True,color=RED,bg=HEADER,size=9,h="left")
    ws4.row_dimensions[ur].height=18
    for i,h in enumerate(["EDGE RANGE","STARTS","WIN RATE","AVG ODDS","RUNNING P/L","✅ BET IT?","",""],1):
        sc(ws4,ur+1,i,h,bold=True,color=MUTED,bg=HEADER,size=8,h="center")
    ws4.row_dimensions[ur+1].height=15
    edge_row(ur+2,"-0.6 to -1.0","UNDER 0.6-1.0",AMBER)
    edge_row(ur+3,"-1.1 to -1.5","UNDER 1.1-1.5",AMBER)
    edge_row(ur+4,"-1.6 to -2.0","UNDER 1.6-2.0",GREEN)
    edge_row(ur+5,"-2.1 to -3.0","UNDER 2.1-3.0",GREEN)
    edge_row(ur+6,"-3.1 and below","UNDER 3.1+",GREEN)
    fn=ur+7
    ws4.merge_cells(f"A{fn}:H{fn}")
    sc(ws4,fn,1,"📌 * = fewer than 30 graded starts  |  RUNNING P/L = flat $10 at -110 (win +$9.09 / loss -$10)  |  MY BETS uses actual bet amounts  |  ✅ Bet it = >57% win rate",
       italic=True,color=DIM,bg=BG,size=8,h="center")
    ws4.row_dimensions[fn].height=13

    # ── MODEL SIGNAL ACCURACY ────────────────────────────────
    sig_mr = fn + 2

    ws4.merge_cells(f"A{sig_mr}:H{sig_mr}")
    sc(ws4,sig_mr,1,"📊  MODEL SIGNAL ACCURACY — Including PUSH Analysis",
       bold=True,color=AMBER,bg=HEADER,size=9,h="left")
    ws4.row_dimensions[sig_mr].height=18

    for i,h in enumerate(["SIGNAL TYPE","TOTAL","OVER %","UNDER %","WITHIN ±0.5","AVG ERROR","",""],1):
        sc(ws4,sig_mr+1,i,h,bold=True,color=MUTED,bg=HEADER,size=8,h="center")
    ws4.row_dimensions[sig_mr+1].height=15

    SIG_COL = f"{LOG}!O4:O1000"
    ACT_COL = f"{LOG}!K4:K1000"   # ACTUAL O/U (K)
    WTH_COL = f"{LOG}!I4:I1000"   # WITHIN 0.5 (I)
    ERR_COL = f"{LOG}!H4:H1000"   # ERROR (H)

    # Compute signal stats from Python for accuracy
    from collections import defaultdict
    sig_stats = defaultdict(lambda: {"total":0,"over":0,"under":0,"push":0,"within":0,"errors":[]})
    for r in history:
        sig = r.get("model_signal","")
        if not sig or sig == "—": continue
        aou = r.get("actual_ou","") or ""
        # Determine actual O/U from the record
        try:
            actual_v = float(r.get("actual_ks") or 0)
            vegas_v  = float(r.get("vegas_ou") or 0)
            if actual_v and vegas_v:
                aou = "OVER" if actual_v > vegas_v else "UNDER" if actual_v < vegas_v else "PUSH"
        except: pass
        try:
            proj_v   = float(r.get("proj_ks") or 0)
            actual_v2 = float(r.get("actual_ks") or 0)
            if proj_v and actual_v2:
                err = actual_v2 - proj_v
                sig_stats[sig]["errors"].append(err)
                if abs(err) <= 0.5: sig_stats[sig]["within"] += 1
        except: pass
        sig_stats[sig]["total"] += 1
        if aou == "OVER":  sig_stats[sig]["over"]  += 1
        elif aou == "UNDER": sig_stats[sig]["under"] += 1
        elif aou == "PUSH":  sig_stats[sig]["push"]  += 1

    # Group: ALL signals, PUSH only, OVER signals combined, UNDER signals combined
    push_d = sig_stats.get("PUSH", {"total":0,"over":0,"under":0,"within":0,"errors":[]})

    # Aggregate all OVER signals
    over_all = {"total":0,"over":0,"under":0,"within":0,"errors":[]}
    under_all = {"total":0,"over":0,"under":0,"within":0,"errors":[]}
    all_sigs  = {"total":0,"over":0,"under":0,"push":0,"within":0,"errors":[]}
    for sig, d in sig_stats.items():
        all_sigs["total"]  += d["total"]
        all_sigs["over"]   += d["over"]
        all_sigs["under"]  += d["under"]
        all_sigs["push"]   += d.get("push",0)
        all_sigs["within"] += d["within"]
        all_sigs["errors"] += d["errors"]
        if sig.startswith("OVER"):
            for k in ["total","over","under","within"]: over_all[k] += d[k]
            over_all["errors"] += d["errors"]
        elif sig.startswith("UNDER"):
            for k in ["total","over","under","within"]: under_all[k] += d[k]
            under_all["errors"] += d["errors"]

    def _sig_row(row, label, d, lc=WHITE):
        bg = rb(row)
        n = d["total"]
        if n == 0:
            for c in range(1,7): sc(ws4,row,c,"—",color=MUTED,bg=bg,h="center")
            sc(ws4,row,1,label,bold=True,color=lc,bg=bg,h="left")
            ws4.row_dimensions[row].height=15; return
        over_pct  = f"{d['over']/n:.1%}"  if d['over']  else "—"
        under_pct = f"{d['under']/n:.1%}" if d['under'] else "—"
        within_pct= f"{d['within']/n:.1%}" if d['within'] else "—"
        avg_err   = round(sum(d['errors'])/len(d['errors']),2) if d['errors'] else "—"
        over_c  = GREEN if d["over"] > d["under"] else MUTED
        under_c = RED   if d["under"] > d["over"] else MUTED
        sc(ws4,row,1,label,     bold=True,color=lc,bg=bg,h="left")
        sc(ws4,row,2,n,         color=WHITE,bg=bg,h="center")
        sc(ws4,row,3,over_pct,  color=over_c,bold=True,bg=bg,h="center")
        sc(ws4,row,4,under_pct, color=under_c,bold=True,bg=bg,h="center")
        sc(ws4,row,5,within_pct,color=TEAL,bold=True,bg=bg,h="center")
        sc(ws4,row,6,avg_err,   color=AMBER,bg=bg,h="center")
        ws4.row_dimensions[row].height=15

    _sig_row(sig_mr+2, "ALL Signals",        all_sigs,  WHITE)
    _sig_row(sig_mr+3, "PUSH (±0.5 edge)",   push_d,    AMBER)
    _sig_row(sig_mr+4, "OVER Signals (all)", over_all,  GREEN)
    _sig_row(sig_mr+5, "UNDER Signals (all)",under_all, RED)

    # Also show per-bucket breakdown for PUSH specifically
    sig_mr2 = sig_mr + 7
    ws4.merge_cells(f"A{sig_mr2}:H{sig_mr2}")
    sc(ws4,sig_mr2,1,
       "When MODEL SIGNAL = PUSH:  Shows whether actuals still leaned OVER or UNDER the line — useful for finding hidden bias in neutral signals",
       italic=True,color=DIM,bg=BG,size=8,h="left")
    ws4.row_dimensions[sig_mr2].height=13

    sc(ws4,sig_mr+6,1,
       f"PUSH total: {push_d['total']} starts  |  "
       f"Went OVER: {push_d['over']} ({push_d['over']/push_d['total']:.1%} of PUSHes)  |  "
       f"Went UNDER: {push_d['under']} ({push_d['under']/push_d['total']:.1%})  |  "
       f"Landed within ±0.5: {push_d['within']} ({push_d['within']/push_d['total']:.1%})" if push_d["total"] else "No PUSH data yet",
       italic=True,color=AMBER,bg=rb(sig_mr+6),size=8,h="left")
    ws4.row_dimensions[sig_mr+6].height=13

    ws4.freeze_panes="A2"

    return wb

def _safe_float(v):
    if v is None or v=="": return None
    try: return float(v)
    except: return None

# ── Send email ─────────────────────────────────────────────────
def send_email(filepath):
    if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
        print(f"⚠️  No email credentials — skipping send. File: {filepath}")
        return
    msg=MIMEMultipart()
    msg["From"]=EMAIL_ADDRESS; msg["To"]=EMAIL_ADDRESS
    msg["Subject"]=f"⚾ MLB K Projections — {DATE_LABEL}"
    body=f"""MLB STRIKEOUT PROJECTIONS — {DATE_LABEL}
====================================================

Master tracker attached.

📅 Today's Slate  — fill in Vegas O/U + My Take (yellow cells)
📈 My Bets        — log any real money bets here
📜 History Log    — actuals sync automatically tomorrow morning
🎯 Accuracy       — updates automatically

⚾  Good luck today.
"""
    msg.attach(MIMEText(body,"plain"))
    with open(filepath,"rb") as f:
        part=MIMEBase("application","octet-stream"); part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition",f'attachment; filename="{Path(filepath).name}"')
    msg.attach(part)
    with smtplib.SMTP(SMTP_SERVER,SMTP_PORT) as server:
        server.starttls(); server.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS,EMAIL_ADDRESS,msg.as_string())
    print(f"✅ Email sent to {EMAIL_ADDRESS}")


def resync_all():
    """
    Re-reads every MLB_Ks_Master_*.xlsx in the Output folder.
    Pulls Vegas O/U, My Take, and Actual Ks from each one.
    Rebuilds history.csv and bets.csv from scratch with correct data.
    """
    import glob
    from openpyxl import load_workbook

    print("🔄 RESYNC MODE — scanning all master files in Output folder...")
    master_files = sorted(glob.glob(str(OUTPUT_DIR / "MLB_Ks_Master_*.xlsx")))
    # Exclude backfill and today's file
    master_files = [f for f in master_files if "_BACKFILL" not in f]
    print(f"   Found {len(master_files)} master files:")
    for f in master_files:
        print(f"   → {Path(f).name}")

    history = load_history()
    bets    = load_bets()
    existing_hist = {(r["date"], r["pitcher"]): r for r in history}
    existing_bets = {(b["date"], b["pitcher"]): b for b in bets}
    h_updated = 0; b_updated = 0

    for filepath in master_files:
        fname = Path(filepath).name
        # Extract date from filename
        import re as _re
        m = _re.search(r"(\d{4}-\d{2}-\d{2})", fname)
        if not m: continue
        file_date = m.group(1)
        try:
            wb = load_workbook(filepath, data_only=True)
        except Exception as e:
            print(f"   ⚠️  Could not open {fname}: {e}")
            continue

        # ── Pull from History Log ──────────────────────────────
        hist_sheet = None
        for sn in wb.sheetnames:
            if "History" in sn: hist_sheet = wb[sn]; break
        if hist_sheet:
            for r in range(4, 500):
                p = hist_sheet.cell(r, 2).value
                if not p: break
                date   = _normalize_date(hist_sheet.cell(r, 1).value)
                actual = hist_sheet.cell(r, 6).value
                vegas  = hist_sheet.cell(r, 7).value
                my_bet = hist_sheet.cell(r, 14).value
                proj   = hist_sheet.cell(r, 5).value
                if not date or not p: continue
                name = str(p).strip()
                key  = (date, name)
                # Clean my_bet
                if my_bet:
                    mb = str(my_bet).strip()
                    if "(" in mb: mb = mb.split("(")[0].strip()
                    my_bet = mb
                if key in existing_hist:
                    rec = existing_hist[key]; changed = False
                    if actual not in (None,"") and str(rec.get("actual_ks","")) in ("","None"):
                        rec["actual_ks"] = str(actual); changed = True
                    if vegas not in (None,"") and str(rec.get("vegas_ou","")) in ("","None"):
                        rec["vegas_ou"] = str(vegas); changed = True
                    if my_bet and not rec.get("my_bet"):
                        rec["my_bet"] = my_bet
                    if changed:
                        try:
                            pj=float(rec.get("proj_ks","") or 0)
                            ac=float(rec["actual_ks"] or 0)
                            vg=float(rec["vegas_ou"] or 0)
                            if pj and ac and vg:
                                mc="OVER" if pj>vg else "UNDER" if pj<vg else "PUSH"
                                ao="OVER" if ac>vg else "UNDER" if ac<vg else "PUSH"
                                rec["model_ok"]="✅" if mc==ao else "❌" if mc in ("OVER","UNDER") and ao in ("OVER","UNDER") else "—"
                        except: pass
                        h_updated += 1
                else:
                    if actual or vegas:
                        new_rec = {
                            "date":date,"pitcher":name,
                            "team":str(hist_sheet.cell(r,3).value or ""),
                            "opp":str(hist_sheet.cell(r,4).value or ""),
                            "proj_ks":str(proj or ""),"actual_ks":str(actual or ""),
                            "vegas_ou":str(vegas or ""),
                            "conf":str(hist_sheet.cell(r,13).value or ""),
                            "my_bet":str(my_bet or ""),"model_ok":"",
                        }
                        try:
                            pj=float(new_rec["proj_ks"] or 0)
                            ac=float(new_rec["actual_ks"] or 0)
                            vg=float(new_rec["vegas_ou"] or 0)
                            if pj and ac and vg:
                                mc="OVER" if pj>vg else "UNDER" if pj<vg else "PUSH"
                                ao="OVER" if ac>vg else "UNDER" if ac<vg else "PUSH"
                                new_rec["model_ok"]="✅" if mc==ao else "❌" if mc in ("OVER","UNDER") and ao in ("OVER","UNDER") else "—"
                        except: pass
                        history.append(new_rec)
                        existing_hist[key] = new_rec
                        h_updated += 1

        # ── Pull from My Bets ──────────────────────────────────
        bets_sheet = None
        for sn in wb.sheetnames:
            if "Bets" in sn and "My" in sn: bets_sheet = wb[sn]; break
        if bets_sheet:
            for r in range(3, 200):
                p = bets_sheet.cell(r, 2).value
                if not p or str(p) in ("PITCHER","TOTALS","ROI","WIN RATE"): continue
                date_v  = bets_sheet.cell(r, 1).value
                date_s  = _normalize_date(date_v) if date_v else ""
                vegas   = bets_sheet.cell(r, 6).value
                my_take = bets_sheet.cell(r, 7).value
                if not my_take or not vegas: continue
                odds    = bets_sheet.cell(r, 8).value
                actual  = bets_sheet.cell(r, 9).value
                correct = bets_sheet.cell(r, 11).value
                units   = bets_sheet.cell(r, 12).value
                pl      = bets_sheet.cell(r, 14).value
                key = (date_s, str(p).strip())
                if key not in existing_bets:
                    try: units_v = abs(float(units or 0))
                    except: units_v = 0
                    existing_bets[key] = {
                        "date":date_s,"pitcher":str(p).strip(),
                        "team":str(bets_sheet.cell(r,3).value or ""),
                        "opp":str(bets_sheet.cell(r,4).value or ""),
                        "proj_ks":bets_sheet.cell(r,5).value or "",
                        "vegas":vegas,"my_take":str(my_take).strip(),
                        "odds":odds or -110,
                        "actual_ks":actual or "","correct":str(correct or ""),
                        "units":units_v,"outcome":"","pl":pl or "",
                    }
                    b_updated += 1
                else:
                    rec = existing_bets[key]
                    if actual not in (None,"") and str(rec.get("actual_ks","")) in ("","None"):
                        rec["actual_ks"] = actual
                        rec["correct"]   = str(correct or "")
                        rec["pl"]        = pl or ""
                        b_updated += 1

    save_history(list(existing_hist.values()))
    bets_list = list(existing_bets.values())
    save_bets(bets_list)

    graded = sum(1 for r in history if r.get("model_ok") in ("✅","❌"))
    correct_ct = sum(1 for r in history if r.get("model_ok")=="✅")
    print(f"   ✅ History: {h_updated} rows updated, {len(history)} total, {correct_ct}/{graded} graded")
    print(f"   ✅ Bets: {b_updated} rows updated, {len(bets_list)} total")
    return list(existing_hist.values()), bets_list

# ── Main ───────────────────────────────────────────────────────
def main():
    import urllib.parse

    if BACKFILL_MODE:
        print(f"⚾  MLB K Tracker — {TODAY}  [BACKFILL MODE — no sync, no email]")
    elif RESYNC_MODE:
        print(f"⚾  MLB K Tracker — {TODAY}  [RESYNC MODE — rebuilding from all Output files]")
    else:
        print(f"⚾  MLB K Tracker — {TODAY}")
    print("─" * 50)

    # 1. Load history and bets
    history = load_history()
    bets    = load_bets()
    print(f"📜 History: {len(history)} starts | 📈 Bets: {len(bets)}")

    # 2. Always scan Output files first to capture any manual inputs
    if not BACKFILL_MODE:
        history, bets = scan_output_files(history, bets)

    # 3. Then sync or resync
    if RESYNC_MODE:
        history, bets = resync_all()
    elif not BACKFILL_MODE:
        history, bets = sync_yesterday(history, bets)
    else:
        print(f"⏭️  Skipping sync (backfill mode)")

    # 3. Pull probable pitchers for the target date
    print(f"🌐 Fetching probable pitchers for {TODAY}...")
    probable = get_probable_pitchers_today()
    if not probable:
        print("   ⚠️  Could not fetch probables — check the date or try again")
        probable = []
    else:
        print(f"   ✅ Found {len(probable)} pitchers")

    # 4. Pull team K% (uses current season stats — always valid)
    print("🌐 Fetching team K% data...")
    team_kpct = get_team_kpct()
    if team_kpct:
        print(f"   ✅ Got K% for {len(team_kpct)} teams")
    else:
        print("   ⚠️  Could not fetch team K% — showing —")

    # 5. Pull game logs for each pitcher
    # In backfill mode: filter out starts ON or AFTER target date
    # so projections only use data that existed at game time
    print("🌐 Fetching pitcher game logs...")
    pitchers = []
    for name, pid, team, opp, game_time in probable:
        if not pid:
            pid = search_player_id(name)
        starts = get_pitcher_gamelogs(pid) if pid else []

        if BACKFILL_MODE and starts:
            starts = [s for s in starts if s[0] < TODAY]

        real_starts = [s for s in starts if s[2] >= 2.0]
        avg_ks = round(sum(s[1] for s in real_starts) / len(real_starts), 1) if real_starts else None
        opp_kpct = team_kpct.get(opp)
        proj   = project_ks(starts, opp_kpct=opp_kpct) if real_starts else None

        if len(real_starts) >= 3 and proj:
            conf = "HIGH"
        elif len(real_starts) >= 1 and proj:
            conf = "MEDIUM"
        else:
            conf = "LOW"

        opp_kpct = team_kpct.get(opp)

        pitchers.append({
            "name":     name,
            "team":     team,
            "opp":      opp,
            "proj":     proj or 5.0,
            "avg_ks":   avg_ks,
            "ip_ceil":  5.5,
            "conf":     conf,
            "opp_kpct": opp_kpct,
            "starts":   starts,
            "vegas":    None,
        })
        print(f"   {name}: {len(starts)} prior starts, proj={proj or 'N/A'}, opp_k%={opp_kpct}")
        time.sleep(0.2)

    pitchers.sort(key=lambda x: x["proj"] or 0, reverse=True)
    print(f"✅ Built slate: {len(pitchers)} pitchers")

    # 6. Build Excel
    wb = build_master(pitchers, history, bets)
    if BACKFILL_MODE:
        out_path = OUTPUT_DIR / f"MLB_Ks_Master_{TODAY}_BACKFILL.xlsx"
        wb.save(out_path)
        print(f"💾 Saved: {out_path.name}")
        print(f"")
        print(f"📋 BACKFILL INSTRUCTIONS:")
        print(f"   1. Open {out_path.name}")
        print(f"   2. Fill in VEGAS O/U + MY TAKE on Today's Slate tab")
        print(f"   3. Fill in ACTUAL Ks on the History Log tab")
        print(f"   4. Save as MLB_Ks_Master_{TODAY}.xlsx in Output folder")
        print(f"   5. Accuracy tab grades itself automatically")
    else:
        out_path = OUTPUT_DIR / f"MLB_Ks_Master_{TODAY}.xlsx"
        wb.save(out_path)
        print(f"💾 Saved: {out_path.name}")
        send_email(out_path)

if __name__ == "__main__":
    main()

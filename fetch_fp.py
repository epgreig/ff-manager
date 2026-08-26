"""Fetch FantasyPros consensus projections via the public API (free tier).

The free tier caps every response at 10 players, but the players= ID filter
reaches any player, so we batch 10 IDs per request. Candidate players are
seeded from FantasyFootballCalculator ADP plus (when present) the WinWithOdds
table, and mapped to FantasyPros IDs via the DynastyProcess ID database.

Spends API quota (50 requests/day): ~1 request per 10 offensive candidates
plus 2 for K/DST. A same-day compiled cache short-circuits the fetch; use
--force to refetch anyway.

Usage: python3 fetch_fp.py [--force]
"""

import json
import sys
import time
from datetime import date, timedelta
from pathlib import Path

import config
from common import ROOT, find_wwo, http_get, load_env, norm_name, parse_wwo, read_csv_dicts, write_csv_dicts

API = "https://api.fantasypros.com/public/v2/json/nfl"
DP_URL = "https://raw.githubusercontent.com/dynastyprocess/data/master/files/db_playerids.csv"
FFC_URL = (
    "https://fantasyfootballcalculator.com/api/v1/adp/"
    f"{config.FFC_FORMAT}?teams={config.LEAGUE_TEAMS}&year={config.SEASON}"
)

CACHE = ROOT / "fp_cache"
DP_PATH = CACHE / "dp_ids.csv"
OUT_PATH = CACHE / "fp_projections.csv"
STAMP_PATH = CACHE / "fp_projections.date"

OFFENSE = ("QB", "RB", "WR", "TE")


def dp_id_map() -> dict[str, dict]:
    """norm-name|pos -> DP row, refreshed weekly."""
    if not DP_PATH.exists() or date.fromtimestamp(DP_PATH.stat().st_mtime) < date.today() - timedelta(days=7):
        print("Downloading DynastyProcess ID map...")
        DP_PATH.parent.mkdir(exist_ok=True)
        DP_PATH.write_bytes(http_get(DP_URL))
    rows = read_csv_dicts(DP_PATH)
    out = {}
    for r in rows:
        if r.get("fantasypros_id", "").isdigit() and r.get("position") in OFFENSE:
            out[f"{norm_name(r['name'])}|{r['position']}"] = r
    return out


def candidate_names() -> dict[str, list[str]]:
    """Per-position candidate display names, best first, capped by config.COUNTS."""
    per_pos: dict[str, list[str]] = {p: [] for p in OFFENSE}
    seen: set[str] = set()

    def add(name: str, pos: str):
        key = f"{norm_name(name)}|{pos}"
        if pos in per_pos and key not in seen and len(per_pos[pos]) < config.COUNTS[pos]:
            per_pos[pos].append(name)
            seen.add(key)

    print("Fetching FFC ADP for candidate seeding...")
    ffc = json.loads(http_get(FFC_URL))
    for p in sorted(ffc["players"], key=lambda p: p["adp"]):
        add(p["name"], p["position"])

    wwo_path = find_wwo(ROOT / "data", config.WWO_FILE)
    if wwo_path:
        wwo = parse_wwo(wwo_path, config.WWO_COLS, config.SCORING_RULES)
        for r in sorted(wwo, key=lambda r: -r["points"]):
            add(r["name"], r["position"])
        print(f"Seeded candidates from FFC ADP + {wwo_path.name}")
    else:
        print("No WinWithOdds file in data/ yet — seeding from FFC ADP only")
    return per_pos


class QuotaExhausted(Exception):
    pass


def api_get(path_and_query: str, key: str) -> dict:
    """API call with disk cache so crashes/reruns never re-spend quota.

    The free tier allows ~50 requests per rolling 24h window; past the limit
    the API returns empty count:0 responses, then hard 429s. Both raise
    QuotaExhausted so the caller can compile a partial result and resume later.
    Cached batches are reused for 24h (keyed by fetch date).
    """
    import hashlib
    import urllib.error

    cache_file = CACHE / "batches" / f"{date.today()}_{hashlib.md5(path_and_query.encode()).hexdigest()[:12]}.json"
    yesterday_file = Path(str(cache_file).replace(str(date.today()), str(date.today() - timedelta(days=1))))
    for f in (cache_file, yesterday_file):
        if f.exists():
            return json.loads(f.read_text())
    try:
        resp = json.loads(http_get(f"{API}/{path_and_query}", headers={"x-api-key": key}))
    except urllib.error.HTTPError as e:
        if e.code == 429:
            raise QuotaExhausted("HTTP 429") from e
        raise
    if not resp.get("players"):
        raise QuotaExhausted(f"empty response (count={resp.get('count')})")
    cache_file.parent.mkdir(parents=True, exist_ok=True)
    cache_file.write_text(json.dumps(resp))
    return resp


def flatten(player: dict) -> dict:
    s = player.get("stats", {})
    std = s.get("points", "")
    return {
        "fpid": player["fpid"],
        "name": player["name"],
        "position": player["position_id"],
        "team": player.get("team_id", ""),
        "points_std": std,
        "points_ppr": s.get("points_ppr", std),
        "points_half": s.get("points_half", std),
    }


def main():
    force = "--force" in sys.argv
    if STAMP_PATH.exists() and STAMP_PATH.read_text().strip() == str(date.today()) and not force:
        print(f"{OUT_PATH} already fetched today — skipping (use --force to refetch).")
        return

    key = load_env().get("FP_API_KEY")
    if not key:
        raise SystemExit("FP_API_KEY missing from .env")

    dp = dp_id_map()
    cands = candidate_names()

    fpids, unmapped = [], []
    for pos, names in cands.items():
        for name in names:
            row = dp.get(f"{norm_name(name)}|{pos}")
            if row:
                fpids.append(row["fantasypros_id"])
            else:
                unmapped.append(f"{name} ({pos})")
    if unmapped:
        print(f"Not in DP ID map (skipped): {', '.join(unmapped)}")

    print(f"Candidate pool: {len(fpids)} offensive players + top-10 K/DST")

    # Preload cached batches no matter which query produced them, so the cache
    # is immune to candidate reordering between runs. Fresh (<24h) batches
    # satisfy a player outright; older ones are kept as a fallback used only
    # if the quota runs out before that player is refetched.
    players: dict[int, dict] = {}
    stale: dict[int, dict] = {}
    batches_dir = CACHE / "batches"
    if batches_dir.exists():
        import time as _t
        for f in batches_dir.glob("*.json"):
            target = players if _t.time() - f.stat().st_mtime < 86400 else stale
            for p in json.loads(f.read_text()).get("players") or []:
                target[p["fpid"]] = flatten(p)
    if players or stale:
        print(f"Preloaded {len(players)} fresh + {len(stale)} stale players from batch cache")

    missing = sorted(fid for fid in set(fpids) if int(fid) not in players)
    have_pos = {r["position"] for r in players.values()}
    complete = True
    queries = [
        f"{config.SEASON}/projections?week=0&players={':'.join(missing[i:i + 10])}"
        for i in range(0, len(missing), 10)
    ] + [
        f"{config.SEASON}/projections?week=0&position={pos}"
        for pos in ("K", "DST") if pos not in have_pos
    ]
    print(f"{len(missing)} players missing -> {len(queries)} API requests needed")
    for n, q in enumerate(queries):
        try:
            resp = api_get(q, key)
        except QuotaExhausted as e:
            complete = False
            print(f"\nAPI quota exhausted at batch {n + 1}/{len(queries)} ({e}).")
            print("Compiling partial output; rerun fetch_fp.py after the 24h window rolls —")
            print("cached batches are reused, only the missing ones will be requested.")
            break
        for p in resp.get("players") or []:
            players[p["fpid"]] = flatten(p)
        time.sleep(0.4)

    filled = 0
    for fid, p in stale.items():
        if fid not in players:
            players[fid] = p
            filled += 1
    if filled:
        print(f"Filled {filled} players from stale cache (older than 24h — refetch when quota allows)")

    rows = sorted(players.values(), key=lambda r: (r["position"], -float(r["points_half"] or 0)))
    write_csv_dicts(
        OUT_PATH, rows,
        ["fpid", "name", "position", "team", "points_std", "points_ppr", "points_half"],
    )
    if complete:
        STAMP_PATH.write_text(str(date.today()))
    from collections import Counter

    print(f"Wrote {len(rows)} players to {OUT_PATH}{'' if complete else ' (PARTIAL)'}")
    print("  by position:", dict(Counter(r["position"] for r in rows)))


if __name__ == "__main__":
    main()

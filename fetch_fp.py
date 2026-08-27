"""Fetch FantasyPros consensus projections via the public API (free tier).

The free tier caps every response at 10 players, but the players= ID filter
reaches any player, so we batch 10 IDs per request. Candidate players are
seeded from FantasyFootballCalculator ADP plus (when present) the WinWithOdds
table, and mapped to FantasyPros IDs via the DynastyProcess ID database.

Spends API quota (~50 requests/day, reset timing unclear): ~1 request per 10
offensive candidates plus 2 for K/DST.

Usage: python3 fetch_fp.py --full   # refetch everything, ignoring cached batches
       python3 fetch_fp.py --cache  # no network at all: compile from cached batches
With no flag, prompts for one of the two. A --full run that dies to quota can
be rerun: batches it already fetched today are served from disk without
re-spending, so it resumes where it died.
"""

import json
import sys
import time
from datetime import date, timedelta
from pathlib import Path

import config
from common import (ROOT, find_wwo, http_get, load_aliases, load_env, norm_name, parse_wwo,
                    read_csv_dicts, write_csv_dicts)

API = "https://api.fantasypros.com/public/v2/json/nfl"
DP_URL = "https://raw.githubusercontent.com/dynastyprocess/data/master/files/db_playerids.csv"
FFC_URL = (
    "https://fantasyfootballcalculator.com/api/v1/adp/"
    f"{config.FFC_FORMAT}?teams={config.LEAGUE_TEAMS}&year={config.SEASON}"
)

CACHE = ROOT / "fp_cache"
DP_PATH = CACHE / "dp_ids.csv"
OUT_PATH = CACHE / "fp_projections.csv"

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
    raw = http_get(FFC_URL)
    (CACHE / "ffc_adp.json").write_bytes(raw)  # blend.py joins bye + ADP from this
    ffc = json.loads(raw)
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


def api_get(path_and_query: str, key: str, reuse_secs: int = 1800) -> dict:
    """API call with a short-lived disk cache: a quota-killed run rerun within
    reuse_secs resumes without re-spending, while a deliberate later --full
    genuinely refetches. Past the quota the API returns empty count:0
    responses, then hard 429s; both raise QuotaExhausted so the caller can
    compile a partial result and resume later.
    """
    import hashlib
    import time as _t
    import urllib.error

    h = hashlib.md5(path_and_query.encode()).hexdigest()[:12]
    cache_file = CACHE / "batches" / f"{date.today()}_{h}.json"
    for f in (CACHE / "batches").glob(f"*_{h}.json") if (CACHE / "batches").exists() else []:
        if _t.time() - f.stat().st_mtime < reuse_secs:
            return json.loads(f.read_text())

    # The API appears to throttle (~1 req/sec): an over-pace request can come
    # back empty or 429 without being a real quota problem. Back off and retry
    # a couple of times before concluding the quota is gone.
    last = ""
    for attempt, backoff in enumerate((0, 4, 10)):
        if backoff:
            print(f"    (empty/429 response — backing off {backoff}s and retrying)")
            _t.sleep(backoff)
        try:
            resp = json.loads(http_get(f"{API}/{path_and_query}", headers={"x-api-key": key}))
        except urllib.error.HTTPError as e:
            if e.code == 429:
                last = "HTTP 429"
                continue
            raise
        if resp.get("players"):
            cache_file.parent.mkdir(parents=True, exist_ok=True)
            cache_file.write_text(json.dumps(resp))
            return resp
        last = f"empty response (count={resp.get('count')})"
    raise QuotaExhausted(f"{last} after retries with backoff")


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


def pick_mode() -> str:
    if "--full" in sys.argv:
        return "full"
    if "--cache" in sys.argv:
        return "cache"
    try:
        ans = input("Mode — 'full' (refetch everything from the API) or 'cache' (no network, compile from cached batches)? ").strip().lower()
    except EOFError:
        ans = ""
    if ans in ("full", "f"):
        return "full"
    if ans in ("cache", "c"):
        return "cache"
    raise SystemExit("Pick one: python3 fetch_fp.py --full | --cache")


def compile_and_write(players: dict, complete: bool):
    rows = sorted(players.values(), key=lambda r: (r["position"], -float(r["points_half"] or 0)))
    write_csv_dicts(
        OUT_PATH, rows,
        ["fpid", "name", "position", "team", "points_std", "points_ppr", "points_half"],
    )
    from collections import Counter

    print(f"Wrote {len(rows)} players to {OUT_PATH}{'' if complete else ' (PARTIAL)'}")
    print("  by position:", dict(Counter(r["position"] for r in rows)))


def load_batch_cache() -> tuple[dict, dict]:
    """(fresh, stale) player dicts from every cached batch, split at 24h."""
    import time as _t

    fresh: dict[int, dict] = {}
    stale: dict[int, dict] = {}
    batches_dir = CACHE / "batches"
    if batches_dir.exists():
        for f in batches_dir.glob("*.json"):
            target = fresh if _t.time() - f.stat().st_mtime < 86400 else stale
            for p in json.loads(f.read_text()).get("players") or []:
                target[p["fpid"]] = flatten(p)
    return fresh, stale


def main():
    mode = pick_mode()

    if mode == "cache":
        fresh, stale = load_batch_cache()
        players = {**stale, **fresh}
        if not players:
            raise SystemExit("Batch cache is empty — nothing to compile. Run with --full.")
        print(f"Compiling from cache only: {len(fresh)} fresh + "
              f"{len(stale) - len(set(stale) & set(fresh))} stale-only players, no API calls.")
        compile_and_write(players, complete=False)
        return

    key = load_env().get("FP_API_KEY")
    if not key:
        raise SystemExit("FP_API_KEY missing from .env")

    dp = dp_id_map()
    cands = candidate_names()

    aliases = load_aliases()
    fpids, unmapped = [], []
    id2name: dict[str, str] = {}
    for pos, names in cands.items():
        for name in names:
            if norm_name(name) in aliases:
                fpids.append(aliases[norm_name(name)])
                id2name[aliases[norm_name(name)]] = name
                continue
            row = dp.get(f"{norm_name(name)}|{pos}")
            if row:
                fpids.append(row["fantasypros_id"])
                id2name[row["fantasypros_id"]] = name
            else:
                unmapped.append(f"{name} ({pos})")
    if unmapped:
        print(f"UNMAPPED — not in DP ID map or aliases.csv, will NOT be fetched: {', '.join(unmapped)}")
        print("  -> reconcile by adding rows to aliases.csv (source_name,fpid)")

    print(f"Candidate pool: {len(fpids)} offensive players + top-10 K/DST")

    # Full refetch: every candidate is requested fresh. The existing cache
    # (any age) serves only as a fallback for players the quota cuts off.
    fresh, cached = load_batch_cache()
    fallback = {**cached, **fresh}
    if fallback:
        print(f"{len(fallback)} players held as cache fallback")

    players: dict[int, dict] = {}
    todo = sorted(set(fpids), key=int)
    complete = True
    queries = [
        f"{config.SEASON}/projections?week=0&players={':'.join(todo[i:i + 10])}"
        for i in range(0, len(todo), 10)
    ] + [f"{config.SEASON}/projections?week=0&position={pos}" for pos in ("K", "DST")]
    print(f"Fetching {len(todo)} players + K/DST -> {len(queries)} API requests")
    CANARY = "22968"  # Jahmyr Gibbs — always projected; distinguishes "no
    no_proj: list[str] = []  # projections for this batch" from real quota death
    for n, q in enumerate(queries):
        try:
            resp = api_get(q, key)
        except QuotaExhausted as e:
            canary_ok = False
            if "players=" in q:
                try:
                    canary_ok = bool(api_get(
                        f"{config.SEASON}/projections?week=0&players={CANARY}",
                        key, reuse_secs=0).get("players"))
                except QuotaExhausted:
                    canary_ok = False
            if canary_ok:
                ids = q.split("players=")[1].split(":")
                no_proj.extend(ids)
                print(f"  batch {n + 1}: FP has no projections for: "
                      + ", ".join(id2name.get(i, i) for i in ids))
                continue
            complete = False
            print(f"\nAPI quota exhausted at batch {n + 1}/{len(queries)} ({e}; canary confirmed).")
            print("Compiling partial output. Rerun --full when quota allows: batches fetched")
            print("in the last 30 minutes are served from disk, so it resumes where it died.")
            break
        for p in resp.get("players") or []:
            players[p["fpid"]] = flatten(p)
        time.sleep(1.5)  # stay safely under the ~1 req/sec throttle
    if no_proj:
        print(f"{len(no_proj)} candidates have no FP projection (named above) — excluded, not errors.")

    filled = 0
    for fid, p in fallback.items():
        if fid not in players:
            players[fid] = p
            filled += 1
    if filled:
        print(f"Filled {filled} players from cache fallback (not refreshed this run)")

    if complete:
        # Reconciliation: every requested id must have come back or be
        # explicitly accounted for as projection-less.
        lost = [fid for fid in set(fpids) if int(fid) not in players and fid not in no_proj]
        if lost:
            print(f"RECONCILIATION FAILURE: {len(lost)} requested fpids not returned by the API: "
                  + ", ".join(f"{id2name.get(f, '?')} ({f})" for f in lost)
                  + " — investigate before trusting the output.")
    compile_and_write(players, complete)


if __name__ == "__main__":
    main()

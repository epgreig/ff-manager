"""Blend FantasyPros and WinWithOdds projections into one canonical table.

Pure local step — spends no API quota. Reads:
  fp_cache/fp_projections.csv   (from fetch_fp.py)
  data/wwo.tsv | wwo.html       (WinWithOdds table; optional)
  aliases.csv                   (wwo_name -> fpid overrides, persists forever)

Writes:
  out/blended.csv               (import this into Google Sheets)
  out/unmatched.csv             (both-direction match report, biggest names first)

Usage: python3 blend.py
"""

import json
from collections import defaultdict

import config
from common import (ROOT, find_wwo, load_aliases, norm_name, parse_wwo, player_key,
                    read_csv_dicts, write_csv_dicts)

FP_PATH = ROOT / "fp_cache" / "fp_projections.csv"
ALIAS_PATH = ROOT / "aliases.csv"
OUT_PATH = ROOT / "out" / "blended.csv"
UNMATCHED_PATH = ROOT / "out" / "unmatched.csv"


# Column layout of FantasyPros SITE csv exports (data/FantasyPros_*_{POS}.csv),
# after Player,Team. Duplicate YDS/TDS headers force positional parsing.
FP_CSV_LAYOUT = {
    "QB": ["pass_att", "pass_cmp", "pass_yd", "pass_td", "int", "rush_att", "rush_yd", "rush_td", "fumble"],
    "RB": ["rush_att", "rush_yd", "rush_td", "rec", "rec_yd", "rec_td", "fumble"],
    "WR": ["rec", "rec_yd", "rec_td", "rush_att", "rush_yd", "rush_td", "fumble"],
    "TE": ["rec", "rec_yd", "rec_td", "fumble"],
}


def _f(txt):
    try:
        return float(str(txt).replace(",", ""))
    except ValueError:
        return 0.0


def fp_site_fallback(have_fpids, have_keys, aliases):
    """Fill players missing from the API data using FP website CSV exports.

    Points are computed from component stats under league scoring (never their
    FPTS column, which is used only as a cross-check). Returns (rows, unmapped).
    """
    import csv as _csv

    dp = {}
    dp_path = ROOT / "fp_cache" / "dp_ids.csv"
    if dp_path.exists():
        for r in read_csv_dicts(dp_path):
            if r.get("fantasypros_id", "").isdigit():
                dp[f"{norm_name(r['name'])}|{r['position']}"] = r["fantasypros_id"]

    dp_any = {}
    if dp_path.exists():
        for r in read_csv_dicts(dp_path):
            if r.get("fantasypros_id", "").isdigit():
                dp_any.setdefault(norm_name(r["name"]), r["fantasypros_id"])

    rules = config.SCORING_RULES
    added, unmapped = [], []
    for pos, stat_cols in FP_CSV_LAYOUT.items():
        path = ROOT / "data" / f"FantasyPros_Fantasy_Football_Projections_{pos}.csv"
        if not path.exists():
            continue
        with open(path, newline="") as fh:
            rows = list(_csv.reader(fh))
        n_pos, max_err = 0, 0.0
        for r in rows[1:]:
            if len(r) < len(stat_cols) + 2 or not r[0].strip():
                continue
            name = r[0].strip()
            vals = {c: _f(r[i + 2]) for i, c in enumerate(stat_cols)}
            base = sum(vals.get(c, 0) * rules.get(c, 0) for c in stat_cols if c != "rec")
            rec = vals.get("rec", 0.0)
            std, half, ppr = base, base + 0.5 * rec, base + 1.0 * rec
            max_err = max(max_err, abs(half - _f(r[-1])))
            fpid = aliases.get(norm_name(name)) or dp.get(f"{norm_name(name)}|{pos}")
            if fpid is None:
                unmapped.append(f"{name} ({pos})")
                continue
            if fpid in have_fpids or player_key(name, pos) in have_keys:
                continue  # API data wins
            n_pos += 1
            added.append({"fpid": fpid, "name": name, "position": pos, "team": r[1].strip(),
                          "points_std": round(std, 1), "points_ppr": round(ppr, 1),
                          "points_half": round(half, 1), "from_csv": True})
        if n_pos:
            import datetime
            age = datetime.date.fromtimestamp(path.stat().st_mtime)
            print(f"FP-site CSV fallback: {n_pos} {pos}s filled from {path.name} "
                  f"(downloaded {age}) | max |computed-FPTS| = {max_err:.1f}")
            if max_err > 3:
                print(f"  WARNING: computed points diverge from the file's FPTS by up to "
                      f"{max_err:.1f} — check the download's scoring setting.")
    # K and DST: their CSVs carry no per-distance FG split and only season-total
    # points allowed, so league scoring CANNOT be reconstructed from components.
    # We take FantasyPros' own FPTS and label it, rather than guess.
    for pos in ("K", "DST"):
        path = ROOT / "data" / f"FantasyPros_Fantasy_Football_Projections_{pos}.csv"
        if not path.exists():
            continue
        with open(path, newline="") as fh:
            rows = list(_csv.reader(fh))
        if "FPTS" not in rows[0]:
            print(f"  {path.name}: no FPTS column, skipped")
            continue
        i_fpts = rows[0].index("FPTS")
        n_pos = 0
        for r in rows[1:]:
            if len(r) <= i_fpts or not r[0].strip():
                continue
            name, pts = r[0].strip(), _f(r[i_fpts])
            if pts <= 0:
                continue
            fpid = (aliases.get(norm_name(name)) or dp.get(f"{norm_name(name)}|{pos}")
                    or dp_any.get(norm_name(name))
                    or f"{pos.lower()}-{norm_name(name).replace(' ', '-')}")
            if fpid in have_fpids or player_key(name, pos) in have_keys:
                continue
            n_pos += 1
            added.append({"fpid": fpid, "name": name, "position": pos,
                          "team": r[1].strip() if len(r) > 1 else "",
                          "points_std": pts, "points_ppr": pts, "points_half": pts,
                          "from_csv": True, "fp_default_scoring": True})
        if n_pos:
            print(f"FP-site CSV fallback: {n_pos} {pos}s added from {path.name} "
                  f"-- FANTASYPROS DEFAULT SCORING, not league scoring "
                  f"({'no FG distance split' if pos == 'K' else 'season PA only, no per-game tiers'})")
    return added, unmapped


def load_wwo():
    path = find_wwo(ROOT / "data", config.WWO_FILE)
    if path is None:
        return None
    print(f"WinWithOdds source: {path.name} (points computed under league scoring)")
    return parse_wwo(path, config.WWO_COLS, config.SCORING_RULES)


def write_idp():
    """Rank the FantasyPros IDP CSVs under league IDP scoring -> out/idp.csv."""
    rows = []
    for path in sorted((ROOT / "data").glob("FantasyPros_*_Projections_*.csv")):
        pos = path.stem.rsplit("_", 1)[1]
        if pos not in ("DB", "DL", "LB"):
            continue  # offense/K/DST exports live in data/ too
        for r in read_csv_dicts(path):
            if not r.get("Player", "").strip():  # interleaved high/low rows
                continue
            pts = sum(float(r[c] or 0) * w for c, w in config.IDP_SCORING.items() if c in r)
            rows.append({"name": r["Player"], "position": pos, "team": r.get("Team", ""),
                         "idp_pts": round(pts, 1)})
    # Manually supplied IDP lines (players no source projects — e.g. Travis
    # Hunter, whom FantasyPros classifies WR-only). Same scoring path as the
    # real files; provenance is recorded in the file itself.
    manual = ROOT / "data" / "manual_idp.csv"
    if manual.exists():
        for r in read_csv_dicts(manual):
            if not r.get("Player", "").strip():
                continue
            pts = sum(float(r[c] or 0) * w for c, w in config.IDP_SCORING.items() if c in r)
            # Yahoo counts every stat a player records, whichever slot he fills.
            # For a dual-eligible player the value of using the DB slot is his
            # offence AND his defence, so add his blended offensive total.
            off = 0.0
            if r.get("add_offense", "").strip().lower() in ("y", "yes", "true", "1"):
                for b in (read_csv_dicts(OUT_PATH) if OUT_PATH.exists() else []):
                    if norm_name(b["name"]) == norm_name(r["Player"]):
                        off = float(b["blend_pts"])
                        break
                if not off:
                    print(f"  WARNING: {r['Player']} flagged add_offense but no offensive "
                          f"row found in blended.csv — using defence only.")
            rows.append({"name": r["Player"], "position": r.get("Pos", "DB").strip() or "DB",
                         "team": r.get("Team", ""), "idp_pts": round(pts + off, 1)})
            print(f"IDP MANUAL ESTIMATE (not a sourced projection): {r['Player']} "
                  f"= {pts:.1f} defensive" + (f" + {off:.1f} offensive = {pts + off:.1f}" if off else ""))

    if rows:
        rows.sort(key=lambda r: -r["idp_pts"])
        write_csv_dicts(ROOT / "out" / "idp.csv", rows, ["name", "position", "team", "idp_pts"])
        print(f"IDP: wrote {len(rows)} players to out/idp.csv "
              f"(top: {rows[0]['name']} {rows[0]['position']} {rows[0]['idp_pts']})")


def main():
    if not FP_PATH.exists():
        raise SystemExit("Run fetch_fp.py first — no fp_cache/fp_projections.csv")
    fp = read_csv_dicts(FP_PATH)
    aliases = load_aliases(ALIAS_PATH)

    csv_rows, csv_unmapped = fp_site_fallback(
        {r["fpid"] for r in fp}, {player_key(r["name"], r["position"]) for r in fp}, aliases)
    fp.extend(csv_rows)
    if csv_unmapped:
        print(f"FP-site CSV names not in DP map/aliases (skipped): {', '.join(csv_unmapped)}")

    fp_pts_col = f"points_{config.SCORING}"
    for r in fp:
        r["fp_pts"] = float(r[fp_pts_col] or 0)

    # Bye weeks + FFC ADP, joined by normalized name|pos from the cached FFC feed.
    ffc_info = {}
    ffc_path = ROOT / "fp_cache" / "ffc_adp.json"
    if ffc_path.exists():
        pos_map = {"DEF": "DST", "PK": "K"}
        for p in json.loads(ffc_path.read_text()).get("players", []):
            pos = pos_map.get(p["position"], p["position"])
            ffc_info[player_key(p["name"], pos)] = (p.get("bye", ""), p.get("adp", ""))
    else:
        print("No fp_cache/ffc_adp.json yet (written by fetch_fp.py) — bye/ffc_adp will be blank.")

    wwo = load_wwo()
    if wwo is None:
        print("No data/wwo.* file — writing FantasyPros-only table.")
        wwo = []

    by_key = defaultdict(list)
    by_name = defaultdict(list)
    by_fpid = {}
    for r in fp:
        by_key[player_key(r["name"], r["position"])].append(r)
        by_name[player_key(r["name"], "")].append(r)
        by_fpid[r["fpid"]] = r

    wwo_pts = {}  # fpid -> points
    wwo_delta = {}  # fpid -> 7-day move on WWO's own (full-PPR) scale; staleness indicator
    wwo_unmatched = []
    for w in wwo:
        row = None
        if norm_name(w["name"]) in aliases:
            row = by_fpid.get(aliases[norm_name(w["name"])])
        if row is None:
            hits = by_key.get(player_key(w["name"], w["position"]), [])
            if not hits:
                hits = by_name.get(player_key(w["name"], ""), [])
                hits = hits if len(hits) == 1 else []
            row = hits[0] if hits else None
        if row is not None:
            wwo_pts[row["fpid"]] = w["points"]
            wwo_delta[row["fpid"]] = w.get("delta", 0.0)
        else:
            wwo_unmatched.append(w)

    # Per-position scale diagnostic on matched pairs.
    scale = {}
    if wwo_pts:
        sums = defaultdict(lambda: [0.0, 0.0, 0])
        for r in fp:
            if r["fpid"] in wwo_pts and r["fp_pts"] > 0:
                s = sums[r["position"]]
                s[0] += wwo_pts[r["fpid"]]
                s[1] += r["fp_pts"]
                s[2] += 1
        print(f"\nScale check (scoring={config.SCORING}) — mean WWO/FP ratio by position:")
        for pos, (sw, sf, n) in sorted(sums.items()):
            scale[pos] = sw / sf
            flag = "  <-- level disagreement, check scoring assumptions" if abs(scale[pos] - 1) > 0.08 else ""
            print(f"  {pos:4s} {scale[pos]:.3f}  (n={n}){flag}")

    w_wwo = config.BLEND_WEIGHT_WWO
    out_rows, overridden = [], []
    for r in fp:
        wp = wwo_pts.get(r["fpid"])
        if wp is not None and config.RESCALE_WWO and scale.get(r["position"]):
            wp = wp / scale[r["position"]]
        if wp is None:
            blend, src = r["fp_pts"], "fp_only"
        else:
            blend, src = w_wwo * wp + (1 - w_wwo) * r["fp_pts"], "blend"
        if r.get("from_csv"):
            src += "_csv"  # FP side came from the site download, not the API
        if r.get("fp_default_scoring"):
            src += "_fpdefault"  # points are FantasyPros' scoring, not the league's
        bye, fadp = ffc_info.get(player_key(r["name"], r["position"]), ("", ""))
        # Applied only at write time, so the per-position rescale above still
        # sees him among his real position's players.
        pos_out = config.POSITION_OVERRIDES.get(r["name"].strip().lower(), r["position"])
        if pos_out != r["position"]:
            overridden.append(f"{r['name']} {r['position']} -> {pos_out}")
        out_rows.append({
            "fpid": r["fpid"], "name": r["name"], "position": pos_out, "team": r["team"],
            "bye": bye, "ffc_adp": fadp,
            "fp_pts": round(r["fp_pts"], 1),
            "wwo_pts": "" if wp is None else round(wp, 1),
            "blend_pts": round(blend, 1), "source": src,
            "diff": "" if wp is None else round(wp - r["fp_pts"], 1),
            "wwo_7d_delta": "" if wp is None else wwo_delta.get(r["fpid"], 0.0),
        })
    out_rows.sort(key=lambda r: -r["blend_pts"])
    write_csv_dicts(OUT_PATH, out_rows,
                    ["fpid", "name", "position", "team", "bye", "ffc_adp", "fp_pts", "wwo_pts",
                     "blend_pts", "source", "diff", "wwo_7d_delta"])
    if overridden:
        print("Position overrides (off the offensive board, still in the data): "
              + "; ".join(overridden))
    no_bye = sum(1 for r in out_rows if r["bye"] == "")
    if no_bye:
        print(f"{no_bye} players without bye/ffc_adp (not in FFC's ~230-player feed — fine for deep names)")

    # Unmatched report, biggest projections first. FP side only flags players
    # WWO should plausibly cover (offense with meaningful points).
    report = [
        {"source": "wwo", "name": w["name"], "position": w["position"], "points": w["points"]}
        for w in sorted(wwo_unmatched, key=lambda w: -w["points"])
    ]
    if wwo:
        fp_unmatched = [
            r for r in fp
            if r["fpid"] not in wwo_pts and r["position"] in ("QB", "RB", "WR", "TE") and r["fp_pts"] > 50
        ]
        report += [
            {"source": "fp", "name": r["name"], "position": r["position"], "points": round(r["fp_pts"], 1)}
            for r in sorted(fp_unmatched, key=lambda r: -r["fp_pts"])
        ]
    write_csv_dicts(UNMATCHED_PATH, report, ["source", "name", "position", "points"])

    n_blend = sum(1 for r in out_rows if r["source"].startswith("blend"))
    print(f"\nWrote {len(out_rows)} players to {OUT_PATH} ({n_blend} blended, "
          f"{len(out_rows) - n_blend} single-source)")
    if report:
        print(f"Unmatched report: {len(report)} rows in {UNMATCHED_PATH} — top of list:")
        for r in report[:8]:
            print(f"  [{r['source']}] {r['name']} ({r['position']}) {r['points']}")
        print("Fix important ones by adding rows to aliases.csv (source_name,fpid).")
    write_idp()


if __name__ == "__main__":
    main()

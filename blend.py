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

from collections import defaultdict

import config
from common import (ROOT, find_wwo, load_aliases, norm_name, parse_wwo, player_key,
                    read_csv_dicts, write_csv_dicts)

FP_PATH = ROOT / "fp_cache" / "fp_projections.csv"
ALIAS_PATH = ROOT / "aliases.csv"
OUT_PATH = ROOT / "out" / "blended.csv"
UNMATCHED_PATH = ROOT / "out" / "unmatched.csv"


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
        for r in read_csv_dicts(path):
            if not r.get("Player", "").strip():  # interleaved high/low rows
                continue
            pts = sum(float(r[c] or 0) * w for c, w in config.IDP_SCORING.items() if c in r)
            rows.append({"name": r["Player"], "position": pos, "team": r.get("Team", ""),
                         "idp_pts": round(pts, 1)})
    if rows:
        rows.sort(key=lambda r: -r["idp_pts"])
        write_csv_dicts(ROOT / "out" / "idp.csv", rows, ["name", "position", "team", "idp_pts"])
        print(f"IDP: wrote {len(rows)} players to out/idp.csv "
              f"(top: {rows[0]['name']} {rows[0]['position']} {rows[0]['idp_pts']})")


def main():
    if not FP_PATH.exists():
        raise SystemExit("Run fetch_fp.py first — no fp_cache/fp_projections.csv")
    fp = read_csv_dicts(FP_PATH)
    fp_pts_col = f"points_{config.SCORING}"
    for r in fp:
        r["fp_pts"] = float(r[fp_pts_col] or 0)

    aliases = load_aliases(ALIAS_PATH)

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
    out_rows = []
    for r in fp:
        wp = wwo_pts.get(r["fpid"])
        if wp is not None and config.RESCALE_WWO and scale.get(r["position"]):
            wp = wp / scale[r["position"]]
        if wp is None:
            blend, src = r["fp_pts"], "fp_only"
        else:
            blend, src = w_wwo * wp + (1 - w_wwo) * r["fp_pts"], "blend"
        out_rows.append({
            "fpid": r["fpid"], "name": r["name"], "position": r["position"], "team": r["team"],
            "fp_pts": round(r["fp_pts"], 1),
            "wwo_pts": "" if wp is None else round(wp, 1),
            "blend_pts": round(blend, 1), "source": src,
            "diff": "" if wp is None else round(wp - r["fp_pts"], 1),
            "wwo_7d_delta": "" if wp is None else wwo_delta.get(r["fpid"], 0.0),
        })
    out_rows.sort(key=lambda r: -r["blend_pts"])
    write_csv_dicts(OUT_PATH, out_rows,
                    ["fpid", "name", "position", "team", "fp_pts", "wwo_pts", "blend_pts",
                     "source", "diff", "wwo_7d_delta"])

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

    n_blend = sum(1 for r in out_rows if r["source"] == "blend")
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

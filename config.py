"""Central knobs for the projection pipeline. Edit here, not in the scripts."""

# Scoring format used for the blended output: "std", "ppr", or "half".
# League confirmed: half PPR, 4-pt passing TD, otherwise Yahoo defaults.
SCORING = "half"

# League scoring rules, used to compute points from WWO's component stats
# (their own "Projections" column is full PPR — wrong for this league).
SCORING_RULES = {
    "pass_yd": 0.04, "pass_td": 4.0, "int": -1.0,
    "rush_yd": 0.1, "rush_td": 6.0,
    "rec": 0.5, "rec_yd": 0.1, "rec_td": 6.0,
    "fumble": -2.0,
}

# Source weights where a player is covered. FantasyPros gets whatever is left
# over. WWO sits at 0: investigation on 2026-08-31 showed it is a projection
# model refreshed only on news (98% of players unchanged over three days), not
# the live prop market it advertises. Ciely (The Athletic) replaces it — one
# sharp analyst rather than a consensus, so genuinely diversifying, and his
# 08-31 file already prices news ours does not.
BLEND_WEIGHT_WWO = 0.0
BLEND_WEIGHT_CIELY = 0.35

# Ciely's workbook: 'Ranks w Proj' holds four side-by-side blocks. Bounds are
# explicit half-open column ranges because a scan would run into the next block.
CIELY_FILE = "2026-FFB-Projections-0831.xlsx"
CIELY_SHEET = "Ranks w Proj"
CIELY_BLOCKS = {"QB": (0, 14), "RB": (15, 28), "WR": (29, 41), "TE": (42, 52)}
CIELY_COLS = {"PASS YARDS": "pass_yd", "PASS TD": "pass_td", "INT": "int",
              "RUSH YARDS": "rush_yd", "RUSH TD": "rush_td", "REC": "rec",
              "RECV YARDS": "rec_yd", "RECV TD": "rec_td"}

# If True, rescale WWO points by the per-position mean ratio vs FP before
# blending (corrects level disagreement between sources, keeps player ordering).
# Needed: prop lines are medians and season totals are right-skewed (injury
# truncation), so WWO sits ~15% below FP's means at RB. Without rescaling,
# blended players get dragged down relative to fp_only players.
RESCALE_WWO = True

# Max players to request from the FantasyPros API per position (offense only;
# K/DST come from the free top-10 endpoint and IDPs are handled outside this
# pipeline). Actual counts are limited by the candidate seed (FFC ADP + WWO).
COUNTS = {"QB": 45, "RB": 120, "WR": 140, "TE": 60, "K": 25}

# How long a cached projection counts as current in --cache mode. 24h is fine
# in the quiet weeks; drop it to 6 or less on draft week, when one suspension
# can move a player 80 points between runs.
FRESH_HOURS = 24

# Board position overrides, keyed by lower-case name. A dual-eligible player
# belongs in the slot you would actually start him in; the board has one row per
# player, so putting him in both would double-count him. The overridden position
# matches no panel, so he disappears from the offensive board while his row —
# and his projection — stay in the data. data/manual_idp.csv's add_offense
# lookup matches on NAME, so his offensive points still feed his IDP value.
POSITION_OVERRIDES = {"travis hunter": "DUAL"}

# The WinWithOdds download in data/ (pinned by name — data/ also holds the
# FantasyPros IDP CSVs, so "newest file" discovery is not safe).
WWO_FILE = "season_long_proj_table.csv"

# IDP scoring applied to the FantasyPros IDP CSVs (their FPTS column matches
# no standard scoring). Confirmed from league settings 2026-08-25 — customized
# above Yahoo defaults: sack 3 (not 2), INT 4 (not 3). Safety 4, block kick 2,
# XP returned 4 are also in the league rules but FP doesn't project them.
IDP_SCORING = {
    "TACKLE": 1.0, "ASSIST": 0.5, "SACK": 3.0, "PD": 1.0,
    "INT": 4.0, "FF": 2.0, "FR": 2.0, "TD": 6.0,
}

# WinWithOdds input: data/wwo.tsv (paste the copied table into a text file)
# or data/wwo.html (saved page). Column auto-detection can be overridden here
# with exact header names once we see the real file, e.g. {"name": "Player",
# "position": "Pos", "points": "FPTS"}.
WWO_COLS = {"name": None, "position": None, "points": None}

# FantasyFootballCalculator ADP used for candidate seeding.
FFC_FORMAT = "half-ppr"
LEAGUE_TEAMS = 12
SEASON = 2026

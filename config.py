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

# Weight on WinWithOdds (prop-derived) points where both sources cover a player.
# FantasyPros gets 1 - this.
BLEND_WEIGHT_WWO = 0.65

# If True, rescale WWO points by the per-position mean ratio vs FP before
# blending (corrects level disagreement between sources, keeps player ordering).
# Needed: prop lines are medians and season totals are right-skewed (injury
# truncation), so WWO sits ~15% below FP's means at RB. Without rescaling,
# blended players get dragged down relative to fp_only players.
RESCALE_WWO = True

# Max players to request from the FantasyPros API per position (offense only;
# K/DST come from the free top-10 endpoint and IDPs are handled outside this
# pipeline). Actual counts are limited by the candidate seed (FFC ADP + WWO).
COUNTS = {"QB": 40, "RB": 70, "WR": 100, "TE": 40}

# The WinWithOdds download in data/ (pinned by name — data/ also holds the
# FantasyPros IDP CSVs, so "newest file" discovery is not safe).
WWO_FILE = "season_long_proj_table.csv"

# IDP scoring applied to the FantasyPros IDP CSVs (their FPTS column matches
# no standard scoring). These are Yahoo's defaults from memory — VERIFY against
# the league settings page, especially INT (2 vs 3).
IDP_SCORING = {
    "TACKLE": 1.0, "ASSIST": 0.5, "SACK": 2.0, "PD": 1.0,
    "INT": 3.0, "FF": 2.0, "FR": 2.0, "TD": 6.0,
}

# WinWithOdds input: data/wwo.tsv (paste the copied table into a text file)
# or data/wwo.html (saved page). Column auto-detection can be overridden here
# with exact header names once we see the real file, e.g. {"name": "Player",
# "position": "Pos", "points": "FPTS"}.
WWO_COLS = {"name": None, "position": None, "points": None}

# FantasyFootballCalculator ADP used for candidate seeding.
FFC_FORMAT = "half-ppr"
LEAGUE_TEAMS = 10
SEASON = 2026

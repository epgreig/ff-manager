"""Shared helpers: name normalization, .env loading, WWO table parsing."""

import csv
import io
import re
import urllib.request
from html.parser import HTMLParser
from pathlib import Path

ROOT = Path(__file__).resolve().parent

NAME_SUFFIXES = {"jr", "sr", "ii", "iii", "iv", "v"}


def norm_name(name: str) -> str:
    """Lowercase, strip punctuation and generational suffixes, collapse spaces."""
    s = name.lower()
    s = re.sub(r"[.'’,\-]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    parts = [p for p in s.split() if p not in NAME_SUFFIXES]
    return " ".join(parts)


def player_key(name: str, position: str) -> str:
    return f"{norm_name(name)}|{position.upper()}"


def load_aliases(path: Path = ROOT / "aliases.csv") -> dict[str, str]:
    """source_name -> fpid overrides, keyed by normalized name. Used by both
    the fetcher (candidate mapping) and the blender (WWO matching)."""
    if not path.exists():
        return {}
    with open(path, newline="") as f:
        return {norm_name(r["source_name"]): r["fpid"] for r in csv.DictReader(f) if r.get("fpid")}


def load_env(path: Path = ROOT / ".env") -> dict:
    env = {}
    if path.exists():
        for line in path.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def http_get(url: str, headers: dict | None = None, timeout: int = 30) -> bytes:
    hdrs = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"}
    hdrs.update(headers or {})
    req = urllib.request.Request(url, headers=hdrs)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def read_csv_dicts(path: Path) -> list[dict]:
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def write_csv_dicts(path: Path, rows: list[dict], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


class _TableParser(HTMLParser):
    """Extract rows of the largest <table> in an HTML document."""

    def __init__(self):
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.tables.append([])
        elif tag == "tr" and self.tables:
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None:
            self._row.append(" ".join("".join(self._cell).split()))
            self._cell = None
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.tables[-1].append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _detect_col(headers: list[str], preferred: str | None, needles: list[str]) -> int | None:
    if preferred:
        for i, h in enumerate(headers):
            if h.strip().lower() == preferred.strip().lower():
                return i
        raise SystemExit(f"WWO_COLS names '{preferred}' but headers are: {headers}")
    for needle in needles:
        for i, h in enumerate(headers):
            if needle in h.strip().lower():
                return i
    return None


def find_wwo(data_dir: Path, preferred: str | None = None) -> Path | None:
    """Locate the WinWithOdds download in data/ (pinned name first, then any
    csv/tsv/html/txt that isn't a FantasyPros download)."""
    if not data_dir.exists():
        return None
    if preferred and (data_dir / preferred).exists():
        return data_dir / preferred
    for pattern in ("*.csv", "*.tsv", "*.html", "*.txt"):
        hits = sorted(
            (p for p in data_dir.glob(pattern) if not p.name.startswith("FantasyPros")),
            key=lambda p: -p.stat().st_mtime,
        )
        if hits:
            return hits[0]
    return None


def _num(txt: str) -> float:
    try:
        return float(txt.replace(",", ""))
    except (ValueError, AttributeError):
        return 0.0


# WWO component-column headers -> scoring rule keys.
WWO_COMPONENTS = {
    "Pass Yards": "pass_yd", "Pass TDs": "pass_td", "Ints": "int",
    "Rush Yards": "rush_yd", "Rush TDs": "rush_td",
    "Receptions": "rec", "Rec Yards": "rec_yd", "Rec TDs": "rec_td",
    "Fumbles": "fumble",
}


def parse_wwo(path: Path, col_overrides: dict, scoring_rules: dict | None = None) -> list[dict]:
    """Parse the WinWithOdds table (.csv download, pasted .tsv, or saved .html).

    Returns rows of {name, position, points}. When the file carries component
    stat columns and scoring_rules is given, points are computed from the
    components under league scoring (WWO's own total is full PPR); otherwise
    the site's points column is used as-is. Raises SystemExit with the header
    list if the needed columns can't be auto-detected.
    """
    text = path.read_text(errors="replace")
    if text.lstrip().startswith("<"):
        p = _TableParser()
        p.feed(text)
        if not p.tables:
            raise SystemExit(f"No <table> found in {path}")
        rows = max(p.tables, key=len)
    elif "," in text.splitlines()[0] and "\t" not in text.splitlines()[0]:
        rows = list(csv.reader(io.StringIO(text)))
    else:
        rows = [line.split("\t") for line in text.splitlines() if line.strip()]

    headers = [h.strip() for h in rows[0]]
    comp_idx = {rule: headers.index(h) for h, rule in WWO_COMPONENTS.items() if h in headers}
    use_components = scoring_rules is not None and len(comp_idx) >= 5
    i_name = _detect_col(headers, col_overrides.get("name"), ["player", "name"])
    i_pos = _detect_col(headers, col_overrides.get("position"), ["pos"])
    i_pts = _detect_col(
        headers, col_overrides.get("points"), ["fpts", "proj", "fantasy", "points", "pts"]
    )
    missing = [lbl for lbl, i in [("name", i_name), ("position", i_pos), ("points", i_pts)] if i is None]
    if missing:
        raise SystemExit(
            f"Could not auto-detect column(s) {missing} in {path}.\n"
            f"Headers found: {headers}\n"
            f"Set exact header names in config.WWO_COLS."
        )

    out = []
    for r in rows[1:]:
        if len(r) <= max(i_name, i_pos, i_pts):
            continue
        if use_components:
            pts = sum(_num(r[i]) * scoring_rules[rule] for rule, i in comp_idx.items() if i < len(r))
        else:
            try:
                pts = float(r[i_pts].replace(",", ""))
            except ValueError:
                continue
        if pts <= 0:
            continue
        pos = re.sub(r"\d+$", "", r[i_pos].strip().upper())  # WR3 -> WR
        out.append({"name": r[i_name].strip(), "position": pos, "points": round(pts, 2)})
    if not out:
        raise SystemExit(f"Parsed 0 usable rows from {path} (headers: {headers})")
    return out

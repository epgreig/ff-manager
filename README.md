# ff-manager

Fantasy football draft-prep pipeline: blends FantasyPros consensus projections
(via their public API) with WinWithOdds prop-derived projections into one
canonical table for import into the draft Google Sheet.

## Daily refresh

```
python3 fetch_fp.py    # hits the FP API (~22-35 requests of the 50/day quota)
python3 blend.py       # pure local: match, diagnose, blend
```

Then in Google Sheets: File > Import > Upload `out/blended.csv` > Replace data
at selected cell (on the raw import tab).

After `blend.py`, glance at the top of `out/unmatched.csv`. If an important
player failed to match, add a row to `aliases.csv` (`wwo_name,fpid` — the fpid
is in `out/blended.csv`) and rerun `blend.py`. Aliases persist forever; each
mismatch is fixed exactly once.

## Inputs

- **FantasyPros**: fetched automatically. Free API tier caps responses at 10
  players, so `fetch_fp.py` batches explicit player-ID requests (IDs come from
  the DynastyProcess ID database; candidates seeded from FantasyFootballCalculator
  ADP + the WWO table). Same-day cache prevents accidental quota burn; `--force`
  refetches.
- **WinWithOdds**: manual. Copy the site's table and paste into `data/wwo.tsv`
  (or save the page as `data/wwo.html`). Column headers are auto-detected;
  override in `config.WWO_COLS` if detection fails.

## Knobs

All in `config.py`: scoring format (std/ppr/half — must match the Yahoo league),
WWO blend weight, per-position candidate counts, optional per-position rescaling
of WWO onto FP's level.

## Known data gaps

Rule: missing data is flagged, never filled in with estimates. No number
appears in any output file unless it came from a source.

- **Travis Hunter's defensive stats**: no source projects them (FantasyPros
  classifies him WR-only; the IDP CSVs have no row for him). His DB-slot value
  (offense + defense combined) is therefore UNKNOWN — the pipeline carries only
  his sourced offensive projection. Needs a real defensive projection source
  before any DB-slot comparison is made.

## Notes

- K/DST: fetched as FP's plain top-10 (they're last-round picks; no blending).
- IDP (DL/LB/DB slots): outside this pipeline entirely — grab FP's IDP rankings
  page manually once per season.
- `blend.py` prints a per-position WWO/FP scale ratio; a position far from 1.00
  means the sources disagree on level (likely scoring assumptions) — investigate
  before trusting the blend, or set `RESCALE_WWO = True`.

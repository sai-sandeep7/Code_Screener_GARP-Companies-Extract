"""
GARP Screener Dashboard Generator  —  v3
=========================================
Reads every CONSOLIDATED_ALL_SCREENS_dd_mm_yyyy.xlsx produced by the screener
scraper and renders one self-contained analytics dashboard (dashboard.html).

DESIGN BRIEF
------------
v2 shipped everything the workbook contained. v3 ships only what changes an
investment decision. Cut, and why:

  * raw screen sheets (p1_main, p2_v3 …) — scraper pagination artifacts. A
    company being on "page 2" carries no information; only the screen GROUP
    it belongs to does.
  * 20x20 Jaccard overlap heatmap — interesting, never actionable.
  * A-vs-B set algebra panel — power-user toy; the conviction score answers
    the same question ("who passes many screens?") in one column.
  * P/E decile table, price histograms, aggregate market-cap doughnut —
    describe the dataset, not any position.
  * EPS and standalone Price columns in the screener — price level is not
    comparable across stocks. Both live in the company profile instead.
  * separate History + Appearance leaderboard views — folded into Changes.

Kept, because each is a fact worth showing:

  * Conviction        how many independent screens flag the name (a count)
  * Peer P/E          median P/E of that name's own industry
  * vs Peers          its P/E against that median, as a percentage
  * Earnings yield    100 / P-E, so cheapness compares across names
  * What changed      new entrants, exits, real price moves between runs
  * Industry exposure concentration in the filtered basket

There is deliberately NO composite score, rank or recommendation. Every
column is either a reported figure or a one-step arithmetic transform of one,
so the ranking and the judgement stay with the analyst.

Usage
-----
    python dashboard.py [folder]

Requires pandas + openpyxl. Charts via Chart.js from CDN.
"""

import sys
import os
import re
import glob
import json
import math
import datetime
from collections import defaultdict, Counter

import pandas as pd

try:
    import openpyxl  # noqa: F401
except ModuleNotFoundError:
    openpyxl = None


def ensure_openpyxl_installed():
    if openpyxl is None:
        raise ImportError(
            "openpyxl is required to read Excel files. "
            "Install it using: python -m pip install openpyxl"
        )


# ==========================
# CONFIG
# ==========================
FILE_PATTERN = "CONSOLIDATED_ALL_SCREENS_*.xlsx"
DATE_RE = re.compile(r"(\d{2})_(\d{2})_(\d{4})")

NON_SCREEN_SHEETS = {"Summary", "All Companies", "Numeric Data"}
NON_SCREEN_PREFIXES = ("Broad_", "Industry_")
NON_SCREEN_SUFFIXES = ("_combined",)

# Market-cap bands in Rs Crore
MCAP_BANDS = [
    ("Nano", 0, 500),
    ("Micro", 500, 2000),
    ("Small", 2000, 10000),
    ("Mid", 10000, 50000),
    ("Large", 50000, float("inf")),
]

# Minimum members before an industry's median P/E is trusted as a peer
# benchmark. Below this we fall back to the broad industry, then the universe.
MIN_PEERS = 3

# Data-quality gate on P/E.
#
# Screener output contains reported P/Es that are arithmetically valid but
# analytically meaningless: sub-1 ratios from demerger accounting or one-off
# extraordinary gains (Taparia Tools 0.11, Raymond 0.72), and 400+ ratios from
# near-zero earnings. Left in, they dominate any cheapness ranking — 8 of the
# top 50 scores in this dataset were sub-3 P/E shells.
#
# Names outside the band keep their row and their raw P/E, but are excluded
# from peer medians and from scoring, and are flagged in the UI. That is
# preferable to silently deleting them.
PE_MIN_TRUSTED = 3.0
PE_MAX_TRUSTED = 150.0


def pe_is_trusted(pe):
    return pe is not None and PE_MIN_TRUSTED <= pe <= PE_MAX_TRUSTED


# ==========================
# DISCOVERY & PARSING
# ==========================
def find_run_files(folder):
    runs = []
    for p in glob.glob(os.path.join(folder, FILE_PATTERN)):
        m = DATE_RE.search(os.path.basename(p))
        if not m:
            continue
        dd, mm, yyyy = m.groups()
        try:
            d = datetime.date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        runs.append((d, p))
    runs.sort(key=lambda x: x[0])
    return runs


def is_screen_sheet(name, group_sheet_names=None):
    if name in NON_SCREEN_SHEETS:
        return False
    if name.startswith(NON_SCREEN_PREFIXES):
        return False
    if name.endswith(NON_SCREEN_SUFFIXES):
        return False
    if group_sheet_names and name in group_sheet_names:
        return False
    if name.endswith("...") and "comb" in name.lower():
        return False
    return True


def parse_numeric(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        try:
            if pd.isna(val):
                return None
        except TypeError:
            pass
        if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
            return None
        return float(val)
    s = str(val).strip().replace(",", "").replace("₹", "")
    if s in ("", "-", "None", "nan", "NaN"):
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return None if (math.isnan(f) or math.isinf(f)) else f


def clean_str(v, default="Unclassified"):
    if v is None:
        return default
    s = str(v).strip()
    return default if (s == "" or s.lower() == "nan") else s


def load_run(path, run_date):
    """
    Parse one workbook into {companies, groups}.

    Individual screen sheets are read only to backfill companies missing from
    "All Companies" and to rebuild group membership when "Appears In" is
    absent. Their page-level membership is intentionally NOT carried into the
    dashboard — see the module docstring.
    """
    ensure_openpyxl_installed()
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheet_names = xls.sheet_names

    companies = {}
    groups = defaultdict(set)

    group_sheet_names = set()
    if "Summary" in sheet_names:
        try:
            sdf = xls.parse("Summary")
            if {"Screen Group", "Type"} <= set(sdf.columns):
                group_sheet_names = set(
                    sdf[sdf["Type"] == "Screen Group"]["Screen Group"].astype(str)
                )
        except Exception:
            pass

    if "All Companies" in sheet_names:
        for _, row in xls.parse("All Companies").iterrows():
            name = clean_str(row.get("Company Name"), "")
            if not name:
                continue
            glist = [g.strip() for g in str(row.get("Appears In") or "").split(",") if g.strip()]
            companies[name] = {
                "broad_industry": clean_str(row.get("Broad Industry")),
                "industry": clean_str(row.get("Industry")),
                "market_cap": parse_numeric(row.get("Market Cap (₹ Cr)")),
                "pe": parse_numeric(row.get("Stock P/E")),
                "price": parse_numeric(row.get("Current Price (₹)")),
                "appears_in": glist,
            }
            for g in glist:
                groups[g].add(name)

    if "Numeric Data" in sheet_names:
        try:
            for _, row in xls.parse("Numeric Data").iterrows():
                name = clean_str(row.get("Company Name"), "")
                if not name:
                    continue
                c = companies.setdefault(name, {
                    "broad_industry": clean_str(row.get("Broad Industry")),
                    "industry": clean_str(row.get("Industry")),
                    "market_cap": None, "pe": None, "price": None, "appears_in": [],
                })
                for field, col in (("market_cap", "Market Cap (Cr) - Numeric"),
                                   ("pe", "Stock P/E - Numeric"),
                                   ("price", "Current Price (₹) - Numeric")):
                    v = parse_numeric(row.get(col))
                    if v is not None:
                        c[field] = v
        except Exception:
            pass

    screen_sheet_count = 0
    for sn in sheet_names:
        if not is_screen_sheet(sn, group_sheet_names):
            continue
        try:
            df = xls.parse(sn)
        except Exception:
            continue
        if "Company Name" not in df.columns:
            continue
        screen_sheet_count += 1
        for _, row in df.iterrows():
            name = clean_str(row.get("Company Name"), "")
            if not name or name in companies:
                continue
            companies[name] = {
                "broad_industry": clean_str(row.get("Broad Industry")),
                "industry": clean_str(row.get("Industry")),
                "market_cap": parse_numeric(row.get("Market Cap (₹ Cr)")),
                "pe": parse_numeric(row.get("Stock P/E")),
                "price": parse_numeric(row.get("Current Price (₹)")),
                "appears_in": [],
            }

    if not groups:  # "Appears In" missing — rebuild from *_combined sheets
        for sn in sheet_names:
            if not (sn.endswith("_combined") or (sn.endswith("...") and "comb" in sn.lower())):
                continue
            try:
                df = xls.parse(sn)
            except Exception:
                continue
            if "Company Name" not in df.columns:
                continue
            label = sn.replace("_combined", "").rstrip(". ")
            for n in df["Company Name"].dropna().tolist():
                nm = clean_str(n, "")
                if nm:
                    groups[label].add(nm)
                    if nm in companies and label not in companies[nm]["appears_in"]:
                        companies[nm]["appears_in"].append(label)

    return {
        "date": run_date,
        "path": path,
        "companies": companies,
        "groups": dict(groups),
        "screen_sheet_count": screen_sheet_count,
    }


def load_all_runs(folder):
    runs = []
    for d, p in find_run_files(folder):
        try:
            runs.append(load_run(p, d))
            print(f"  ✓ {os.path.basename(p)}")
        except Exception as e:
            print(f"  ⚠ Skipping {os.path.basename(p)}: {e}")
    return runs


# ==========================
# ANALYTICS
# ==========================
def band_of(value):
    if value is None:
        return "Unknown"
    for label, lo, hi in MCAP_BANDS:
        if lo <= value < hi:
            return label
    return MCAP_BANDS[-1][0]


def median(vals):
    v = sorted(x for x in vals if x is not None)
    if not v:
        return None
    m = len(v) // 2
    return v[m] if len(v) % 2 else (v[m - 1] + v[m]) / 2


def stats_for(values):
    vals = sorted(v for v in values if v is not None)
    if not vals:
        return None
    s = pd.Series(vals)
    return {
        "count": int(s.count()),
        "min": round(float(s.min()), 2),
        "median": round(float(s.median()), 2),
        "max": round(float(s.max()), 2),
    }


def peer_benchmarks(companies):
    """
    Median P/E per specific industry, per broad industry, and universe-wide.
    Only trusted P/Es contribute, so one 0.11 shell cannot drag an industry
    benchmark down and make every peer look expensive.
    """
    by_ind, by_broad, all_pe = defaultdict(list), defaultdict(list), []
    for c in companies.values():
        if not pe_is_trusted(c["pe"]):
            continue
        by_ind[c["industry"]].append(c["pe"])
        by_broad[c["broad_industry"]].append(c["pe"])
        all_pe.append(c["pe"])
    return (
        {k: median(v) for k, v in by_ind.items() if len(v) >= MIN_PEERS},
        {k: median(v) for k, v in by_broad.items() if len(v) >= MIN_PEERS},
        median(all_pe),
    )


def enrich(companies):
    """
    Attach displayed fields to every company. No ranking, no weighting:
      ey     earnings yield %   = 100 / P-E
      bench  peer P/E           = median P/E of its industry
      disc   vs peers %         = (bench - pe) / bench * 100
      flag   why a P/E is not usable, or None
    """
    ind_med, broad_med, uni_med = peer_benchmarks(companies)
    rows = {}
    for name, c in companies.items():
        pe = c["pe"]
        bench = ind_med.get(c["industry"]) or broad_med.get(c["broad_industry"]) or uni_med
        bench_src = ("industry" if c["industry"] in ind_med else
                     "broad" if c["broad_industry"] in broad_med else "universe")

        # Flag, don't delete. A flagged name still appears in the table with
        # its raw P/E, but carries no discount, no yield and no score, so it
        # cannot win a ranking on the strength of a broken input.
        if pe is None:
            flag = "no P/E reported"
        elif pe < PE_MIN_TRUSTED:
            flag = f"P/E {pe:g} below {PE_MIN_TRUSTED:g} — likely one-off gain or demerger artifact"
        elif pe > PE_MAX_TRUSTED:
            flag = f"P/E {pe:g} above {PE_MAX_TRUSTED:g} — near-zero earnings"
        else:
            flag = None

        trusted = flag is None
        disc = round((bench - pe) / bench * 100.0, 1) if (trusted and bench) else None
        rows[name] = {
            **c,
            "ey": round(100.0 / pe, 2) if trusted else None,
            "bench": round(bench, 2) if bench else None,
            "bench_src": bench_src,
            "disc": disc,
            "flag": flag,
            "band": band_of(c["market_cap"]),
            "cv": len(c.get("appears_in") or []),
        }

    return rows


def movers(prev_run, curr_run, limit=25):
    """Largest P/E and market-cap moves among names present in both runs."""
    if prev_run is None:
        return {"pe": [], "mcap": []}
    out = {"pe": [], "mcap": []}
    for name, cur in curr_run["companies"].items():
        old = prev_run["companies"].get(name)
        if not old:
            continue
        for key, bucket in (("pe", "pe"), ("market_cap", "mcap")):
            a, b = old.get(key), cur.get(key)
            if a is None or b is None or a == 0:
                continue
            out[bucket].append({
                "name": name, "from": a, "to": b,
                "pct": round((b - a) / abs(a) * 100.0, 2),
            })
    for k in out:
        out[k].sort(key=lambda x: -abs(x["pct"]))
        out[k] = out[k][:limit]
    return out


def freshness(prev_run, curr_run):
    """Flags a run whose numbers are byte-identical to the previous one."""
    if prev_run is None:
        return None
    common = set(curr_run["companies"]) & set(prev_run["companies"])
    if not common:
        return None
    unchanged = sum(
        1 for n in common
        if (prev_run["companies"][n]["market_cap"], prev_run["companies"][n]["pe"],
            prev_run["companies"][n]["price"]) ==
           (curr_run["companies"][n]["market_cap"], curr_run["companies"][n]["pe"],
            curr_run["companies"][n]["price"])
    )
    share = unchanged / len(common)
    return {
        "compared": len(common), "unchanged": unchanged,
        "unchanged_pct": round(share * 100, 1), "stale": share >= 0.90,
        "prev_date": prev_run["date"].isoformat(),
    }


def group_stats(run, enriched):
    out = {}
    for gn, names in run["groups"].items():
        infos = [enriched[n] for n in names if n in enriched]
        if not infos:
            continue
        out[gn] = {
            "count": len(infos),
            "pe": median([i["pe"] for i in infos]),
            "mcap": median([i["market_cap"] for i in infos]),
            "disc": median([i["disc"] for i in infos]),
        }
    return dict(sorted(out.items(), key=lambda x: -x[1]["count"]))


def always_present(runs):
    if not runs:
        return []
    inter = set(runs[0]["companies"])
    for r in runs[1:]:
        inter &= set(r["companies"])
    return sorted(inter)


def company_history(runs):
    hist = defaultdict(lambda: {"appearances": 0, "dates": [], "pe": [], "mcap": [],
                                "price": [], "cv": []})
    for r in runs:
        ds = r["date"].isoformat()
        for name, info in r["companies"].items():
            h = hist[name]
            h["appearances"] += 1
            h["dates"].append(ds)
            h["pe"].append(info["pe"])
            h["mcap"].append(info["market_cap"])
            h["price"].append(info["price"])
            h["cv"].append(len(info.get("appears_in") or []))
    return dict(hist)


def median_history(runs):
    out = {"dates": [], "pe": [], "count": []}
    for r in runs:
        out["dates"].append(r["date"].isoformat())
        out["pe"].append(median([c["pe"] for c in r["companies"].values()]))
        out["count"].append(len(r["companies"]))
    return out


# ==========================
# BUILD DATASET
# ==========================
def build_dataset(folder):
    runs = load_all_runs(folder)
    if not runs:
        return None

    dataset = {
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "run_dates": [r["date"].isoformat() for r in runs],
        "bands": [b[0] for b in MCAP_BANDS],
        "pe_band": [PE_MIN_TRUSTED, PE_MAX_TRUSTED],
        "always_present": always_present(runs),
        "median_history": median_history(runs),
        "company_history": company_history(runs),
        "runs": [],
    }

    prev = None
    for r in runs:
        enriched = enrich(r["companies"])
        prev_names = set(prev["companies"]) if prev else set()
        curr_names = set(r["companies"])

        rows = []
        for name in sorted(curr_names):
            e = enriched[name]
            rows.append({
                "n": name, "b": e["broad_industry"], "i": e["industry"],
                "mc": e["market_cap"], "pe": e["pe"], "pr": e["price"],
                "ey": e["ey"], "disc": e["disc"], "bench": e["bench"],
                "bsrc": e["bench_src"], "band": e["band"], "flag": e["flag"],
                "cv": e["cv"], "g": e.get("appears_in") or [],
                "new": bool(prev) and name not in prev_names,
            })

        pes = [c["pe"] for c in r["companies"].values()]
        dataset["runs"].append({
            "date": r["date"].isoformat(),
            "total": len(curr_names),
            "rows": rows,
            "groups": {gn: sorted(v) for gn, v in r["groups"].items()},
            "group_stats": group_stats(r, enriched),
            "new_vs_prev": sorted(curr_names - prev_names) if prev else [],
            "dropped_vs_prev": sorted(prev_names - curr_names) if prev else [],
            "dropped_info": ({n: {"b": prev["companies"][n]["broad_industry"],
                                  "mc": prev["companies"][n]["market_cap"],
                                  "pe": prev["companies"][n]["pe"]}
                              for n in sorted(prev_names - curr_names)} if prev else {}),
            "movers": movers(prev, r),
            "freshness": freshness(prev, r),
            "pe_stats": stats_for(pes),
            "mcap_stats": stats_for([c["market_cap"] for c in r["companies"].values()]),
            "band_mix": Counter(band_of(c["market_cap"]) for c in r["companies"].values()),
            "conv_mix": dict(sorted(Counter(
                len(c.get("appears_in") or []) for c in r["companies"].values()).items())),
        })
        prev = r

    return dataset


# ==========================
# CSS
# ==========================
CSS = r"""
:root{
  --bg:#f6f8fc; --card:#fff; --soft:#fafbfe; --soft2:#f1f4fa;
  --line:#ebeff7; --line2:#dfe6f2;
  --ink:#0f1729; --ink2:#3b465e; --muted:#78849c; --dim:#9aa5bb;
  --blue:#3b6ef5; --blue-s:#eaf0ff; --indigo:#6c5ce7; --indigo-s:#f0edff;
  --green:#0fa968; --green-s:#e4f7ef; --red:#e5384f; --red-s:#fdeaed;
  --amber:#f0a020; --amber-s:#fdf3e2; --teal:#0aa2b8; --teal-s:#e3f6f9;
  --r:14px;
  --sh:0 1px 2px rgba(15,23,41,.04),0 4px 16px rgba(15,23,41,.04);
  --sh-lg:0 14px 40px rgba(15,23,41,.12);
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
  background:var(--bg);color:var(--ink);font-size:14px;line-height:1.5;-webkit-font-smoothing:antialiased}
::-webkit-scrollbar{width:9px;height:9px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:#d5dcec;border-radius:8px;border:2px solid transparent;background-clip:content-box}
::-webkit-scrollbar-thumb:hover{background:#bfc8de;background-clip:content-box}

/* shell */
.app{display:flex;min-height:100vh}
aside{width:228px;flex:0 0 228px;background:#0f1729;padding:22px 14px;position:sticky;top:0;
  height:100vh;display:flex;flex-direction:column;gap:3px;overflow:auto}
.brand{display:flex;align-items:center;gap:11px;padding:2px 8px 24px}
.brand .logo{width:34px;height:34px;border-radius:10px;flex:0 0 34px;
  background:linear-gradient(140deg,var(--blue),var(--indigo));display:grid;place-items:center;
  font-weight:800;font-size:15px;color:#fff}
.brand .t1{font-weight:700;font-size:14.5px;color:#fff;letter-spacing:-.2px}
.brand .t2{font-size:11px;color:#7b87a3;margin-top:1px}
.navsec{font-size:9.5px;font-weight:700;letter-spacing:.11em;color:#5d6a86;
  text-transform:uppercase;padding:18px 10px 8px}
.navbtn{display:flex;align-items:center;gap:11px;width:100%;text-align:left;cursor:pointer;
  background:transparent;border:0;color:#a8b3c9;padding:10px 12px;border-radius:10px;
  font-size:13.5px;font-weight:500;transition:.15s;font-family:inherit}
.navbtn:hover{background:rgba(255,255,255,.06);color:#fff}
.navbtn.active{background:var(--blue);color:#fff;font-weight:600;
  box-shadow:0 6px 18px rgba(59,110,245,.35)}
.navbtn .ic{width:17px;text-align:center;font-size:13px}
.navbtn .badge{margin-left:auto;font-size:10.5px;font-weight:600;background:rgba(255,255,255,.12);
  padding:1px 8px;border-radius:99px}
.navbtn.active .badge{background:rgba(255,255,255,.25)}
.sidefoot{margin-top:auto;padding:14px 10px 0;border-top:1px solid rgba(255,255,255,.08);
  font-size:11px;color:#6b7791;line-height:1.6}
.sidefoot b{color:#c3ccdd;display:block;font-size:11.5px}

main{flex:1;min-width:0}
.topbar{position:sticky;top:0;z-index:40;background:rgba(246,248,252,.9);backdrop-filter:blur(12px);
  padding:20px 26px 0}
.topline{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.topline h1{margin:0;font-size:22px;font-weight:700;letter-spacing:-.5px}
.topline .sub{font-size:12.5px;color:var(--muted);margin-top:2px}
.spacer{flex:1}
.searchbox{position:relative}
.searchbox .mag{position:absolute;left:13px;top:50%;transform:translateY(-50%);color:var(--dim);font-size:13px}
.searchbox input{padding-left:34px;min-width:250px}
.iconbtn{width:38px;height:38px;border-radius:10px;background:var(--card);border:1px solid var(--line2);
  display:grid;place-items:center;cursor:pointer;color:var(--ink2);box-shadow:var(--sh);
  transition:.15s;font-size:14px;position:relative}
.iconbtn:hover{color:var(--blue);border-color:#c6d4f7}
.iconbtn .n{position:absolute;top:-5px;right:-5px;background:var(--blue);color:#fff;
  font-size:9.5px;font-weight:700;border-radius:99px;padding:1px 5px}
.content{padding:18px 26px 60px}

/* controls */
input,select,button{font-family:inherit;font-size:13px}
input[type=text],input[type=number],select{background:var(--card);color:var(--ink);
  border:1px solid var(--line2);border-radius:10px;padding:8px 11px;outline:none;
  transition:.15s;box-shadow:var(--sh)}
input::placeholder{color:var(--dim)}
input:focus,select:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(59,110,245,.13)}
input[type=number]{width:80px}
.range{display:flex;gap:6px}
.btn{background:var(--card);border:1px solid var(--line2);color:var(--ink2);padding:8px 14px;
  border-radius:10px;cursor:pointer;transition:.15s;font-weight:500;box-shadow:var(--sh)}
.btn:hover{border-color:#c6d4f7;color:var(--ink)}
.btn.primary{background:var(--blue);border-color:var(--blue);color:#fff;
  box-shadow:0 6px 16px rgba(59,110,245,.3)}
.btn.primary:hover{background:#2f5fe0;color:#fff}
.btn.sm{padding:5px 10px;font-size:12px;border-radius:8px}
.fgroup{display:flex;flex-direction:column;gap:5px}
.fgroup>label{font-size:10.5px;color:var(--muted);font-weight:600}
.drawer{display:none;background:var(--card);border:1px solid var(--line);border-radius:var(--r);
  padding:16px 18px;margin-top:14px;box-shadow:var(--sh)}
.drawer.open{display:block}
.drawer .row{display:grid;grid-template-columns:repeat(auto-fit,minmax(172px,1fr));gap:13px}
.drawer .dhead{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}
.drawer .dhead h4{margin:0;font-size:13.5px;font-weight:650}
.chip{display:inline-flex;align-items:center;gap:6px;padding:5px 12px;border-radius:99px;
  background:var(--card);border:1px solid var(--line2);font-size:12px;cursor:pointer;
  transition:.15s;user-select:none;color:var(--ink2);font-weight:500}
.chip:hover{border-color:#c6d4f7}
.chip.on{background:var(--blue-s);border-color:#c2d3fb;color:var(--blue);font-weight:600}
.chip.x{background:var(--blue);border-color:var(--blue);color:#fff}
.chipbar{display:flex;gap:7px;flex-wrap:wrap;margin-top:12px;align-items:center}
.chipbar .ttl{font-size:11px;color:var(--muted);font-weight:600}
.multi{position:relative}
.multi>.btn{width:100%;display:flex;justify-content:space-between;gap:8px;align-items:center;font-weight:500}
.multi .pop{display:none;position:absolute;z-index:70;top:calc(100% + 6px);left:0;width:290px;
  max-height:320px;overflow:auto;background:var(--card);border:1px solid var(--line2);
  border-radius:12px;padding:9px;box-shadow:var(--sh-lg)}
.multi.open .pop{display:block}
.multi .pop .opt{display:flex;align-items:center;gap:9px;padding:6px 8px;border-radius:7px;
  cursor:pointer;font-size:12.5px}
.multi .pop .opt:hover{background:var(--soft2)}
.multi .pop .opt .cnt{margin-left:auto;color:var(--dim);font-size:11px}
.multi .pop .tools{display:flex;gap:6px;margin-bottom:7px}
.multi .pop input{width:100%;margin-bottom:7px}

/* tabs */
.tabs{display:flex;gap:4px;margin:18px 0 0;border-bottom:1px solid var(--line2)}
.tab{background:transparent;border:0;padding:10px 18px;cursor:pointer;font-size:13.5px;
  font-weight:550;color:var(--muted);border-bottom:2px solid transparent;transition:.15s;
  font-family:inherit;margin-bottom:-1px}
.tab:hover{color:var(--ink)}
.tab.active{color:var(--blue);border-bottom-color:var(--blue);font-weight:650}

/* KPI cards */
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(232px,1fr));gap:16px;margin-bottom:18px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:18px 20px;
  box-shadow:var(--sh);display:flex;align-items:center;gap:15px;transition:.18s;position:relative;
  overflow:hidden}
.kpi:hover{box-shadow:var(--sh-lg);transform:translateY(-1px)}
.kpi::before{content:"";position:absolute;left:0;top:14px;bottom:14px;width:3px;border-radius:0 3px 3px 0;
  background:var(--blue)}
.kpi.green::before{background:var(--green)}.kpi.red::before{background:var(--red)}
.kpi.amber::before{background:var(--amber)}.kpi.indigo::before{background:var(--indigo)}
.kpi.teal::before{background:var(--teal)}
.kpi .ic{width:46px;height:46px;border-radius:50%;flex:0 0 46px;display:grid;place-items:center;
  font-size:18px;background:var(--blue-s);color:var(--blue)}
.kpi.green .ic{background:var(--green-s);color:var(--green)}
.kpi.red .ic{background:var(--red-s);color:var(--red)}
.kpi.amber .ic{background:var(--amber-s);color:var(--amber)}
.kpi.indigo .ic{background:var(--indigo-s);color:var(--indigo)}
.kpi.teal .ic{background:var(--teal-s);color:var(--teal)}
.kpi .val{font-size:24px;font-weight:750;letter-spacing:-.7px;font-variant-numeric:tabular-nums;
  line-height:1.15;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.kpi .lbl{font-size:12.5px;color:var(--muted);font-weight:500;margin-top:2px}
.kpi .sub{font-size:11.5px;color:var(--muted);margin-top:3px}

/* panels */
.panel{background:var(--card);border:1px solid var(--line);border-radius:var(--r);
  padding:20px 22px;margin-bottom:18px;box-shadow:var(--sh)}
.panel>h3{margin:0 0 16px;font-size:14.5px;font-weight:650;display:flex;align-items:center;
  gap:9px;flex-wrap:wrap}
.panel>h3 .hint{font-weight:400;font-size:11.5px;color:var(--dim)}
.two{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.two.wide{grid-template-columns:1.35fr 1fr}
@media(max-width:1200px){.two,.two.wide{grid-template-columns:1fr}}
.flex-between{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px}
.scroll{max-height:430px;overflow:auto}
.scroll.tall{max-height:620px}
.chart-wrap{position:relative;height:280px}
.chart-wrap.sm{height:210px}

/* tables */
table{width:100%;border-collapse:separate;border-spacing:0;font-size:13px}
thead th{position:sticky;top:0;z-index:2;background:var(--soft);color:var(--muted);font-weight:600;
  font-size:11px;letter-spacing:.02em;padding:11px 12px;text-align:left;
  border-bottom:1px solid var(--line2);white-space:nowrap}
thead th.sortable{cursor:pointer;user-select:none}
thead th.sortable:hover{color:var(--blue)}
thead th .arr{opacity:.4;margin-left:4px;font-size:9px}
tbody td{padding:11px 12px;border-bottom:1px solid var(--line);white-space:nowrap;color:var(--ink2)}
tbody tr:last-child td{border-bottom:0}
tbody tr:hover td{background:var(--soft)}
td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}
td.name{font-weight:600;color:var(--ink);max-width:280px;overflow:hidden;text-overflow:ellipsis}
.empty{padding:34px;text-align:center;color:var(--dim)}
.clickable{cursor:pointer;color:var(--ink);font-weight:600}
.clickable:hover{color:var(--blue)}
.sublabel{display:block;font-size:11px;color:var(--dim);font-weight:400;margin-top:1px}

/* bits */
.tag{display:inline-block;padding:3px 10px;border-radius:99px;font-size:11px;font-weight:600}
.tag.new{background:var(--green-s);color:var(--green)}
.tag.out{background:var(--red-s);color:var(--red)}
.tag.g{background:var(--soft2);color:var(--ink2);margin:1px 4px 1px 0;font-weight:500;font-size:10.5px}
.pos{color:var(--green);font-weight:650}.neg{color:var(--red);font-weight:650}
.pill{display:inline-block;min-width:26px;text-align:center;padding:2px 9px;border-radius:99px;
  font-size:11.5px;font-weight:700;background:var(--soft2);color:var(--ink2)}
.pill.hot{background:var(--amber-s);color:var(--amber)}
.flagged{color:var(--amber);font-weight:650;cursor:help}
.pill.fire{background:var(--green-s);color:var(--green)}
.bar{height:7px;border-radius:99px;background:var(--soft2);overflow:hidden;min-width:64px}
.bar>i{display:block;height:100%;border-radius:99px;background:var(--blue)}
.pager{display:flex;gap:8px;align-items:center;justify-content:flex-end;margin-top:14px;
  color:var(--muted);font-size:12.5px}
.kv{display:grid;grid-template-columns:auto 1fr;gap:10px 18px;font-size:13px}
.kv .k{color:var(--muted)}
.legend{display:flex;flex-direction:column;gap:11px;justify-content:center}
.legend .li{display:flex;align-items:center;gap:9px;font-size:12.5px}
.legend .dot{width:9px;height:9px;border-radius:50%;flex:0 0 9px}
.legend .lv{margin-left:auto;font-weight:700;font-variant-numeric:tabular-nums}
.banner{display:flex;gap:12px;align-items:flex-start;background:#fff8ec;border:1px solid #f5e0b8;
  border-left:4px solid var(--amber);border-radius:12px;padding:14px 18px;margin-bottom:18px;
  font-size:13px;color:#7a5717;box-shadow:var(--sh)}
.banner b{display:block;color:#5c3f0c;margin-bottom:2px}
.toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:var(--ink);
  color:#fff;padding:11px 20px;border-radius:11px;box-shadow:var(--sh-lg);z-index:200;
  opacity:0;transition:.25s;pointer-events:none;font-size:13px}
.toast.show{opacity:1}
.view{display:none;animation:fade .2s ease}
.view.active{display:block}
@keyframes fade{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:none}}
"""


# ==========================
# RENDER
# ==========================
def render_html(dataset):
    latest = dataset["runs"][-1]
    payload = json.dumps(dataset, default=str, allow_nan=False, separators=(",", ":"))
    payload = payload.replace("</", "<\\/")

    tpl = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GARP Screener</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>__CSS__</style>
</head>
<body>
<div class="app">
  <aside>
    <div class="brand">
      <div class="logo">G</div>
      <div><div class="t1">GARP Screener</div><div class="t2">Equity research desk</div></div>
    </div>
    <div class="navsec">Analysis</div>
    <button class="navbtn active" data-view="overview"><span class="ic">◎</span>Overview</button>
    <button class="navbtn" data-view="ideas"><span class="ic">▤</span>Screener<span class="badge" id="nb-ideas"></span></button>
    <button class="navbtn" data-view="changes"><span class="ic">⇄</span>What Changed</button>
    <button class="navbtn" data-view="company"><span class="ic">☰</span>Company</button>
    <div class="sidefoot">
      <b>Run __LATEST__</b>
      __NRUNS__ workbook(s)<br>Built __GENERATED__
    </div>
  </aside>

  <main>
    <div class="topbar">
      <div class="topline">
        <div>
          <h1 id="view-title">Overview</h1>
          <div class="sub" id="view-sub"></div>
        </div>
        <span class="spacer"></span>
        <div class="searchbox"><span class="mag">🔍</span>
          <input type="text" id="f-q" placeholder="Search company or industry…"></div>
        <select id="f-run" title="Run date"></select>
        <button class="iconbtn" id="btn-filters" title="Filters">⚙</button>
        <button class="btn primary" id="btn-export">⤓ Export</button>
      </div>

      <div class="drawer" id="drawer">
        <div class="dhead"><h4>Refine universe</h4>
          <div><button class="btn sm" id="btn-reset">Reset</button>
               <button class="btn sm" id="btn-closef">Done</button></div></div>
        <div class="row">
          <div class="fgroup"><label>Broad industry</label><div class="multi" id="f-broad"></div></div>
          <div class="fgroup"><label>Industry</label><div class="multi" id="f-ind"></div></div>
          <div class="fgroup"><label>Screen group</label><div class="multi" id="f-grp"></div></div>
          <div class="fgroup"><label>Group match</label>
            <select id="f-grpmode"><option value="any">ANY of</option><option value="all">ALL of</option><option value="none">NONE of</option></select></div>
          <div class="fgroup"><label>Min conviction</label><select id="f-conv"></select></div>
          <div class="fgroup"><label>Market cap ₹Cr</label>
            <span class="range"><input type="number" id="f-mcmin" placeholder="min"><input type="number" id="f-mcmax" placeholder="max"></span></div>
          <div class="fgroup"><label>P/E</label>
            <span class="range"><input type="number" id="f-pemin" placeholder="min"><input type="number" id="f-pemax" placeholder="max"></span></div>
          <div class="fgroup"><label>Min discount to peers %</label>
            <span class="range"><input type="number" id="f-dmin" placeholder="min"></span></div>
          <div class="fgroup"><label>Min earnings yield %</label>
            <span class="range"><input type="number" id="f-eymin" placeholder="min %"></span></div>
        </div>
        <div class="chipbar" id="f-quick"><span class="ttl">Presets</span></div>
      </div>
      <div class="chipbar" id="active-chips"></div>
    </div>

    <div class="content">
      <div id="overview" class="view active"></div>
      <div id="ideas" class="view"></div>
      <div id="changes" class="view"></div>
      <div id="company" class="view"></div>
    </div>
  </main>
</div>
<div class="toast" id="toast"></div>
<script id="payload" type="application/json">__PAYLOAD__</script>
<script>const DATA=JSON.parse(document.getElementById('payload').textContent);</script>
<script>__JS__</script>
</body>
</html>
"""
    return (tpl.replace("__CSS__", CSS).replace("__JS__", APP_JS)
            .replace("__LATEST__", latest["date"])
            .replace("__NRUNS__", str(len(dataset["runs"])))
            .replace("__GENERATED__", dataset["generated_at"].replace("T", " "))
            .replace("__PAYLOAD__", payload))


APP_JS = r"""
const C={blue:'#3b6ef5',indigo:'#6c5ce7',green:'#0fa968',red:'#e5384f',amber:'#f0a020',
         teal:'#0aa2b8',muted:'#78849c',grid:'#ebeff7'};
const PALETTE=[C.blue,C.indigo,C.green,C.amber,C.teal,C.red,'#8ea9f9','#a99af0','#6dcaa4'];

Chart.defaults.color=C.muted;
Chart.defaults.borderColor=C.grid;
Chart.defaults.font.family="Inter, system-ui, sans-serif";
Chart.defaults.font.size=11;
Chart.defaults.maintainAspectRatio=false;
Chart.defaults.plugins.legend.display=false;
Chart.defaults.plugins.tooltip.backgroundColor='#0f1729';
Chart.defaults.plugins.tooltip.padding=10;
Chart.defaults.plugins.tooltip.cornerRadius=8;
Chart.defaults.plugins.tooltip.displayColors=false;
Chart.defaults.elements.bar.borderRadius=6;
Chart.defaults.elements.bar.borderSkipped=false;
Chart.defaults.elements.point.radius=0;
Chart.defaults.elements.point.hoverRadius=5;
Chart.defaults.elements.line.borderWidth=2.5;
Chart.defaults.scale.grid.drawTicks=false;
Chart.defaults.scale.border={display:false};
Chart.defaults.scale.ticks.padding=8;

const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
function el(t,c,h){const e=document.createElement(t);if(c)e.className=c;if(h!==undefined)e.innerHTML=h;return e}
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function n(v,d=2){if(v==null||isNaN(v))return '—';return Number(v).toLocaleString('en-IN',{maximumFractionDigits:d})}
function cr(v){if(v==null||isNaN(v))return '—';
  if(v>=100000)return '₹'+(v/100000).toFixed(2)+'L Cr';
  if(v>=1000)return '₹'+(v/1000).toFixed(2)+'k Cr';
  return '₹'+n(v,0)+' Cr'}
function pctHtml(v,d=1){if(v==null||isNaN(v))return '—';
  return `<span class="${v>=0?'pos':'neg'}">${v>=0?'▲':'▼'} ${Math.abs(v).toFixed(d)}%</span>`}
function med(a){const v=a.filter(x=>x!=null&&!isNaN(x)).sort((x,y)=>x-y);if(!v.length)return null;
  const m=Math.floor(v.length/2);return v.length%2?v[m]:(v[m-1]+v[m])/2}
function sum(a){return a.reduce((s,x)=>s+(x||0),0)}
function toast(m){const t=$('#toast');t.textContent=m;t.classList.add('show');
  clearTimeout(t._h);t._h=setTimeout(()=>t.classList.remove('show'),1800)}
const charts={};
function chart(k,ctx,cfg){if(charts[k])charts[k].destroy();charts[k]=new Chart(ctx,cfg);return charts[k]}

function kpi(label,val,sub,cls,icon){
  const c=el('div','kpi'+(cls?' '+cls:''));
  c.appendChild(el('div','ic',icon||'◆'));
  const b=el('div');b.style.minWidth='0';
  b.appendChild(el('div','val',val));
  b.appendChild(el('div','lbl',esc(label)));
  if(sub)b.appendChild(el('div','sub',sub));
  c.appendChild(b);return c;
}
function convPill(v){return `<span class="pill ${v>=8?'fire':v>=5?'hot':''}">${v}</span>`}

/* ---------- state ---------- */
const RUNS=DATA.runs;
let RUN=RUNS[RUNS.length-1];
const PREV=()=>{const i=RUNS.indexOf(RUN);return i>0?RUNS[i-1]:null};
const F={q:'',broad:new Set(),ind:new Set(),grp:new Set(),grpmode:'any',conv:0,
  mcmin:null,mcmax:null,pemin:null,pemax:null,dmin:null,eymin:null,bands:new Set(),
  newonly:false,hideflag:false};
let VIEW='overview',SORT={key:'cv',dir:-1},PAGE=0;
const PAGE_SIZE=50;

function matches(r){
  if(F.q){const q=F.q.toLowerCase();
    if(!(r.n.toLowerCase().includes(q)||r.i.toLowerCase().includes(q)||r.b.toLowerCase().includes(q)))return false}
  if(F.broad.size&&!F.broad.has(r.b))return false;
  if(F.ind.size&&!F.ind.has(r.i))return false;
  if(F.grp.size){const gs=new Set(r.g),sel=[...F.grp];
    if(F.grpmode==='any'&&!sel.some(g=>gs.has(g)))return false;
    if(F.grpmode==='all'&&!sel.every(g=>gs.has(g)))return false;
    if(F.grpmode==='none'&&sel.some(g=>gs.has(g)))return false}
  if(F.conv&&r.cv<F.conv)return false;
  if(F.bands.size&&!F.bands.has(r.band))return false;
  if(F.newonly&&!r.new)return false;
  if(F.hideflag&&r.flag)return false;
  if(F.mcmin!=null&&(r.mc==null||r.mc<F.mcmin))return false;
  if(F.mcmax!=null&&(r.mc==null||r.mc>F.mcmax))return false;
  if(F.pemin!=null&&(r.pe==null||r.pe<F.pemin))return false;
  if(F.pemax!=null&&(r.pe==null||r.pe>F.pemax))return false;
  if(F.dmin!=null&&(r.disc==null||r.disc<F.dmin))return false;
  if(F.eymin!=null&&(r.ey==null||r.ey<F.eymin))return false;
  return true;
}
const filtered=()=>RUN.rows.filter(matches);
function activeCount(){
  return (F.q?1:0)+F.broad.size+F.ind.size+F.grp.size+F.bands.size+(F.conv?1:0)+(F.newonly?1:0)+
    ['mcmin','mcmax','pemin','pemax','dmin','eymin'].filter(k=>F[k]!=null).length;
}

/* ---------- multiselect ---------- */
function multiSelect(host,title,options,store,onChange){
  host.innerHTML='';
  const btn=el('button','btn');host.appendChild(btn);
  const pop=el('div','pop');host.appendChild(pop);
  const search=el('input');search.type='text';search.placeholder='filter…';pop.appendChild(search);
  const tools=el('div','tools');
  const bNone=el('button','btn sm','Clear');tools.appendChild(bNone);pop.appendChild(tools);
  const list=el('div');pop.appendChild(list);
  function label(){btn.innerHTML=`<span>${store.size?store.size+' selected':title}</span><span style="color:var(--dim)">▾</span>`}
  function draw(){
    const q=search.value.toLowerCase();list.innerHTML='';
    options.filter(o=>o[0].toLowerCase().includes(q)).slice(0,300).forEach(([name,count])=>{
      const o=el('div','opt');const cb=el('input');cb.type='checkbox';cb.checked=store.has(name);
      o.appendChild(cb);o.appendChild(el('span','',esc(name)));o.appendChild(el('span','cnt',count));
      o.addEventListener('click',ev=>{if(ev.target!==cb)cb.checked=!cb.checked;
        cb.checked?store.add(name):store.delete(name);label();onChange()});
      list.appendChild(o)})}
  btn.addEventListener('click',e=>{e.stopPropagation();
    $$('.multi').forEach(m=>{if(m!==host)m.classList.remove('open')});host.classList.toggle('open')});
  pop.addEventListener('click',e=>e.stopPropagation());
  search.addEventListener('input',draw);
  bNone.addEventListener('click',()=>{store.clear();draw();label();onChange()});
  label();draw();
  return {refresh:()=>{label();draw()}};
}
document.addEventListener('click',()=>$$('.multi').forEach(m=>m.classList.remove('open')));

/* ---------- filter bar ---------- */
const MS={};
function buildFilters(){
  const rs=$('#f-run');rs.innerHTML='';
  RUNS.forEach(r=>{const o=el('option','',r.date);o.value=r.date;rs.appendChild(o)});
  rs.value=RUN.date;
  rs.addEventListener('change',()=>{RUN=RUNS.find(r=>r.date===rs.value);rebuild();render()});

  const cs=$('#f-conv');cs.innerHTML='';
  const mx=Math.max(1,...RUN.rows.map(r=>r.cv));
  for(let i=0;i<=mx;i++){const o=el('option','',i===0?'any':'≥ '+i);o.value=i;cs.appendChild(o)}
  cs.addEventListener('change',()=>{F.conv=+cs.value;PAGE=0;render()});

  rebuild();
  $('#f-grpmode').addEventListener('change',e=>{F.grpmode=e.target.value;render()});
  const bind=(sel,key,num)=>{const i=$(sel);i.addEventListener('input',()=>{
    const v=i.value.trim();F[key]=v===''?(num?null:''):(num?parseFloat(v):v);
    if(num&&isNaN(F[key]))F[key]=null;PAGE=0;render()})};
  bind('#f-q','q',false);
  [['#f-mcmin','mcmin'],['#f-mcmax','mcmax'],['#f-pemin','pemin'],['#f-pemax','pemax'],
   ['#f-dmin','dmin'],['#f-eymin','eymin']].forEach(([s,k])=>bind(s,k,true));

  const qb=$('#f-quick');
  DATA.bands.forEach(b=>{
    const c=el('div','chip',esc(b));
    c.addEventListener('click',()=>{F.bands.has(b)?F.bands.delete(b):F.bands.add(b);
      c.classList.toggle('on');PAGE=0;render()});qb.appendChild(c)});
  const preset=(label,fn)=>{const c=el('div','chip',label);
    c.addEventListener('click',()=>{fn(c);PAGE=0;render()});qb.appendChild(c);return c};
  preset('High conviction ≥5',c=>{F.conv=F.conv>=5?0:5;$('#f-conv').value=F.conv;c.classList.toggle('on')});
  preset('Cheaper than peers',c=>{F.dmin=F.dmin===0?null:0;$('#f-dmin').value=F.dmin==null?'':0;c.classList.toggle('on')});
  preset('Investable ≥ ₹500 Cr',c=>{F.mcmin=F.mcmin===500?null:500;
    $('#f-mcmin').value=F.mcmin==null?'':500;c.classList.toggle('on')});
  preset('Hide flagged P/E',c=>{F.hideflag=!F.hideflag;c.classList.toggle('on')});

  $('#btn-filters').addEventListener('click',e=>{e.stopPropagation();$('#drawer').classList.toggle('open')});
  $('#btn-closef').addEventListener('click',()=>$('#drawer').classList.remove('open'));
  $('#btn-reset').addEventListener('click',reset);
  $('#btn-export').addEventListener('click',exportCSV);
}
function rebuild(){
  const cb=new Map(),ci=new Map(),cg=new Map();
  RUN.rows.forEach(r=>{cb.set(r.b,(cb.get(r.b)||0)+1);ci.set(r.i,(ci.get(r.i)||0)+1);
    r.g.forEach(g=>cg.set(g,(cg.get(g)||0)+1))});
  const srt=m=>[...m.entries()].sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
  MS.broad=multiSelect($('#f-broad'),'All',srt(cb),F.broad,()=>{PAGE=0;render()});
  MS.ind=multiSelect($('#f-ind'),'All',srt(ci),F.ind,()=>{PAGE=0;render()});
  MS.grp=multiSelect($('#f-grp'),'All',srt(cg),F.grp,()=>{PAGE=0;render()});
}
function reset(){
  F.q='';F.broad.clear();F.ind.clear();F.grp.clear();F.bands.clear();F.grpmode='any';
  F.conv=0;F.newonly=false;F.hideflag=false;
  ['mcmin','mcmax','pemin','pemax','dmin','eymin'].forEach(k=>F[k]=null);
  $('#f-q').value='';['#f-mcmin','#f-mcmax','#f-pemin','#f-pemax','#f-dmin','#f-eymin'].forEach(s=>$(s).value='');
  $('#f-grpmode').value='any';$('#f-conv').value=0;
  $$('#f-quick .chip').forEach(c=>c.classList.remove('on'));
  Object.values(MS).forEach(m=>m&&m.refresh());
  PAGE=0;render();toast('Filters cleared');
}
function renderChips(){
  const host=$('#active-chips');host.innerHTML='';
  const items=[];const add=(t,u)=>items.push([t,u]);
  if(F.q)add(`"${F.q}"`,()=>{F.q='';$('#f-q').value=''});
  F.broad.forEach(v=>add(v,()=>{F.broad.delete(v);MS.broad.refresh()}));
  F.ind.forEach(v=>add(v,()=>{F.ind.delete(v);MS.ind.refresh()}));
  F.grp.forEach(v=>add(`${F.grpmode}: ${v}`,()=>{F.grp.delete(v);MS.grp.refresh()}));
  F.bands.forEach(v=>add(v+' cap',()=>{F.bands.delete(v);
    $$('#f-quick .chip').forEach(c=>{if(c.textContent===v)c.classList.remove('on')})}));
  if(F.conv)add(`Conviction ≥${F.conv}`,()=>{F.conv=0;$('#f-conv').value=0});
  if(F.newonly)add('New entrants only',()=>{F.newonly=false});
  if(F.hideflag)add('Flagged P/E hidden',()=>{F.hideflag=false;
    $$('#f-quick .chip').forEach(c=>{if(c.textContent==='Hide flagged P/E')c.classList.remove('on')})});
  const rng=(a,b,l,fa,fb)=>{if(F[a]!=null||F[b]!=null)
    add(`${l} ${F[a]!=null?F[a]:'…'}–${F[b]!=null?F[b]:'…'}`,
      ()=>{F[a]=null;F[b]=null;$(fa).value='';if(fb)$(fb).value=''})};
  rng('mcmin','mcmax','Cap','#f-mcmin','#f-mcmax');
  rng('pemin','pemax','P/E','#f-pemin','#f-pemax');
  if(F.dmin!=null)add(`Discount ≥${F.dmin}%`,()=>{F.dmin=null;$('#f-dmin').value=''});
  if(F.eymin!=null)add(`E.Yield ≥${F.eymin}%`,()=>{F.eymin=null;$('#f-eymin').value=''});
  if(!items.length)return;
  host.appendChild(el('span','ttl','Filters'));
  items.forEach(([t,u])=>{const c=el('div','chip x',`${esc(t)} <span>✕</span>`);
    c.addEventListener('click',()=>{u();PAGE=0;render()});host.appendChild(c)});
  const cl=el('div','chip','Clear all');cl.addEventListener('click',reset);host.appendChild(cl);
}
function exportCSV(){
  const rows=sortRows(filtered());
  const head=['Company','Broad Industry','Industry','Market Cap (Cr)','Price','P/E','Peer P/E',
              'vs Peers %','Earnings Yield %','Screens Matched','Screen Groups'];
  const body=rows.map(r=>[r.n,r.b,r.i,r.mc,r.pr,r.pe,r.bench,r.disc,r.ey,r.cv,r.g.join(' | ')]);
  const csv=[head,...body].map(l=>l.map(v=>{const s=v==null?'':String(v);
    return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s}).join(',')).join('\n');
  const a=el('a');a.href=URL.createObjectURL(new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8'}));
  a.download=`garp_${RUN.date}_${rows.length}.csv`;a.click();
  toast(`Exported ${rows.length} rows`);
}

/* ---------- table ---------- */
const COLS=[
  {k:'n',t:'Company',cls:'name',fmt:r=>
    `<span class="clickable" data-co="${esc(r.n)}">${esc(r.n)}</span>
     <span class="sublabel">${esc(r.i)}</span>`},
  {k:'mc',t:'Market Cap',num:1,fmt:r=>`${cr(r.mc)}<span class="sublabel">${esc(r.band)}</span>`},
  {k:'pr',t:'Price ₹',num:1,fmt:r=>n(r.pr,2)},
  {k:'pe',t:'P/E',num:1,fmt:r=>r.flag
     ? `<span class="flagged" title="${esc(r.flag)}">${n(r.pe,2)} ⚠</span>`
     : n(r.pe,1)},
  {k:'bench',t:'Peer P/E',num:1,fmt:r=>r.bench==null?'—'
     :`${n(r.bench,1)}<span class="sublabel">${esc(r.bsrc)}</span>`},
  {k:'disc',t:'vs Peers',num:1,fmt:r=>r.disc==null?'—':pctHtml(r.disc)},
  {k:'ey',t:'Earn. Yield',num:1,fmt:r=>r.ey==null?'—':n(r.ey,1)+'%'},
  {k:'cv',t:'Screens',num:1,fmt:r=>convPill(r.cv)},
];
function sortRows(rows){
  const k=SORT.key,d=SORT.dir;
  return [...rows].sort((a,b)=>{
    let x=a[k],y=b[k];
    if(x==null&&y==null)return 0;
    if(x==null)return 1;
    if(y==null)return -1;
    if(typeof x==='string')return d*x.localeCompare(y);
    return d*(x-y);
  });
}
function renderTable(host,rows,paged=true){
  const sorted=sortRows(rows),total=sorted.length;
  const pages=Math.max(1,Math.ceil(total/PAGE_SIZE));
  if(PAGE>=pages)PAGE=0;
  const slice=paged?sorted.slice(PAGE*PAGE_SIZE,(PAGE+1)*PAGE_SIZE):sorted;
  let h='<div class="scroll tall"><table><thead><tr>';
  COLS.forEach(c=>{const on=SORT.key===c.k;
    h+=`<th class="sortable ${c.num?'num':''}" data-k="${c.k}">${c.t}<span class="arr">${on?(SORT.dir<0?'▼':'▲'):'⇅'}</span></th>`});
  h+='</tr></thead><tbody>';
  if(!slice.length)h+=`<tr><td colspan="${COLS.length}" class="empty">No companies match these filters.</td></tr>`;
  slice.forEach(r=>{h+='<tr>'+COLS.map(c=>`<td class="${c.cls||''} ${c.num?'num':''}">${c.fmt(r)}</td>`).join('')+'</tr>'});
  h+='</tbody></table></div>';
  if(paged&&total>PAGE_SIZE)
    h+=`<div class="pager"><span>${PAGE*PAGE_SIZE+1}–${Math.min(total,(PAGE+1)*PAGE_SIZE)} of ${total}</span>
      <button class="btn sm" data-pg="prev">‹ Prev</button><span>${PAGE+1} / ${pages}</span>
      <button class="btn sm" data-pg="next">Next ›</button></div>`;
  host.innerHTML=h;
  host.querySelectorAll('th.sortable').forEach(th=>th.addEventListener('click',()=>{
    const k=th.dataset.k;
    if(SORT.key===k)SORT.dir*=-1;else{SORT.key=k;SORT.dir=(k==='n')?1:-1}
    PAGE=0;render()}));
  host.querySelectorAll('[data-pg]').forEach(b=>b.addEventListener('click',()=>{
    PAGE=b.dataset.pg==='prev'?Math.max(0,PAGE-1):Math.min(pages-1,PAGE+1);render()}));
  wireLinks(host);
}
function miniTable(headers,rows,cls='scroll'){
  let h=`<div class="${cls}"><table><thead><tr>`+
    headers.map(x=>`<th class="${x.startsWith('#')||/Cap|P\/E|%|Count|Screens|Price|Yield/.test(x)?'num':''}">${x.replace(/^#/,'')}</th>`).join('')+
    '</tr></thead><tbody>';
  if(!rows.length)h+=`<tr><td colspan="${headers.length}" class="empty">Nothing here.</td></tr>`;
  rows.forEach(r=>{h+='<tr>'+r.map(c=>`<td class="${typeof c==='number'?'num':''}">${c}</td>`).join('')+'</tr>'});
  return h+'</tbody></table></div>';
}
function wireLinks(host){
  host.querySelectorAll('[data-co]').forEach(a=>a.addEventListener('click',e=>{
    e.preventDefault();openCompany(a.dataset.co)}));
}

/* ---------- nav ---------- */
const TITLES={overview:['Overview','Universe health and where the opportunity sits'],
  ideas:['Screener','Filter, sort and export — no ranking applied'],
  changes:['What Changed','New entrants, exits and real price moves'],
  company:['Company','Single-name profile and peer context']};
$$('.navbtn').forEach(b=>b.addEventListener('click',()=>{
  $$('.navbtn').forEach(x=>x.classList.remove('active'));b.classList.add('active');
  $$('.view').forEach(v=>v.classList.remove('active'));
  VIEW=b.dataset.view;$('#'+VIEW).classList.add('active');
  $('#view-title').textContent=TITLES[VIEW][0];$('#view-sub').textContent=TITLES[VIEW][1];
  render()}));
const goto=v=>{const b=$(`.navbtn[data-view="${v}"]`);if(b)b.click()};

/* ---------- dispatch ---------- */
function render(){
  const rows=filtered();
  $('#nb-ideas').textContent=rows.length;
  const nf=activeCount();
  $('#btn-filters').innerHTML=nf?`⚙<span class="n">${nf}</span>`:'⚙';
  renderChips();
  ({overview:renderOverview,ideas:renderIdeas,changes:renderChanges,company:renderCompany}[VIEW])(rows);
}
function staleBanner(){
  const f=RUN.freshness;
  if(!f||!f.stale)return null;
  return el('div','banner',`<div>⚠</div><div><b>Prices may be stale</b>
    ${f.unchanged_pct}% of names are identical to the ${f.prev_date} run. If the market
    was open between runs, check that the scraper's metrics cache expired.</div>`);
}

/* =====================================================================
   OVERVIEW
   ===================================================================== */
function renderOverview(rows){
  const root=$('#overview');root.innerHTML='';
  const sb=staleBanner();if(sb)root.appendChild(sb);
  const prev=PREV();

  const cheap=rows.filter(r=>r.disc!=null&&r.disc>0).length;
  const g=el('div','grid');
  g.appendChild(kpi('Stocks in view',n(rows.length,0),
    RUN.total===rows.length?'full universe':`of ${n(RUN.total,0)} screened`,'','▤'));
  g.appendChild(kpi('Median P/E',n(med(rows.map(r=>r.pe)),1),
    `universe ${n(RUN.pe_stats?RUN.pe_stats.median:null,1)}`,'amber','₹'));
  g.appendChild(kpi('Below peer P/E',n(cheap,0),
    rows.length?`${(cheap/rows.length*100).toFixed(0)}% of the basket`:'','green','▼'));
  const flagged=rows.filter(r=>r.flag).length;
  g.appendChild(kpi('In 5+ screens',n(rows.filter(r=>r.cv>=5).length,0),
    'flagged by 5 or more','indigo','★'));
  g.appendChild(kpi('Suspect P/E',n(flagged,0),
    `outside ${DATA.pe_band[0]}–${DATA.pe_band[1]} — shown, not compared`,'amber','⚠'));
  if(prev){
    g.appendChild(kpi('New entrants',n(RUN.new_vs_prev.length,0),`since ${prev.date}`,'teal','↗'));
    g.appendChild(kpi('Exits',n(RUN.dropped_vs_prev.length,0),`since ${prev.date}`,'red','↘'));
  }else{
    g.appendChild(kpi('Median market cap',cr(med(rows.map(r=>r.mc))),'filtered basket','teal','◧'));
  }
  root.appendChild(g);

  // top ideas + conviction donut
  const two=el('div','two wide');
  const p1=el('div','panel');
  p1.appendChild(el('h3','','Most-screened names <span class="hint">ordered by how many screens flag them — a count, not a ranking</span>'));
  const top=[...rows].sort((a,b)=>b.cv-a.cv||(b.mc||0)-(a.mc||0)).slice(0,10);
  p1.insertAdjacentHTML('beforeend',miniTable(
    ['Company','Market Cap','#P/E','#Peer P/E','#vs Peers','#Screens'],
    top.map(r=>[`<span class="clickable" data-co="${esc(r.n)}">${esc(r.n)}</span>
        <span class="sublabel">${esc(r.i)}</span>`,
      cr(r.mc),n(r.pe,1),n(r.bench,1),pctHtml(r.disc),convPill(r.cv)])));
  two.appendChild(p1);

  const p2=el('div','panel');
  p2.appendChild(el('h3','','Conviction spread'));
  const inner=el('div');inner.style.cssText='display:grid;grid-template-columns:1fr 150px;gap:16px;align-items:center';
  const cwrap=el('div','chart-wrap sm');const cv=el('canvas');cwrap.appendChild(cv);inner.appendChild(cwrap);
  const buckets=[['1–2 screens',r=>r.cv<=2],['3–4 screens',r=>r.cv>=3&&r.cv<=4],
                 ['5–7 screens',r=>r.cv>=5&&r.cv<=7],['8+ screens',r=>r.cv>=8]];
  const bvals=buckets.map(([,f])=>rows.filter(f).length);
  const bcolors=['#c9d6f7',C.blue,C.indigo,C.green];
  const lg=el('div','legend');
  buckets.forEach(([lbl],i)=>{lg.insertAdjacentHTML('beforeend',
    `<div class="li"><span class="dot" style="background:${bcolors[i]}"></span>
     <span>${lbl}</span><span class="lv">${bvals[i]}</span></div>`)});
  inner.appendChild(lg);p2.appendChild(inner);
  two.appendChild(p2);
  root.appendChild(two);
  setTimeout(()=>chart('conv',cv,{type:'doughnut',
    data:{labels:buckets.map(b=>b[0]),datasets:[{data:bvals,backgroundColor:bcolors,
      borderColor:'#fff',borderWidth:3,hoverOffset:6}]},
    options:{cutout:'70%',plugins:{tooltip:{enabled:true}}}}),0);

  // industry exposure + screen groups
  const two2=el('div','two');
  const p3=el('div','panel');
  p3.appendChild(el('h3','','Industry exposure <span class="hint">concentration risk in this basket</span>'));
  const bm=new Map();
  rows.forEach(r=>{if(!bm.has(r.b))bm.set(r.b,[]);bm.get(r.b).push(r)});
  const bl=[...bm.entries()].map(([k,v])=>({k,n:v.length,pe:med(v.map(x=>x.pe)),
    d:med(v.map(x=>x.disc))})).sort((a,b)=>b.n-a.n).slice(0,10);
  const mx=bl.length?bl[0].n:1;
  p3.insertAdjacentHTML('beforeend',miniTable(['Industry','#Count','','#Median P/E','#vs Peers'],
    bl.map(x=>[`<span class="clickable" data-b="${esc(x.k)}">${esc(x.k)}</span>`,x.n,
      `<div class="bar"><i style="width:${x.n/mx*100}%"></i></div>`,n(x.pe,1),pctHtml(x.d)])));
  two2.appendChild(p3);

  const p4=el('div','panel');
  p4.appendChild(el('h3','','Screen groups <span class="hint">click to filter</span>'));
  const inF=new Set(rows.map(r=>r.n));
  const gl=Object.entries(RUN.group_stats).map(([k,v])=>
    [k,v.count,Object.values(RUN.groups[k]||[]).filter(x=>inF.has(x)).length,v.pe,v.disc]);
  p4.insertAdjacentHTML('beforeend',miniTable(['Screen group','#Names','#In view','#Median P/E','#vs Peers'],
    gl.map(r=>[`<span class="clickable" data-g="${esc(r[0])}">${esc(r[0])}</span>`,
      r[1],r[2],n(r[3],1),pctHtml(r[4])]),'scroll'));
  two2.appendChild(p4);
  root.appendChild(two2);

  if(RUNS.length>1){
    const p5=el('div','panel');
    p5.appendChild(el('h3','','Universe and valuation over time'));
    const cw=el('div','chart-wrap');const c2=el('canvas');cw.appendChild(c2);p5.appendChild(cw);
    root.appendChild(p5);
    const mh=DATA.median_history;
    chart('hist',c2,{data:{labels:mh.dates,datasets:[
      {type:'line',label:'Companies',data:mh.count,borderColor:C.blue,
        backgroundColor:'rgba(59,110,245,.10)',fill:true,tension:.35,yAxisID:'y',pointRadius:4},
      {type:'line',label:'Median P/E',data:mh.pe,borderColor:C.amber,tension:.35,yAxisID:'y1',pointRadius:4}]},
      options:{plugins:{legend:{display:true,labels:{usePointStyle:true,pointStyle:'circle',boxWidth:8}},
        tooltip:{enabled:true}},
        scales:{x:{grid:{display:false}},y:{position:'left'},
                y1:{position:'right',grid:{drawOnChartArea:false}}}}});
  }

  wireLinks(root);
  root.querySelectorAll('[data-b]').forEach(a=>a.addEventListener('click',()=>{
    F.broad.clear();F.broad.add(a.dataset.b);MS.broad.refresh();PAGE=0;goto('ideas')}));
  root.querySelectorAll('[data-g]').forEach(a=>a.addEventListener('click',()=>{
    F.grp.clear();F.grp.add(a.dataset.g);F.grpmode='any';$('#f-grpmode').value='any';
    MS.grp.refresh();PAGE=0;goto('ideas')}));
}

/* =====================================================================
   IDEAS
   ===================================================================== */
function renderIdeas(rows){
  const root=$('#ideas');root.innerHTML='';
  const g=el('div','grid');
  g.appendChild(kpi('Matches',n(rows.length,0),'','','▤'));
  g.appendChild(kpi('Median P/E',n(med(rows.map(r=>r.pe)),1),'','amber','₹'));
  g.appendChild(kpi('Median discount',n(med(rows.map(r=>r.disc)),1)+'%','vs industry peers','green','▼'));
  g.appendChild(kpi('Median earnings yield',n(med(rows.map(r=>r.ey)),1)+'%','100 / P-E','indigo','%'));
  root.appendChild(g);

  const p=el('div','panel');
  const head=el('div','flex-between');
  head.appendChild(el('h3','','Ranked shortlist <span class="hint">click a company for its profile</span>'));
  const quick=el('div');
  [['Most screens',{key:'cv',dir:-1}],['Cheapest vs peers',{key:'disc',dir:-1}],
   ['Lowest P/E',{key:'pe',dir:1}],['Largest cap',{key:'mc',dir:-1}]].forEach(([t,s])=>{
    const b=el('button','btn sm',t);b.style.marginLeft='6px';
    b.addEventListener('click',()=>{SORT=s;PAGE=0;render()});quick.appendChild(b)});
  head.appendChild(quick);p.appendChild(head);
  const w=el('div');p.appendChild(w);root.appendChild(p);
  renderTable(w,rows);
}

/* =====================================================================
   WHAT CHANGED
   ===================================================================== */
function renderChanges(rows){
  const root=$('#changes');root.innerHTML='';
  const prev=PREV();
  const sb=staleBanner();if(sb)root.appendChild(sb);
  if(!prev){
    root.appendChild(el('div','panel',
      `<div class="empty">Only one run is loaded.<br><br>
       New entrants, exits and price moves appear once a second dated workbook
       sits in this folder — re-run the scraper tomorrow.</div>`));
    return;
  }
  const byName=Object.fromEntries(RUN.rows.map(r=>[r.n,r]));
  const g=el('div','grid');
  g.appendChild(kpi('New entrants',n(RUN.new_vs_prev.length,0),`since ${prev.date}`,'green','↗'));
  g.appendChild(kpi('Exits',n(RUN.dropped_vs_prev.length,0),'no longer screening','red','↘'));
  g.appendChild(kpi('Held',n(RUN.total-RUN.new_vs_prev.length,0),'present in both runs','','●'));
  g.appendChild(kpi('Churn',((RUN.new_vs_prev.length+RUN.dropped_vs_prev.length)/
    Math.max(1,RUN.total)*100).toFixed(1)+'%','of the universe','amber','⇄'));
  root.appendChild(g);

  const two=el('div','two');
  const np=el('div','panel');
  np.appendChild(el('h3','',`<span class="tag new">NEW</span> Entered this run <span class="hint">research candidates</span>`));
  const news=RUN.new_vs_prev.map(x=>byName[x]).filter(Boolean).sort((a,b)=>b.cv-a.cv||(b.mc||0)-(a.mc||0));
  np.insertAdjacentHTML('beforeend',miniTable(['Company','Market Cap','#P/E','#vs Peers','#Screens'],
    news.map(r=>[`<span class="clickable" data-co="${esc(r.n)}">${esc(r.n)}</span>
      <span class="sublabel">${esc(r.i)}</span>`,cr(r.mc),n(r.pe,1),pctHtml(r.disc),convPill(r.cv)]),'scroll tall'));
  two.appendChild(np);

  const dp=el('div','panel');
  dp.appendChild(el('h3','',`<span class="tag out">EXIT</span> Left this run <span class="hint">review any you hold</span>`));
  dp.insertAdjacentHTML('beforeend',miniTable(['Company','Industry','Last Cap','#Last P/E'],
    RUN.dropped_vs_prev.map(x=>{const d=RUN.dropped_info[x]||{};
      return [esc(x),esc(d.b||'—'),cr(d.mc),n(d.pe,1)]}),'scroll tall'));
  two.appendChild(dp);
  root.appendChild(two);

  const two2=el('div','two');
  [['P/E movers','pe',1],['Market-cap movers','mcap',0]].forEach(([t,k,d])=>{
    const p=el('div','panel');p.appendChild(el('h3','',t+` <span class="hint">largest moves vs ${prev.date}</span>`));
    p.insertAdjacentHTML('beforeend',miniTable(['Company','From','To','#Change'],
      (RUN.movers[k]||[]).map(m=>[
        `<span class="clickable" data-co="${esc(m.name)}">${esc(m.name)}</span>`,
        k==='mcap'?cr(m.from):n(m.from,d),k==='mcap'?cr(m.to):n(m.to,d),pctHtml(m.pct)]),'scroll tall'));
    two2.appendChild(p)});
  root.appendChild(two2);
  wireLinks(root);
}

/* =====================================================================
   COMPANY
   ===================================================================== */
let SELECTED=null;
function openCompany(name){SELECTED=name;goto('company')}
function renderCompany(rows){
  const root=$('#company');root.innerHTML='';
  const H=DATA.company_history;
  const byName=Object.fromEntries(RUN.rows.map(r=>[r.n,r]));

  const p0=el('div','panel');
  const head=el('div','flex-between');
  head.appendChild(el('h3','','Select a company <span class="hint">respects active filters</span>'));
  const inp=el('input');inp.type='text';inp.placeholder='Type to search…';inp.style.minWidth='250px';
  head.appendChild(inp);p0.appendChild(head);
  const lw=el('div');p0.appendChild(lw);root.appendChild(p0);
  const detail=el('div');root.appendChild(detail);

  function drawList(){
    const q=inp.value.toLowerCase();
    const list=sortRows(rows.filter(r=>r.n.toLowerCase().includes(q))).slice(0,120);
    lw.innerHTML=miniTable(['Company','Market Cap','#P/E','#vs Peers','#Screens'],
      list.map(r=>[`<span class="clickable" data-co="${esc(r.n)}">${esc(r.n)}</span>
        <span class="sublabel">${esc(r.i)}</span>`,cr(r.mc),n(r.pe,1),pctHtml(r.disc),convPill(r.cv)]));
    wireLinks(lw);
  }
  inp.addEventListener('input',drawList);drawList();

  if(!SELECTED||!byName[SELECTED]){
    detail.appendChild(el('div','panel','<div class="empty">Pick a company to see its profile.</div>'));
    return;
  }
  const r=byName[SELECTED];
  const h=H[SELECTED]||{dates:[],pe:[],mcap:[],price:[],cv:[],appearances:1};

  const g=el('div','grid');
  g.appendChild(kpi('Market cap',cr(r.mc),esc(r.band)+' cap','','◧'));
  g.appendChild(kpi('P/E',r.flag?n(r.pe,2)+' ⚠':n(r.pe,1),
    r.flag?esc(r.flag):`peer median ${n(r.bench,1)}`,r.flag?'amber':'amber','₹'));
  g.appendChild(kpi('vs Peer P/E',r.disc==null?'—':(r.disc>=0?'+':'')+r.disc.toFixed(1)+'%',
    r.disc==null?'not comparable':(r.disc>=0?'below industry median':'above industry median'),
    r.disc==null?'':(r.disc>=0?'green':'red'),'▼'));
  g.appendChild(kpi('Screens matched',r.cv,`of ${Object.keys(RUN.groups).length} screen groups`,'indigo','★'));
  root.insertBefore(g,detail);

  const two=el('div','two');
  const pi=el('div','panel');
  pi.appendChild(el('h3','',esc(r.n)));
  pi.insertAdjacentHTML('beforeend',`<div class="kv">
    <div class="k">Industry</div><div>${esc(r.i)}</div>
    <div class="k">Broad industry</div><div>${esc(r.b)}</div>
    <div class="k">Current price</div><div>₹ ${n(r.pr,2)}</div>
    <div class="k">Earnings yield</div><div>${n(r.ey,2)}%</div>
    <div class="k">Peer benchmark</div><div>${n(r.bench,1)} <span class="sublabel">${esc(r.bsrc)} median</span></div>
    <div class="k">Conviction</div><div>${convPill(r.cv)} of ${Object.keys(RUN.groups).length} screen groups</div>
    <div class="k">Screens</div><div>${r.g.map(x=>`<span class="tag g">${esc(x)}</span>`).join('')||'—'}</div>
    <div class="k">Runs present</div><div>${h.appearances} / ${RUNS.length}</div>
  </div>`);
  two.appendChild(pi);

  const pp=el('div','panel');
  pp.appendChild(el('h3','','Peer comparison <span class="hint">same industry, by market cap</span>'));
  const peers=RUN.rows.filter(x=>x.i===r.i).sort((a,b)=>(b.mc||0)-(a.mc||0)).slice(0,15);
  pp.insertAdjacentHTML('beforeend',miniTable(['Company','Market Cap','#P/E','#vs Peers','#Screens'],
    peers.map(x=>[(x.n===r.n?'▸ ':'')+`<span class="clickable" data-co="${esc(x.n)}">${esc(x.n)}</span>`,
      cr(x.mc),n(x.pe,1),pctHtml(x.disc),convPill(x.cv)]),'scroll tall'));
  two.appendChild(pp);
  detail.appendChild(two);

  if(RUNS.length>1){
    const ph=el('div','panel');
    ph.appendChild(el('h3','','History'));
    const cw=el('div','chart-wrap');const cv=el('canvas');cw.appendChild(cv);ph.appendChild(cw);
    detail.appendChild(ph);
    chart('co',cv,{data:{labels:h.dates,datasets:[
      {type:'line',label:'Price ₹',data:h.price,borderColor:C.blue,
        backgroundColor:'rgba(59,110,245,.10)',fill:true,tension:.35,yAxisID:'y',pointRadius:4},
      {type:'line',label:'P/E',data:h.pe,borderColor:C.amber,tension:.35,yAxisID:'y1',pointRadius:4}]},
      options:{plugins:{legend:{display:true,labels:{usePointStyle:true,pointStyle:'circle',boxWidth:8}},
        tooltip:{enabled:true}},
        scales:{x:{grid:{display:false}},y:{position:'left'},
                y1:{position:'right',grid:{drawOnChartArea:false}}}}});
  }
  wireLinks(detail);
}

/* ---------- boot ---------- */
document.addEventListener('keydown',e=>{
  if(e.key==='/'&&document.activeElement.tagName!=='INPUT'){e.preventDefault();$('#f-q').focus()}
  if(e.key==='Escape')$$('.multi').forEach(m=>m.classList.remove('open'))});
buildFilters();
$('#view-sub').textContent=TITLES.overview[1];
render();
"""


# ==========================
# MAIN
# ==========================
def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    folder = os.path.abspath(folder)

    print(f"Scanning for {FILE_PATTERN} in: {folder}")
    dataset = build_dataset(folder)
    if dataset is None:
        print("No CONSOLIDATED_ALL_SCREENS_*.xlsx files found.")
        print("    python dashboard.py /path/to/folder")
        sys.exit(1)

    latest = dataset["runs"][-1]
    flagged = [r for r in latest["rows"] if r["flag"]]
    print(f"\nLoaded {len(dataset['runs'])} run(s): {', '.join(dataset['run_dates'])}")
    print(f"  companies   : {latest['total']}")
    print(f"  groups      : {len(latest['groups'])}")
    print(f"  suspect P/E : {len(flagged)}  (shown, excluded from peer medians)")

    for r in dataset["runs"]:
        f = r.get("freshness")
        if f and f["stale"]:
            print(f"\n  ⚠ STALE DATA in run {r['date']}: {f['unchanged']}/{f['compared']} "
                  f"({f['unchanged_pct']}%) identical to {f['prev_date']}.")

    html = render_html(dataset)
    out_path = os.path.join(folder, "dashboard.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\nDashboard written to: {out_path}  ({os.path.getsize(out_path)/1024/1024:.2f} MB)")


if __name__ == "__main__":
    main()
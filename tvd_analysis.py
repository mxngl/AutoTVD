import argparse
import csv
import io
import re
import json
import sys
import os
import webbrowser
import urllib.request
import urllib.error
from collections import defaultdict
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

GITHUB_REPO_RAW = (
    "https://https://github.com/mxngl/AutoTVD"
)

# Paths inside the repo (relative to GITHUB_REPO_RAW)
QTO_FILES = {
    "arch":   "qto/Architecture_TakeOff.csv",
    "struct": "qto/Structural_Schedule.csv",
}
COST_DATA_FILE = "cost_data.csv"

# Cluster names (from cost_data.csv) that use TAKEOFF quantities.
# All other clusters fall back to Fixed Quantity only.
TAKEOFF_CLUSTERS = {"Substructure", "Shell", "Interiors"}

# Revit categories to exclude from area/length/volume aggregation
EXCLUDE_CATEGORIES = {"Furniture"}

# Finish quantity mirrors: cost AC → (source takeoff AC, quantity field)
# These finishes automatically track the element they're applied to.
#   C3010 Wall Paint        = same SF as interior walls  (C1010)
#   C3020 Floor Finishes    = same SF as interior floors (B1010)
QUANTITY_MIRRORS: dict[str, tuple[str, str]] = {
    "B3010": ("B1020", "area_sf"),   # Roof Coverings tracks Roof Construction area
    "C3010": ("C1010", "area_sf"),
    "C3020": ("B1010", "area_sf"),
}

# Keyword-based AC splitting: when one real AC covers multiple cost line items,
# use keywords (matched case-insensitively against Category + Family + Type) to
# route each takeoff element to a synthetic sub-code.
# Format: { "real_ac": [(keywords, "sub_ac"), ..., ([], "fallback_sub_ac")] }
# First match wins; an entry with an empty keyword list is the catch-all fallback.
AC_KEYWORD_SPLIT: dict[str, list[tuple[list[str], str]]] = {
    "B2010": [
        (["storefront", "curtain wall", "curtain", "glazing"], "B2010.CW"),
        ([], "B2010.PW"),   # everything else → plaster wall with framing
    ],
}

# Assembly codes that represent toilet/bathroom stall elements.
# One C1030 stall is counted per element with any of these ACs
# (counted across ALL categories, including those in EXCLUDE_CATEGORIES).
TOILET_ACS: set[str] = {"D2010"}

# Cluster target values (from MARQUESINA TVD worksheet — Island Team 2026).
# Keys must exactly match the "Cluster Name" column in cost_data.csv.
# Note: "Special Contruction" preserves the typo that appears in cost_data.csv.
CLUSTER_TARGETS: dict[str, float] = {
    "Substructure":              1_781_276,
    "Shell":                     3_826_446,
    "Interiors":                 2_005_842,
    "Services":                  4_041_448,
    "Equipment and Furnishings": 1_319_286,
    "Special Contruction":       1_001_839,   # typo matches cost_data.csv
    "Building Sitework":         1_435_258,
    "General Conditions":        1_294_457,
}
TOTAL_TARGET: float = 16_700_000
GROSS_SF: int = 30_000          # gross square footage for $/SF index

OUTPUT_HTML = os.path.join(os.path.dirname(__file__), "TVD_Dashboard.html")

# Local fallback paths (used when GitHub is unreachable or not yet configured)
LOCAL_BASE = os.path.dirname(__file__)
LOCAL_QTO = {
    "arch":   os.path.join(LOCAL_BASE, "qto", "Architecture_TakeOff.csv"),
    "struct": os.path.join(LOCAL_BASE, "qto", "Structural_Schedule.csv"),
}
LOCAL_COST = os.path.join(LOCAL_BASE, "cost_data.csv")


# ─────────────────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────────────────

def _fetch_text(url: str) -> str:
    """Download a URL and return its text content."""
    with urllib.request.urlopen(url, timeout=15) as r:
        return r.read().decode("utf-8-sig")


def _read_local(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def _fix_bom(fieldnames):
    """Strip UTF-8 BOM artefacts from column names."""
    return [f.replace("ï»¿", "").replace("\ufeff", "") for f in fieldnames]


def parse_qty_str(val: str) -> float:
    """Parse quantity strings like '6590 SF', '136\\' - 0\\"', '42.75 CF'."""
    if not val or not val.strip():
        return 0.0
    ft_in = re.match(r"(-?\d+)'\s*-\s*(\d+(?:\.\d+)?)\s*\"", val.strip())
    if ft_in:
        return float(ft_in.group(1)) + float(ft_in.group(2)) / 12
    m = re.search(r"(-?\d[\d,]*\.?\d*)", val.replace(",", ""))
    return float(m.group(1)) if m else 0.0


def parse_cost(val: str):
    """
    Parse cost strings, handling both EU and US number formats.
      EU: period = thousands separator, comma = decimal  → "6.251,07" → 6251.07
      US: comma  = thousands separator, period = decimal → "1,000.00" → 1000.00
    The format is detected by which separator appears last in the string.
    """
    if not val or not val.strip():
        return None
    v = val.strip().replace("$", "").replace(" ", "")
    if "," in v and "." in v:
        if v.rfind(",") > v.rfind("."):   # EU: comma is the decimal separator
            v = v.replace(".", "").replace(",", ".")
        else:                              # US: period is the decimal separator
            v = v.replace(",", "")
    elif "," in v:                         # EU with no thousands sep: "25,00"
        v = v.replace(",", ".")
    try:
        return float(v)
    except ValueError:
        return None


def fmt_usd(n) -> str:
    if n is None or n == 0:
        return "—"
    return f"${n:,.0f}"


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────

def load_csv_text(text: str) -> list[dict]:
    """Parse CSV text into a list of dicts with BOM-cleaned column names."""
    reader = csv.DictReader(io.StringIO(text))
    fields = reader.fieldnames or []
    fixed = _fix_bom(fields)
    rows = []
    for row in reader:
        rows.append({fixed[i]: row.get(fields[i], "") for i in range(len(fields))})
    return rows


def fetch_qto_and_cost(ci_mode: bool = False):
    """
    In CI mode (--ci flag or running in GitHub Actions): always use local files.
    Otherwise: try GitHub first, fall back to local.
    Returns (arch_rows, struct_rows, cost_rows, source_label).
    """
    use_github = not ci_mode and "YOUR_USERNAME" not in GITHUB_REPO_RAW

    if use_github:
        try:
            arch_text   = _fetch_text(f"{GITHUB_REPO_RAW}/{QTO_FILES['arch']}")
            struct_text = _fetch_text(f"{GITHUB_REPO_RAW}/{QTO_FILES['struct']}")
            cost_text   = _fetch_text(f"{GITHUB_REPO_RAW}/{COST_DATA_FILE}")
            source = f"GitHub - {GITHUB_REPO_RAW}"
            print("Loaded data from GitHub")
        except urllib.error.URLError as e:
            print(f"GitHub fetch failed ({e}); falling back to local files.")
            use_github = False

    if not use_github:
        arch_text   = _read_local(LOCAL_QTO["arch"])
        struct_text = _read_local(LOCAL_QTO["struct"])
        cost_text   = _read_local(LOCAL_COST)
        source = f"Local files - {LOCAL_BASE}"
        print("Loaded data from local files")

    return (
        load_csv_text(arch_text),
        load_csv_text(struct_text),
        load_csv_text(cost_text),
        source,
    )


# ─────────────────────────────────────────────────────────────────────────────
# TAKEOFF PROCESSING
# ─────────────────────────────────────────────────────────────────────────────

def merge_takeoffs(arch: list[dict], struct: list[dict]) -> list[dict]:
    """
    Merge architectural and structural takeoffs by ElementId (no duplicates).
    Structural rows override architectural where ElementId matches.
    """
    combined: dict[str, dict] = {}
    for row in arch:
        combined[row["ElementId"]] = row
    for row in struct:
        combined[row["ElementId"]] = row  # struct wins on overlap
    return list(combined.values())


_UNMAPPED_EXPORT_COLS = [
    "ElementId", "Category", "Family", "Type", "Level", "Mark",
    "Area", "Length", "Volume", "Material", "Comments",
]


def aggregate_quantities(rows: list[dict], exclude_categories: set) -> tuple[dict, int, dict, list[dict]]:
    """
    Aggregate per Assembly Code for non-excluded categories:
      area_sf, length_lf, volume_cf, count

    Also builds all_ac_counts: element counts across ALL categories (including
    excluded ones like Furniture) — used for toilet stall mapping.

    Returns (code_qtys, unmapped_count, all_ac_counts, unmapped_rows).
    unmapped_rows = non-excluded rows with no Assembly Code, trimmed to export columns.
    """
    code_qtys: dict[str, dict] = defaultdict(
        lambda: {"area_sf": 0.0, "length_lf": 0.0, "volume_cf": 0.0, "count": 0}
    )
    all_ac_counts: dict[str, int] = defaultdict(int)
    unmapped = 0
    unmapped_rows: list[dict] = []

    for row in rows:
        ac = row.get("Assembly Code", "").strip()
        cat = row.get("Category", "").strip()

        # Resolve keyword-based sub-codes before any counting
        if ac in AC_KEYWORD_SPLIT:
            name = " ".join([
                row.get("Category", ""),
                row.get("Family",   ""),
                row.get("Type",     ""),
            ]).lower()
            for keywords, sub_ac in AC_KEYWORD_SPLIT[ac]:
                if not keywords or any(kw in name for kw in keywords):
                    ac = sub_ac
                    break

        # Count every element by AC regardless of category (for toilet mapping)
        if ac:
            all_ac_counts[ac] += 1

        # Skip excluded categories from quantity aggregation
        if cat in exclude_categories:
            continue

        if not ac:
            unmapped += 1
            unmapped_rows.append({col: row.get(col, "") for col in _UNMAPPED_EXPORT_COLS})
            continue

        q = code_qtys[ac]
        q["area_sf"]   += parse_qty_str(row.get("Area", ""))
        q["length_lf"] += parse_qty_str(row.get("Length", ""))
        q["volume_cf"] += parse_qty_str(row.get("Volume", ""))
        q["count"]     += 1

    return dict(code_qtys), unmapped, dict(all_ac_counts), unmapped_rows


# ─────────────────────────────────────────────────────────────────────────────
# COST DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────

def load_cost_data(rows: list[dict]) -> list[dict]:
    """Parse cost data rows, skipping blank lines."""
    result = []
    for row in rows:
        ac = row.get("Assembly Code", "").strip()
        if not ac:
            continue
        fq_raw = row.get("Fixed Quantity", "").strip()
        result.append({
            "cluster": row.get("Cluster Name", "").strip(),
            "ac":      ac,
            "group":   row.get("Assembly Group Name", "").strip(),
            "desc":    row.get("Description             ", "").strip(),
            "unit":    row.get("Unit             ", "").strip(),
            "cost":    parse_cost(row.get("Total O&P", "")),
            "fixed_qty": float(fq_raw) if fq_raw else None,
        })
    return result


# ─────────────────────────────────────────────────────────────────────────────
# COST CALCULATION
# ─────────────────────────────────────────────────────────────────────────────

def pick_quantity(
    cr: dict,
    code_qtys: dict,
    all_ac_counts: dict,
) -> tuple[float, str]:
    """
    Determine the quantity to use for a cost line item.

    Rules (in priority order):
      1. Fixed Quantity in cost_data → always wins.
      2. C1030 Toilet Partitions → count elements with TOILET_ACS (all categories).
      3. QUANTITY_MIRRORS entry → use a different AC's takeoff quantity.
      4. Normal takeoff lookup by unit type (clusters A–C only).
      5. Non-A/B/C cluster with no Fixed Quantity → 0.
    """
    fq = cr["fixed_qty"]
    ac = cr["ac"]

    # Rule 1 — fixed quantity overrides everything
    if fq is not None:
        return fq, "Fixed"

    # Rule 2 — toilet stall count (C1030 only, uses TOILET_ACS across all categories)
    if ac == "C1030":
        count = sum(all_ac_counts.get(t, 0) for t in TOILET_ACS)
        label = f"Toilet elements ({', '.join(sorted(TOILET_ACS))})"
        return float(count), label

    # Only clusters A–C use takeoff data from here on
    if cr["cluster"] not in TAKEOFF_CLUSTERS:
        return 0.0, "Fixed only (none set)"

    # Rule 3 — finish mirrors: C3010/C3020 track another AC's area
    if ac in QUANTITY_MIRRORS:
        src_ac, field = QUANTITY_MIRRORS[ac]
        q = code_qtys.get(src_ac)
        if q:
            return q[field], f"Mirror: {src_ac} area"
        return 0.0, f"Mirror source {src_ac} not in takeoff"

    # Rule 4 — normal unit-based lookup
    q = code_qtys.get(ac)
    if q is None:
        return 0.0, "No takeoff match"

    u = cr["unit"].upper()
    if u in ("SF", "GSF"):
        return q["area_sf"],            "Area (SF)"
    if u == "MSF":
        return q["area_sf"] / 1000,     "Area/1000 (MSF)"
    if u == "LF":
        return q["length_lf"],          "Length (LF)"
    if u in ("EA", "FLIGHT"):
        return float(q["count"]),       "Count (EA)"
    if u == "CY":
        return q["volume_cf"] / 27,     "Volume (CY)"
    if u == "CF":
        return q["volume_cf"],          "Volume (CF)"
    return 0.0, f"Unknown unit: {cr['unit']}"


def calculate_costs(
    cost_data: list[dict],
    code_qtys: dict,
    all_ac_counts: dict,
) -> list[dict]:
    """Apply unit costs to quantities and return enriched line items."""
    # ACs that have special quantity rules (no "AC not in takeoff" warning)
    special_acs = set(QUANTITY_MIRRORS.keys()) | {"C1030"}

    results = []
    for cr in cost_data:
        qty, qty_src = pick_quantity(cr, code_qtys, all_ac_counts)
        cost = cr["cost"]
        line_total = qty * cost if (cost is not None and qty) else 0.0

        notes = []
        if cost is None:
            notes.append("No unit cost")
        if (qty == 0
                and cr["fixed_qty"] is None
                and cr["cluster"] in TAKEOFF_CLUSTERS
                and cr["ac"] not in special_acs):
            notes.append("Zero qty from takeoff")
        if (cr["ac"] not in code_qtys
                and cr["fixed_qty"] is None
                and cr["cluster"] in TAKEOFF_CLUSTERS
                and cr["ac"] not in special_acs):
            notes.append("AC not in takeoff")

        results.append({
            "cluster":    cr["cluster"],
            "ac":         cr["ac"],
            "group":      cr["group"],
            "desc":       cr["desc"],
            "unit":       cr["unit"],
            "unit_cost":  cost,
            "qty":        round(qty, 2),
            "qty_src":    qty_src,
            "total":      round(line_total, 2),
            "notes":      " | ".join(notes),
        })
    return results


def build_cluster_summary(results: list[dict]) -> list[dict]:
    totals: dict[str, float] = defaultdict(float)
    for r in results:
        totals[r["cluster"]] += r["total"]

    # Preserve order from results
    seen = []
    for r in results:
        if r["cluster"] not in seen:
            seen.append(r["cluster"])

    rows = [{"cluster": c, "total": totals[c]} for c in seen]
    rows.append({"cluster": "GRAND TOTAL", "total": sum(totals.values())})
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# HTML DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────

CLUSTER_COLORS = {
    "Substructure":             "#A85520",   # dark terracotta — earth/foundation
    "Shell":                    "#C46626",   # primary terracotta — structure
    "Interiors":                "#7A9B76",   # sage green — interior spaces
    "Services":                 "#5A7A56",   # dark sage — mechanical/utility
    "Equipment and Furnishings":"#9BB08A",   # light sage — furnishings
    "Special Contruction":      "#D4834A",   # light terracotta — specialty
    "Building Sitework":        "#6B6B6B",   # medium charcoal — site/ground
    "General Conditions":       "#424242",   # dark charcoal — administration
}

def _cluster_color(name: str) -> str:
    return CLUSTER_COLORS.get(name, "#94a3b8")


def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    """Convert '#rrggbb' to 'rgba(r,g,b,alpha)' for chart backgrounds."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def generate_html(
    results: list[dict],
    summary: list[dict],
    unmapped_count: int,
    data_source: str,
    targets: dict[str, float],
    total_target: float,
    gross_sf: int = 30_000,
    unmapped_rows: list[dict] | None = None,
) -> str:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")

    chart_rows  = [r for r in summary if r["cluster"] != "GRAND TOTAL"]
    grand_total = next(r["total"] for r in summary if r["cluster"] == "GRAND TOTAL")

    # ── delta helpers ─────────────────────────────────────────────────────────
    grand_delta = grand_total - total_target

    def _delta_cls(d: float) -> str:
        return "delta-over" if d > 0 else ("delta-under" if d < 0 else "delta-neutral")

    def _fmt_delta(d: float) -> str:
        if d == 0:
            return "$0"
        sign = "+" if d > 0 else "-"
        return f"{sign}${abs(d):,.0f}"

    def _fmt_pct(d: float, base: float) -> str:
        if not base:
            return ""
        p = d / base * 100
        return f" ({'+'if p>0 else ''}{p:.1f}%)"

    # ── pie chart data (all clusters, high-contrast palette) ─────────────────
    PIE_PALETTE = [
        "#F94144", "#F3722C", "#F8961E", "#F9C74F",
        "#90BE6D", "#43AA8B", "#577590", "#415262",
    ]
    pie_labels    = [r["cluster"] or "Other" for r in chart_rows]
    pie_estimates = [r["total"] for r in chart_rows]
    pie_colors    = PIE_PALETTE[: len(pie_labels)]

    pie_labels_js    = json.dumps(pie_labels)
    pie_estimates_js = json.dumps(pie_estimates)
    pie_colors_js    = json.dumps(pie_colors)
    total_target_js  = json.dumps(total_target)

    # ── chart data — only clusters that have a defined target ─────────────────
    cmp = [(r, targets[r["cluster"]]) for r in chart_rows if r["cluster"] in targets]
    ct_labels    = json.dumps([r["cluster"] for r, _ in cmp])
    ct_estimates = json.dumps([r["total"]   for r, _ in cmp])
    ct_targets   = json.dumps([t            for _, t in cmp])
    ct_deltas    = json.dumps([r["total"] - t for r, t in cmp])
    ct_colors    = json.dumps([_cluster_color(r["cluster"]) for r, _ in cmp])
    ct_tgt_bg    = json.dumps([_hex_to_rgba(_cluster_color(r["cluster"]), 0.25) for r, _ in cmp])

    def _fmt_psf(n: float) -> str:
        return f"${n / gross_sf:,.0f}/SF"

    # ── cluster letter + scroll-target id ────────────────────────────────────
    cluster_letter: dict[str, str] = {}
    for _r in results:
        _c = _r["cluster"]
        if _c not in cluster_letter and _r["ac"]:
            cluster_letter[_c] = _r["ac"][0].upper()

    def _cluster_id(name: str) -> str:
        return "cl-" + re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")

    # ── summary cards ─────────────────────────────────────────────────────────
    # Grand total card first
    gdc = _delta_cls(grand_delta)
    cards_html = f"""
        <div class="card grand-total">
          <div class="card-label">Grand Total</div>
          <div class="card-est">{fmt_usd(grand_total)}</div>
          <div class="card-tvd-row">
            <span class="tvd-label">Target</span>
            <span class="tvd-val">{fmt_usd(total_target)}</span>
          </div>
          <div class="card-delta {gdc}">{_fmt_delta(grand_delta)}{_fmt_pct(grand_delta, total_target)}</div>
          <div class="card-tvd-row" style="margin-top:6px">
            <span class="tvd-label">$/SF ({gross_sf:,} GSF)</span>
            <span class="tvd-val">{_fmt_psf(grand_total)}</span>
          </div>
        </div>"""

    for r in chart_rows:
        clr   = r["cluster"] or "Other"
        est   = r["total"]
        color = _cluster_color(r["cluster"])
        tgt   = targets.get(r["cluster"])

        if tgt:
            delta  = est - tgt
            dc     = _delta_cls(delta)
            t_html = (
                f'<div class="card-tvd-row">'
                f'<span class="tvd-label">Target</span>'
                f'<span class="tvd-val">{fmt_usd(tgt)}</span>'
                f'</div>'
                f'<div class="card-delta {dc}">{_fmt_delta(delta)}{_fmt_pct(delta, tgt)}</div>'
            )
        else:
            t_html = '<div class="card-tvd-row" style="margin-top:6px"><span class="tvd-label" style="font-style:italic">No target set</span></div>'

        cards_html += f"""
        <div class="card card-link" style="border-top:4px solid {color}" onclick="document.getElementById('{_cluster_id(clr)}').scrollIntoView({{behavior:'smooth'}})">
          <div class="card-label">{clr}</div>
          <div class="card-est">{fmt_usd(est)}</div>
          {t_html}
          <div class="card-tvd-row" style="margin-top:6px">
            <span class="tvd-label">$/SF</span>
            <span class="tvd-val">{_fmt_psf(est)}</span>
          </div>
        </div>"""

    # ── detail table rows ─────────────────────────────────────────────────────
    # Pre-compute per-cluster totals for header rows
    cluster_totals: dict[str, float] = {}
    for r in results:
        cluster_totals[r["cluster"]] = cluster_totals.get(r["cluster"], 0) + r["total"]

    detail_rows_html = ""
    prev_cluster = None
    for r in results:
        if r["cluster"] != prev_cluster:
            bg = _cluster_color(r["cluster"])
            cl_total = cluster_totals.get(r["cluster"], 0)
            cl_psf   = f"${cl_total / gross_sf:,.0f}/SF" if cl_total else "—"
            letter = cluster_letter.get(r["cluster"], "")
            detail_rows_html += f"""
            <tr id="{_cluster_id(r['cluster'])}" class="cluster-header" style="background:{bg}20;border-left:4px solid {bg}">
              <td colspan="7" style="font-weight:600;padding:6px 12px;color:{bg}">
                <span style="font-family:monospace;font-weight:700;margin-right:10px;opacity:.55">{letter}</span>{r['cluster']}
              </td>
              <td class="num" style="padding:6px 12px;font-weight:700;color:{bg}">{fmt_usd(cl_total)}</td>
              <td class="num" style="padding:6px 12px;font-size:.72rem;color:{bg};opacity:.75">{cl_psf}</td>
            </tr>"""
            prev_cluster = r["cluster"]

        zero_class = ' class="zero-row"' if r["total"] == 0 else ""
        fixed_badge = ' <span class="badge-fixed">fixed</span>' if r["qty_src"] == "Fixed" else ""
        note_cell = f'<span class="note">{r["notes"]}</span>' if r["notes"] else ""
        unit_cost_str = fmt_usd(r["unit_cost"]) if r["unit_cost"] is not None else "—"

        detail_rows_html += f"""
        <tr{zero_class}>
          <td class="mono">{r['ac']}</td>
          <td>{r['group']}</td>
          <td>{r['desc']}</td>
          <td class="mono">{r['unit']}</td>
          <td class="num">{unit_cost_str}</td>
          <td class="num">{r['qty']:,.1f}{fixed_badge}</td>
          <td class="qty-src">{r['qty_src']}</td>
          <td class="num total-cell">{fmt_usd(r['total'])}</td>
          <td>{note_cell}</td>
        </tr>"""

    unmapped_json = json.dumps(unmapped_rows or [])
    unmapped_cols_json = json.dumps(_UNMAPPED_EXPORT_COLS)

    detail_rows_html += f"""
    <tr class="unmapped-row unmapped-link" onclick="downloadUnmapped()" title="Click to download unmapped elements as CSV for Revit review">
      <td class="mono">—</td>
      <td colspan="5">Unmapped elements (no Assembly Code in takeoff)</td>
      <td class="qty-src">Takeoff</td>
      <td class="num total-cell">—</td>
      <td><span class="note">{unmapped_count} elements &mdash; <u>download for review</u> &#8595;</span></td>
    </tr>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>TVD Cost Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: system-ui, -apple-system, sans-serif; background: #EDE8E3; color: #424242; }}
  header {{ background: #424242; color: #FEFFFE; padding: 24px 32px; display: flex; justify-content: space-between; align-items: center; }}
  header h1 {{ font-size: 1.4rem; font-weight: 700; letter-spacing: -.01em; }}
  header .meta {{ font-size: .8rem; color: #A8A8A8; text-align: right; }}
  .container {{ max-width: 1400px; margin: 0 auto; padding: 28px 32px; }}
  /* cards */
  .cards {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 16px; margin-bottom: 32px; }}
  .card {{ background: #FEFFFE; border-radius: 10px; padding: 18px 20px; box-shadow: 0 1px 4px rgba(0,0,0,.1); }}
  .card.grand-total {{ background: #424242; color: #FEFFFE; border-top: 4px solid #C46626; }}
  .card-link {{ cursor: pointer; transition: transform .12s, box-shadow .12s; }}
  .card-link:hover {{ transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0,0,0,.14); }}
  .card-label {{ font-size: .72rem; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; color: #6B6B6B; margin-bottom: 4px; }}
  .grand-total .card-label {{ color: #A8A8A8; }}
  .card-est {{ font-size: 1.35rem; font-weight: 700; margin: 2px 0 8px; }}
  .card-tvd-row {{ display: flex; justify-content: space-between; align-items: baseline; margin-top: 4px; }}
  .tvd-label {{ font-size: .72rem; color: #6B6B6B; }}
  .tvd-val {{ font-size: .82rem; font-weight: 600; color: #6B6B6B; }}
  .grand-total .tvd-label, .grand-total .tvd-val {{ color: #A8A8A8; }}
  .card-delta {{ font-size: .8rem; font-weight: 600; margin-top: 6px; }}
  .delta-over {{ color: #C44040; }}
  .delta-under {{ color: #7A9B76; }}
  .delta-neutral {{ color: #A8A8A8; }}
  .grand-total .delta-over {{ color: #E8A8A8; }}
  .grand-total .delta-under {{ color: #B5CEAD; }}
  /* charts */
  .chart-section {{ background: #FEFFFE; border-radius: 10px; padding: 24px; box-shadow: 0 1px 4px rgba(0,0,0,.1); margin-bottom: 16px; }}
  .chart-wrap {{ position: relative; }}
  .section-title {{ font-size: .85rem; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: #6B6B6B; margin-bottom: 18px; }}
  /* table */
  .table-section {{ background: #FEFFFE; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,.1); overflow: hidden; margin-bottom: 32px; margin-top: 16px; }}
  .table-header {{ padding: 18px 24px; border-bottom: 1px solid #E0DBD5; }}
  table {{ width: 100%; border-collapse: collapse; font-size: .82rem; }}
  thead th {{ background: #EAE6E0; padding: 10px 12px; text-align: left; font-size: .72rem; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; color: #6B6B6B; white-space: nowrap; }}
  tbody td {{ padding: 8px 12px; border-bottom: 1px solid #EAE6E0; vertical-align: top; }}
  tbody tr:last-child td {{ border-bottom: none; }}
  tbody tr:hover {{ background: #F5F1EC; }}
  .cluster-header td {{ border-bottom: none !important; }}
  .zero-row {{ opacity: .5; }}
  .unmapped-row {{ background: #FFF5EA; }}
  .unmapped-row td {{ font-style: italic; color: #7A4020; }}
  .unmapped-link {{ cursor: pointer; transition: background .12s; }}
  .unmapped-link:hover {{ background: #FFE8D0 !important; }}
  .total-cell {{ font-weight: 600; color: #424242; }}
  .zero-row .total-cell {{ font-weight: normal; }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  .mono {{ font-family: monospace; font-size: .78rem; }}
  .qty-src {{ font-size: .72rem; color: #A8A8A8; }}
  .note {{ font-size: .72rem; color: #C44040; }}
  .badge-fixed {{ display: inline-block; background: rgba(196,102,38,0.15); color: #C46626; font-size: .65rem; font-weight: 600; border-radius: 4px; padding: 1px 5px; margin-left: 4px; text-transform: uppercase; }}
  footer {{ text-align: center; font-size: .75rem; color: #A8A8A8; padding: 24px; }}
  /* toggle buttons */
  .toggle-group {{ display: flex; border-radius: 6px; overflow: hidden; border: 1px solid #E0DBD5; }}
  .toggle-btn {{ padding: 6px 18px; font-size: .78rem; font-weight: 600; background: #FEFFFE; color: #6B6B6B; border: none; cursor: pointer; transition: background .15s, color .15s; letter-spacing: .02em; }}
  .toggle-btn.active {{ background: #424242; color: #FEFFFE; }}
  .toggle-btn:hover:not(.active) {{ background: #F5F1EC; color: #424242; }}
  .chart-section-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; }}
</style>
</head>
<body>
<header>
  <div>
    <h1>TVD Cost Dashboard</h1>
    <div style="font-size:.8rem;color:#A8A8A8;margin-top:4px">Island Team 2026</div>
  </div>
  <div class="meta">
    Last updated: {ts}<br>
    Source: {data_source}
  </div>
</header>

<div class="container">

  <!-- Cost Summary Cards -->
  <div class="section-title" style="margin-bottom:14px">Cost Summary</div>
  <div class="cards">
    {cards_html}
  </div>

  <!-- Pie Chart: Cost by Cluster (toggleable) -->
  <div class="chart-section">
    <div class="chart-section-header">
      <div class="section-title" style="margin-bottom:0">Cost by Cluster</div>
      <div class="toggle-group">
        <button class="toggle-btn active" id="btnDollar" onclick="switchPieMode('dollar')">$ Value</button>
        <button class="toggle-btn" id="btnPct" onclick="switchPieMode('pct')">% of Target</button>
      </div>
    </div>
    <div class="chart-wrap" style="height:320px">
      <canvas id="pieChart"></canvas>
    </div>
  </div>

  <!-- Chart 1: Estimate vs Target by Cluster -->
  <div class="chart-section">
    <div class="section-title">Estimate vs. Target &mdash; by Cluster</div>
    <div class="chart-wrap" style="height:340px">
      <canvas id="comparisonChart"></canvas>
    </div>
  </div>

  <!-- Chart 2: Variance per Cluster -->
  <div class="chart-section">
    <div class="section-title">Variance per Cluster (Estimate &minus; Target)</div>
    <div class="chart-wrap" style="height:220px">
      <canvas id="deltaChart"></canvas>
    </div>
  </div>

  <!-- Line Item Detail Table -->
  <div class="table-section">
    <div class="table-header">
      <div class="section-title" style="margin-bottom:4px">Line Item Detail</div>
      <div style="font-size:.78rem;color:#6B6B6B">
        Clusters A&ndash;C use takeoff quantities &nbsp;&middot;&nbsp;
        Other clusters use fixed quantities only &nbsp;&middot;&nbsp;
        <span style="color:#C44040">Red text</span> = review flag &nbsp;&middot;&nbsp;
        Dimmed rows = $0 line items &nbsp;&middot;&nbsp;
        <span class="badge-fixed">fixed</span> = fixed quantity from cost data
      </div>
    </div>
    <table>
      <thead>
        <tr>
          <th>Assy Code</th>
          <th>Group</th>
          <th>Description</th>
          <th>Unit</th>
          <th class="num">Unit Cost</th>
          <th class="num">Quantity</th>
          <th>Qty Source</th>
          <th class="num">Line Total</th>
          <th>Notes</th>
        </tr>
      </thead>
      <tbody>
        {detail_rows_html}
      </tbody>
    </table>
  </div>

</div>

<footer>
  Island Team 2026 &middot; Last updated: {ts} &middot; Data: {data_source}
</footer>

<script>
// ── Unmapped elements download ────────────────────────────────────────────────
const unmappedData = {unmapped_json};
const unmappedCols = {unmapped_cols_json};

function downloadUnmapped() {{
  const escape = v => {{
    const s = (v == null ? '' : String(v));
    return (s.includes(',') || s.includes('"') || s.includes('\\n'))
      ? '"' + s.replace(/"/g, '""') + '"'
      : s;
  }};
  let csv = unmappedCols.join(',') + '\\n';
  for (const row of unmappedData) {{
    csv += unmappedCols.map(c => escape(row[c])).join(',') + '\\n';
  }}
  const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'AutoTVD_Unmapped_Elements.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}}

// ── Pie Chart: Cost by Cluster ───────────────────────────────────────────────
const pieLabels    = {pie_labels_js};
const pieEstimates = {pie_estimates_js};
const pieColors    = {pie_colors_js};
const totalTarget  = {total_target_js};
const grandEst     = pieEstimates.reduce((a, b) => a + b, 0);
const UNALLOC_COLOR = '#D6D0CA';

let pieMode = 'dollar';

function buildPieDataset(mode) {{
  if (mode === 'dollar') {{
    return {{ labels: pieLabels, data: pieEstimates, colors: pieColors }};
  }}
  // % of total target
  const pcts     = pieEstimates.map(v => parseFloat((v / totalTarget * 100).toFixed(2)));
  const usedPct  = pcts.reduce((a, b) => a + b, 0);
  const remaining = parseFloat((100 - usedPct).toFixed(2));
  if (remaining > 0.01) {{
    return {{
      labels: [...pieLabels, 'Unallocated'],
      data:   [...pcts, remaining],
      colors: [...pieColors, UNALLOC_COLOR],
    }};
  }}
  return {{ labels: pieLabels, data: pcts, colors: pieColors }};
}}

let pieChart = null;

function renderPieChart(mode) {{
  pieMode = mode;
  document.getElementById('btnDollar').classList.toggle('active', mode === 'dollar');
  document.getElementById('btnPct').classList.toggle('active', mode === 'pct');
  if (pieChart) {{ pieChart.destroy(); pieChart = null; }}
  const d = buildPieDataset(mode);
  pieChart = new Chart(document.getElementById('pieChart'), {{
    type: 'doughnut',
    data: {{
      labels: d.labels,
      datasets: [{{
        data: d.data,
        backgroundColor: d.colors,
        borderColor: '#FEFFFE',
        borderWidth: 2,
        hoverOffset: 8,
      }}]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{
          display: true,
          position: 'right',
          labels: {{
            boxWidth: 14,
            padding: 14,
            font: {{ size: 12 }},
            color: '#424242',
          }}
        }},
        tooltip: {{
          callbacks: {{
            label: ctx => {{
              if (pieMode === 'dollar') {{
                const v = ctx.parsed;
                const pct = (v / grandEst * 100).toFixed(1);
                return ctx.label + ': $' + Math.round(v).toLocaleString('en-US') + ' (' + pct + '%)';
              }} else {{
                return ctx.label + ': ' + ctx.parsed.toFixed(1) + '% of target';
              }}
            }}
          }}
        }}
      }}
    }}
  }});
}}

function switchPieMode(mode) {{ renderPieChart(mode); }}

renderPieChart('dollar');

// ── Chart 1: Estimate vs. Target ──────────────────────────────────────────────
new Chart(document.getElementById('comparisonChart'), {{
  type: 'bar',
  data: {{
    labels: {ct_labels},
    datasets: [
      {{
        label: 'Estimate',
        data: {ct_estimates},
        backgroundColor: {ct_colors},
        borderRadius: 5,
        borderSkipped: false,
      }},
      {{
        label: 'Target',
        data: {ct_targets},
        backgroundColor: {ct_tgt_bg},
        borderColor: {ct_colors},
        borderWidth: 1.5,
        borderRadius: 5,
        borderSkipped: false,
      }}
    ]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    plugins: {{
      legend: {{ display: true, position: 'top' }},
      tooltip: {{
        callbacks: {{
          label: ctx => ctx.dataset.label + ': $' + ctx.parsed.x.toLocaleString('en-US', {{maximumFractionDigits: 0}})
        }}
      }}
    }},
    scales: {{
      x: {{
        ticks: {{ callback: v => '$' + (v / 1e6).toFixed(1) + 'M' }},
        grid: {{ color: '#EAE6E0' }}
      }},
      y: {{ grid: {{ display: false }} }}
    }}
  }}
}});

// ── Chart 2: Variance (Estimate - Target) ────────────────────────────────────
const rawDeltas = {ct_deltas};
const deltaBarColors = rawDeltas.map(d => d > 0 ? '#C44040' : '#7A9B76');
new Chart(document.getElementById('deltaChart'), {{
  type: 'bar',
  data: {{
    labels: {ct_labels},
    datasets: [{{
      label: 'Variance (Estimate - Target)',
      data: rawDeltas,
      backgroundColor: deltaBarColors,
      borderRadius: 5,
      borderSkipped: false,
    }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    plugins: {{
      legend: {{ display: false }},
      tooltip: {{
        callbacks: {{
          label: ctx => {{
            const v = ctx.parsed.x;
            const sign = v >= 0 ? '+' : '';
            return 'Variance: ' + sign + '$' + Math.round(Math.abs(v)).toLocaleString('en-US');
          }}
        }}
      }}
    }},
    scales: {{
      x: {{
        ticks: {{
          callback: v => {{
            if (v === 0) return '$0';
            const sign = v > 0 ? '+' : '-';
            return sign + '$' + (Math.abs(v) / 1e6).toFixed(1) + 'M';
          }}
        }},
        grid: {{ color: '#f1f5f9' }}
      }},
      y: {{ grid: {{ display: false }} }}
    }}
  }}
}});
</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="AutoTVD Cost Analysis")
    parser.add_argument(
        "--ci",
        action="store_true",
        help="CI mode: use local files, write to docs/index.html, skip browser",
    )
    args = parser.parse_args()
    ci_mode = args.ci

    print("AutoTVD Cost Analysis" + (" [CI mode]" if ci_mode else ""))

    # 1. Load data
    arch_rows, struct_rows, cost_csv_rows, source = fetch_qto_and_cost(ci_mode)

    # 2. Merge takeoffs (dedup by ElementId)
    all_elements = merge_takeoffs(arch_rows, struct_rows)
    dupes = len(arch_rows) + len(struct_rows) - len(all_elements)
    print(f"   Elements: arch={len(arch_rows)}, struct={len(struct_rows)}, "
          f"combined={len(all_elements)}, duplicates removed={dupes}")

    # 3. Aggregate takeoff quantities (excluding furnishings)
    code_qtys, unmapped_count, all_ac_counts, unmapped_rows = aggregate_quantities(
        all_elements, EXCLUDE_CATEGORIES
    )
    toilet_count = sum(all_ac_counts.get(t, 0) for t in TOILET_ACS)
    print(f"   Assembly codes in takeoff: {len(code_qtys)}")
    print(f"   Unmapped elements (no AC): {unmapped_count}")
    print(f"   Toilet elements found (C1030): {toilet_count}")

    # 4. Load cost data
    cost_data = load_cost_data(cost_csv_rows)
    print(f"   Cost line items: {len(cost_data)}")

    # 5. Calculate line item costs
    results = calculate_costs(cost_data, code_qtys, all_ac_counts)

    # 6. Build cluster summary
    summary = build_cluster_summary(results)

    # 7. Print summary to console
    print()
    print(f"  {'Cluster':<35} {'Total':>14}")
    print("  " + "─" * 51)
    for r in summary:
        marker = " ◄" if r["cluster"] == "GRAND TOTAL" else ""
        print(f"  {r['cluster']:<35} {fmt_usd(r['total']):>14}{marker}")
    print(f"\n  Unmapped elements: {unmapped_count} (no Assembly Code)")

    # 8. Generate HTML dashboard
    if ci_mode:
        out_dir = os.path.join(LOCAL_BASE, "docs")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, "index.html")
    else:
        out_path = OUTPUT_HTML

    html = generate_html(results, summary, unmapped_count, source, CLUSTER_TARGETS, TOTAL_TARGET, GROSS_SF, unmapped_rows)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\nDashboard saved: {out_path}")

    # 9. Open in browser (local mode only)
    if not ci_mode:
        webbrowser.open(f"file:///{out_path.replace(os.sep, '/')}")


if __name__ == "__main__":
    main()

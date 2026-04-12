"""
TVD Cost Analysis — AutoTVD
============================
Reads QTO schedule CSVs from a GitHub repository, applies cost mapping,
and generates an HTML dashboard for Target Value Design review.

Setup:
  1. Create a GitHub repo and push your QTO files + cost_data.csv to it
     following the folder layout described below.
  2. Set GITHUB_REPO_RAW to your repo's raw base URL.
  3. Run:  python tvd_analysis.py
  4. The dashboard opens automatically in your browser.

Expected repo layout:
  /
  ├── tvd_analysis.py          ← this file
  ├── cost_data.csv            ← cost mapping (never changes often)
  └── qto/
      ├── Architecture_TakeOff.csv
      └── Structural_Schedule.csv

Cluster logic:
  • Clusters A–C (Substructure / Shell / Interiors):
        quantities come from the QTO takeoff.
  • All other clusters (Services, Equipment, Sitework …):
        quantities come from the "Fixed Quantity" column in cost_data.csv only.
        If a row has no Fixed Quantity it contributes $0.
"""

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
# CONFIG  ← edit these to match your GitHub repo
# ─────────────────────────────────────────────────────────────────────────────

GITHUB_REPO_RAW = (
    "https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main"
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

# Revit categories to exclude from the takeoff entirely
EXCLUDE_CATEGORIES = {"Furniture"}

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
    """Parse cost strings like ' $ 25,00 ' → 25.0  (handles EU comma decimal)."""
    if not val or not val.strip():
        return None
    v = val.strip().replace("$", "").replace(" ", "")
    if "," in v and "." in v:   # 1,000.00 → thousands sep
        v = v.replace(",", "")
    elif "," in v:              # 25,00 → EU decimal
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


def aggregate_quantities(rows: list[dict], exclude_categories: set) -> tuple[dict, int]:
    """
    Filter out excluded categories, then aggregate per Assembly Code:
      area_sf, length_lf, volume_cf, count

    Returns (code_qtys dict, unmapped_count).
    unmapped = rows with no Assembly Code (blank).
    """
    code_qtys: dict[str, dict] = defaultdict(
        lambda: {"area_sf": 0.0, "length_lf": 0.0, "volume_cf": 0.0, "count": 0}
    )
    unmapped = 0

    for row in rows:
        cat = row.get("Category", "").strip()
        if cat in exclude_categories:
            continue

        ac = row.get("Assembly Code", "").strip()
        if not ac:
            unmapped += 1
            continue

        q = code_qtys[ac]
        q["area_sf"]   += parse_qty_str(row.get("Area", ""))
        q["length_lf"] += parse_qty_str(row.get("Length", ""))
        q["volume_cf"] += parse_qty_str(row.get("Volume", ""))
        q["count"]     += 1

    return dict(code_qtys), unmapped


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

def pick_quantity(cr: dict, code_qtys: dict) -> tuple[float, str]:
    """
    Determine the quantity to use for a cost line item.

    Rules:
      1. If Fixed Quantity is set → always use it (any cluster).
      2. If cluster is in TAKEOFF_CLUSTERS → derive from QTO by unit type.
      3. Otherwise → quantity = 0 (non-A/B/C cluster, no fixed qty → excluded).
    """
    fq = cr["fixed_qty"]
    if fq is not None:
        return fq, "Fixed"

    if cr["cluster"] not in TAKEOFF_CLUSTERS:
        return 0.0, "Fixed only (none set)"

    q = code_qtys.get(cr["ac"])
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


def calculate_costs(cost_data: list[dict], code_qtys: dict) -> list[dict]:
    """Apply unit costs to quantities and return enriched line items."""
    results = []
    for cr in cost_data:
        qty, qty_src = pick_quantity(cr, code_qtys)
        cost = cr["cost"]
        line_total = qty * cost if (cost is not None and qty) else 0.0

        notes = []
        if cost is None:
            notes.append("No unit cost")
        if qty == 0 and cr["fixed_qty"] is None and cr["cluster"] in TAKEOFF_CLUSTERS:
            notes.append("Zero qty from takeoff")
        if cr["ac"] not in code_qtys and cr["fixed_qty"] is None and cr["cluster"] in TAKEOFF_CLUSTERS:
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
    "Substructure":            "#6366f1",
    "Shell":                   "#0ea5e9",
    "Interiors":               "#10b981",
    "Services":                "#f59e0b",
    "Equipment and Furnishings":"#ec4899",
    "Special Contruction":     "#8b5cf6",
    "Building Sitework":       "#14b8a6",
}

def _cluster_color(name: str) -> str:
    return CLUSTER_COLORS.get(name, "#94a3b8")


def generate_html(
    results: list[dict],
    summary: list[dict],
    unmapped_count: int,
    data_source: str,
) -> str:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── summary data for Chart.js ─────────────────────────────────────────────
    chart_rows = [r for r in summary if r["cluster"] != "GRAND TOTAL"]
    grand_total = next(r["total"] for r in summary if r["cluster"] == "GRAND TOTAL")

    chart_labels  = json.dumps([r["cluster"] for r in chart_rows])
    chart_data    = json.dumps([r["total"] for r in chart_rows])
    chart_colors  = json.dumps([_cluster_color(r["cluster"]) for r in chart_rows])

    # ── summary cards ─────────────────────────────────────────────────────────
    cards_html = ""
    for r in summary:
        if r["cluster"] == "GRAND TOTAL":
            cards_html += f"""
            <div class="card grand-total">
              <div class="card-label">GRAND TOTAL</div>
              <div class="card-value">{fmt_usd(r['total'])}</div>
            </div>"""
        else:
            pct = r["total"] / grand_total * 100 if grand_total else 0
            color = _cluster_color(r["cluster"])
            cards_html += f"""
            <div class="card" style="border-top:4px solid {color}">
              <div class="card-label">{r['cluster']}</div>
              <div class="card-value">{fmt_usd(r['total'])}</div>
              <div class="card-pct">{pct:.1f}% of total</div>
            </div>"""

    # ── detail table rows ─────────────────────────────────────────────────────
    detail_rows_html = ""
    prev_cluster = None
    for r in results:
        if r["cluster"] != prev_cluster:
            bg = _cluster_color(r["cluster"])
            detail_rows_html += f"""
            <tr class="cluster-header" style="background:{bg}20;border-left:4px solid {bg}">
              <td colspan="9" style="font-weight:600;padding:6px 12px;color:{bg}">
                {r['cluster']}
              </td>
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

    # Unmapped row
    detail_rows_html += f"""
    <tr class="unmapped-row">
      <td class="mono">—</td>
      <td colspan="5">Unmapped elements (no Assembly Code in takeoff)</td>
      <td class="qty-src">Takeoff</td>
      <td class="num total-cell">—</td>
      <td><span class="note">{unmapped_count} elements — review required</span></td>
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
  body {{ font-family: system-ui, -apple-system, sans-serif; background: #f8fafc; color: #1e293b; }}
  header {{ background: #1e293b; color: #fff; padding: 24px 32px; display: flex; justify-content: space-between; align-items: center; }}
  header h1 {{ font-size: 1.4rem; font-weight: 700; letter-spacing: -.01em; }}
  header .meta {{ font-size: .8rem; color: #94a3b8; text-align: right; }}
  .container {{ max-width: 1400px; margin: 0 auto; padding: 28px 32px; }}
  /* cards */
  .cards {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 16px; margin-bottom: 32px; }}
  .card {{ background: #fff; border-radius: 10px; padding: 18px 20px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
  .card.grand-total {{ background: #1e293b; color: #fff; border-top: 4px solid #6366f1; }}
  .card-label {{ font-size: .72rem; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; color: #64748b; margin-bottom: 6px; }}
  .grand-total .card-label {{ color: #94a3b8; }}
  .card-value {{ font-size: 1.4rem; font-weight: 700; }}
  .card-pct {{ font-size: .78rem; color: #94a3b8; margin-top: 4px; }}
  /* chart */
  .chart-section {{ background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 32px; }}
  .section-title {{ font-size: .85rem; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: #64748b; margin-bottom: 18px; }}
  /* table */
  .table-section {{ background: #fff; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,.08); overflow: hidden; margin-bottom: 32px; }}
  .table-header {{ padding: 18px 24px; border-bottom: 1px solid #e2e8f0; }}
  table {{ width: 100%; border-collapse: collapse; font-size: .82rem; }}
  thead th {{ background: #f1f5f9; padding: 10px 12px; text-align: left; font-size: .72rem; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; color: #64748b; white-space: nowrap; }}
  tbody td {{ padding: 8px 12px; border-bottom: 1px solid #f1f5f9; vertical-align: top; }}
  tbody tr:last-child td {{ border-bottom: none; }}
  tbody tr:hover {{ background: #f8fafc; }}
  .cluster-header td {{ border-bottom: none !important; }}
  .zero-row {{ opacity: .55; }}
  .unmapped-row {{ background: #fff7ed; }}
  .unmapped-row td {{ font-style: italic; color: #92400e; }}
  .total-cell {{ font-weight: 600; color: #1e293b; }}
  .zero-row .total-cell {{ font-weight: normal; }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  .mono {{ font-family: monospace; font-size: .78rem; }}
  .qty-src {{ font-size: .72rem; color: #94a3b8; }}
  .note {{ font-size: .72rem; color: #ef4444; }}
  .badge-fixed {{ display: inline-block; background: #dbeafe; color: #1d4ed8; font-size: .65rem; font-weight: 600; border-radius: 4px; padding: 1px 5px; margin-left: 4px; text-transform: uppercase; }}
  footer {{ text-align: center; font-size: .75rem; color: #94a3b8; padding: 24px; }}
</style>
</head>
<body>
<header>
  <div>
    <h1>TVD Cost Dashboard</h1>
    <div style="font-size:.8rem;color:#94a3b8;margin-top:4px">Target Value Design — Cost Analysis</div>
  </div>
  <div class="meta">
    Generated: {ts}<br>
    Source: {data_source}
  </div>
</header>

<div class="container">

  <!-- Summary cards -->
  <div class="section-title" style="margin-bottom:14px">Cost by Cluster</div>
  <div class="cards">
    {cards_html}
  </div>

  <!-- Bar chart -->
  <div class="chart-section">
    <div class="section-title">Cost Distribution</div>
    <canvas id="clusterChart" height="80"></canvas>
  </div>

  <!-- Detail table -->
  <div class="table-section">
    <div class="table-header">
      <div class="section-title" style="margin-bottom:4px">Line Item Detail</div>
      <div style="font-size:.78rem;color:#64748b">
        Clusters A–C use takeoff quantities &nbsp;·&nbsp;
        Other clusters use fixed quantities only &nbsp;·&nbsp;
        <span style="color:#ef4444">Red text</span> = review flag &nbsp;·&nbsp;
        Dimmed rows = $0 line items &nbsp;·&nbsp;
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
  AutoTVD · {ts} · Data: {data_source}
</footer>

<script>
const ctx = document.getElementById('clusterChart');
new Chart(ctx, {{
  type: 'bar',
  data: {{
    labels: {chart_labels},
    datasets: [{{
      label: 'Total Cost (USD)',
      data: {chart_data},
      backgroundColor: {chart_colors},
      borderRadius: 6,
      borderSkipped: false,
    }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    plugins: {{
      legend: {{ display: false }},
      tooltip: {{
        callbacks: {{
          label: ctx => '$' + ctx.parsed.x.toLocaleString('en-US', {{maximumFractionDigits:0}})
        }}
      }}
    }},
    scales: {{
      x: {{
        ticks: {{
          callback: v => '$' + (v/1e6).toFixed(1) + 'M'
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
    code_qtys, unmapped_count = aggregate_quantities(all_elements, EXCLUDE_CATEGORIES)
    print(f"   Assembly codes in takeoff: {len(code_qtys)}")
    print(f"   Unmapped elements (no AC): {unmapped_count}")

    # 4. Load cost data
    cost_data = load_cost_data(cost_csv_rows)
    print(f"   Cost line items: {len(cost_data)}")

    # 5. Calculate line item costs
    results = calculate_costs(cost_data, code_qtys)

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

    html = generate_html(results, summary, unmapped_count, source)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\nDashboard saved: {out_path}")

    # 9. Open in browser (local mode only)
    if not ci_mode:
        webbrowser.open(f"file:///{out_path.replace(os.sep, '/')}")


if __name__ == "__main__":
    main()

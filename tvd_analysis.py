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

# Elements whose Family, Type, Mark, or Comments contain this marker (case-insensitive)
# are silently excluded from all quantity aggregation and cost calculations.
DNC_MARKER = "DNC"   # "Do Not Count"

# Finish quantity mirrors: cost AC → (source takeoff AC, quantity field)
# These finishes automatically track the element they're applied to.
#   C3010 Wall Paint        = same SF as interior walls  (C1010)
#   C3020 Floor Finishes    = same SF as interior floors (B1010)
QUANTITY_MIRRORS: dict[str, tuple[str, str]] = {
    "B3010": ("B1020", "area_sf"),
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

# n8n webhook URL — set this to your n8n HTTP trigger URL to receive budget-overrun alerts.
# Leave empty ("") to disable.
N8N_WEBHOOK_URL: str = "https://n8n.srv1447965.hstgr.cloud/webhook/ce8a4a9c-cc53-407e-9b72-2c35b3e73b42"

OUTPUT_HTML  = os.path.join(os.path.dirname(__file__), "TVD_Dashboard.html")
HISTORY_DIR  = os.path.join(os.path.dirname(__file__), "history")

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


def aggregate_quantities(rows: list[dict], exclude_categories: set) -> tuple[dict, int, dict, list[dict], int]:
    """
    Aggregate per Assembly Code for non-excluded categories:
      area_sf, length_lf, volume_cf, count

    Also builds all_ac_counts: element counts across ALL categories (including
    excluded ones like Furniture) — used for toilet stall mapping.

    Returns (code_qtys, unmapped_count, all_ac_counts, unmapped_rows, dnc_count).
    unmapped_rows = non-excluded rows with no Assembly Code, trimmed to export columns.
    dnc_count     = elements skipped because they carry the DNC_MARKER.
    """
    code_qtys: dict[str, dict] = defaultdict(
        lambda: {"area_sf": 0.0, "length_lf": 0.0, "volume_cf": 0.0, "count": 0}
    )
    all_ac_counts: dict[str, int] = defaultdict(int)
    unmapped = 0
    unmapped_rows: list[dict] = []
    dnc_count = 0

    for row in rows:
        # Skip elements marked "Do Not Count" in any name field
        dnc_marker = DNC_MARKER.upper()
        if any(
            dnc_marker in row.get(f, "").upper()
            for f in ("Family", "Type", "Mark", "Comments")
        ):
            dnc_count += 1
            continue

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

    return dict(code_qtys), unmapped, dict(all_ac_counts), unmapped_rows, dnc_count


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
    fixed_qty_by_ac: dict | None = None,
) -> tuple[float, str]:
    """
    Determine the quantity to use for a cost line item.

    Rules (in priority order):
      1. Fixed Quantity in cost_data → always wins.
      2. C1030 Toilet Partitions → count elements with TOILET_ACS (all categories).
      3. QUANTITY_MIRRORS entry → use a different AC's takeoff quantity, or its
         fixed_qty if the source AC has no takeoff data (e.g. B1020 is Fixed).
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

    # Rule 3 — finish mirrors: track another AC's quantity
    if ac in QUANTITY_MIRRORS:
        src_ac, field = QUANTITY_MIRRORS[ac]
        q = code_qtys.get(src_ac)
        if q:
            return q[field], f"Mirror: {src_ac} area"
        # Source AC may have a fixed qty instead of takeoff data
        if fixed_qty_by_ac and src_ac in fixed_qty_by_ac:
            return fixed_qty_by_ac[src_ac], f"Mirror: {src_ac} (Fixed)"
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

    fixed_qty_by_ac = {
        cr["ac"]: cr["fixed_qty"]
        for cr in cost_data
        if cr["fixed_qty"] is not None
    }

    results = []
    for cr in cost_data:
        qty, qty_src = pick_quantity(cr, code_qtys, all_ac_counts, fixed_qty_by_ac)
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
# N8N WEBHOOK
# ─────────────────────────────────────────────────────────────────────────────

def fire_budget_webhook(summary: list[dict], grand_total: float, target: float) -> None:
    """POST a budget-overrun alert to N8N_WEBHOOK_URL (if configured)."""
    if not N8N_WEBHOOK_URL:
        return
    delta = grand_total - target
    payload = json.dumps({
        "event":       "budget_overrun",
        "grand_total": round(grand_total, 2),
        "target":      round(target, 2),
        "delta":       round(delta, 2),
        "delta_pct":   round(delta / target * 100, 2) if target else 0,
        "timestamp":   datetime.now().isoformat(),
        "clusters":    [
            {"cluster": r["cluster"], "total": round(r["total"], 2)}
            for r in summary
        ],
    }).encode("utf-8")
    req = urllib.request.Request(
        N8N_WEBHOOK_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            print(f"   n8n webhook fired — HTTP {resp.status}")
    except urllib.error.URLError as exc:
        print(f"   n8n webhook failed: {exc}")


# ─────────────────────────────────────────────────────────────────────────────
# HISTORY SNAPSHOT I/O
# ─────────────────────────────────────────────────────────────────────────────

def save_snapshot(label: str, results: list[dict], summary: list[dict],
                  unmapped_count: int) -> str:
    """Save a named run snapshot to history/ as a dated JSON file."""
    os.makedirs(HISTORY_DIR, exist_ok=True)
    date_str = datetime.now().strftime("%Y-%m-%d")
    ts_str   = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe     = re.sub(r"[^a-z0-9]+", "_", label.lower()).strip("_")
    path     = os.path.join(HISTORY_DIR, f"{ts_str}_{safe}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({
            "label":         label,
            "date":          date_str,
            "results":       results,
            "summary":       summary,
            "unmapped_count": unmapped_count,
        }, f, indent=2)
    return path


def load_history() -> list[dict]:
    """Load all JSON snapshots from history/, sorted by filename (oldest first)."""
    if not os.path.isdir(HISTORY_DIR):
        return []
    versions = []
    for fname in sorted(os.listdir(HISTORY_DIR)):
        if not fname.endswith(".json"):
            continue
        path = os.path.join(HISTORY_DIR, fname)
        try:
            with open(path, "r", encoding="utf-8") as f:
                v = json.load(f)
            if "label" in v and "summary" in v:
                versions.append(v)
        except (json.JSONDecodeError, KeyError):
            print(f"   Warning: skipping unreadable history snapshot: {fname}")
    return versions


def _make_demo_snapshot(results: list[dict], unmapped_count: int) -> None:
    """
    Create a demo 'Test Version' snapshot in history/ (only if no snapshots
    exist yet).  A handful of line-item totals are scaled to simulate a
    different design iteration.
    """
    import copy
    results2 = copy.deepcopy(results)
    SCALE = {
        "B1010":    0.70,   # Structural Bamboo cheaper → simulate timber swap
        "B2010.CW": 1.30,   # Curtain Wall grew in area
        "D5010":    1.12,   # Electrical systems +12 %
        "D5090":    0.90,   # HVAC savings from passive design
        "C1010":    1.20,   # More partition SF in this iteration
    }
    for r in results2:
        if r["ac"] in SCALE:
            r["total"] = round(r["total"] * SCALE[r["ac"]], 2)

    # Rebuild summary from modified line items
    totals: dict[str, float] = {}
    seen: list[str] = []
    for r in results2:
        totals[r["cluster"]] = totals.get(r["cluster"], 0) + r["total"]
        if r["cluster"] not in seen:
            seen.append(r["cluster"])
    summary2 = [{"cluster": c, "total": totals[c]} for c in seen]
    summary2.append({"cluster": "GRAND TOTAL", "total": sum(totals.values())})

    path = save_snapshot("Test Version – Scheme A", results2, summary2, unmapped_count - 73)
    print(f"   Demo history snapshot created: {path}")


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




def generate_html(
    results: list[dict],
    summary: list[dict],
    unmapped_count: int,
    data_source: str,
    targets: dict[str, float],
    total_target: float,
    gross_sf: int = 30_000,
    unmapped_rows: list[dict] | None = None,
    history_versions: list[dict] | None = None,
) -> str:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── JS-embedded version data ──────────────────────────────────────────────
    _current_ver = {
        "label":         f"Current  ({ts})",
        "date":          ts,
        "results":       results,
        "summary":       summary,
        "unmapped_count": unmapped_count,
    }
    history_versions_js  = json.dumps(history_versions or [])
    current_version_js   = json.dumps(_current_ver)
    cluster_colors_js    = json.dumps(CLUSTER_COLORS)
    cluster_targets_js   = json.dumps(targets)
    gross_sf_js          = json.dumps(gross_sf)

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

    def _fmt_psf(n: float) -> str:
        return f"${n / gross_sf:,.0f}/SF"

    # ── cluster letter + scroll-target id ────────────────────────────────────
    cluster_letter: dict[str, str] = {}
    for _r in results:
        _c = _r["cluster"]
        if _c not in cluster_letter and _r["ac"]:
            cluster_letter[_c] = _r["ac"][0].upper()

    # ── chart data — only clusters that have a defined target ─────────────────
    cmp = [(r, targets[r["cluster"]]) for r in chart_rows if r["cluster"] in targets]
    ct_labels    = json.dumps([
        f"{cluster_letter.get(r['cluster'], '')} {r['cluster']}".strip()
        for r, _ in cmp
    ])
    ct_estimates = json.dumps([r["total"]   for r, _ in cmp])
    ct_targets   = json.dumps([t            for _, t in cmp])
    ct_deltas         = json.dumps([r["total"] - t for r, t in cmp])
    cluster_letter_js = json.dumps(cluster_letter)

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
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></script>
<style>
  /* ── colour tokens ─────────────────────────────────────────────────────────── */
  :root {{
    --bg:              #EDE8E3;
    --surface:         #FEFFFE;
    --surface-alt:     #EAE6E0;
    --surface-hover:   #F5F1EC;
    --text:            #424242;
    --text-muted:      #6B6B6B;
    --text-dim:        #A8A8A8;
    --border:          #E0DBD5;
    --accent:          #C46626;
    --header-bg:       #424242;
    --header-text:     #FEFFFE;
    --nav-bg:          #353535;
    --nav-text:        #A8A8A8;
    --shadow:          rgba(0,0,0,.10);
    --shadow-hover:    rgba(0,0,0,.14);
    --unmapped-bg:     #FFF5EA;
    --unmapped-text:   #7A4020;
    --unmapped-hover:  #FFE8D0;
    --badge-bg:        rgba(196,102,38,0.15);
    --badge-text:      #C46626;
    --total-card-bg:   #424242;
    --compare-label-bg:#EDE8E3;
  }}
  body.dark {{
    --bg:              #1a1a1c;
    --surface:         #2a2a2e;
    --surface-alt:     #35353a;
    --surface-hover:   #333338;
    --text:            #e4e4e4;
    --text-muted:      #9a9a9a;
    --text-dim:        #585858;
    --border:          #3c3c42;
    --accent:          #D4783A;
    --header-bg:       #111113;
    --header-text:     #e4e4e4;
    --nav-bg:          #0d0d0f;
    --nav-text:        #888888;
    --shadow:          rgba(0,0,0,.30);
    --shadow-hover:    rgba(0,0,0,.45);
    --unmapped-bg:     #28201a;
    --unmapped-text:   #D49070;
    --unmapped-hover:  #38281a;
    --badge-bg:        rgba(212,120,58,0.22);
    --badge-text:      #D4783A;
    --total-card-bg:   #1e1e22;
    --compare-label-bg:#2a2a2e;
  }}
  /* ── base ──────────────────────────────────────────────────────────────────── */
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: system-ui, -apple-system, sans-serif; background: var(--bg); color: var(--text); transition: background .2s, color .2s; }}
  /* ── header ────────────────────────────────────────────────────────────────── */
  header {{ background: var(--header-bg); color: var(--header-text); padding: 20px 32px; display: flex; justify-content: space-between; align-items: center; transition: background .2s; }}
  header h1 {{ font-size: 1.4rem; font-weight: 700; letter-spacing: -.01em; }}
  header .meta {{ font-size: .8rem; color: var(--nav-text); text-align: right; }}
  .header-subtitle {{ font-size: .8rem; color: var(--nav-text); margin-top: 4px; }}
  .header-right {{ display: flex; align-items: center; gap: 24px; }}
  .container {{ max-width: 1400px; margin: 0 auto; padding: 28px 32px; }}
  /* ── cards ─────────────────────────────────────────────────────────────────── */
  .cards {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 16px; margin-bottom: 32px; }}
  .card {{ background: var(--surface); border-radius: 10px; padding: 18px 20px; box-shadow: 0 1px 4px var(--shadow); transition: background .2s, box-shadow .2s; }}
  .card.grand-total {{ background: var(--total-card-bg); color: var(--header-text); border-top: 4px solid var(--accent); }}
  .card-link {{ cursor: pointer; transition: transform .12s, box-shadow .12s; }}
  .card-link:hover {{ transform: translateY(-2px); box-shadow: 0 6px 16px var(--shadow-hover); }}
  .card-label {{ font-size: .72rem; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; color: var(--text-muted); margin-bottom: 4px; }}
  .grand-total .card-label {{ color: var(--nav-text); }}
  .card-est {{ font-size: 1.35rem; font-weight: 700; margin: 2px 0 8px; }}
  .card-tvd-row {{ display: flex; justify-content: space-between; align-items: baseline; margin-top: 4px; }}
  .tvd-label {{ font-size: .72rem; color: var(--text-muted); }}
  .tvd-val {{ font-size: .82rem; font-weight: 600; color: var(--text-muted); }}
  .grand-total .tvd-label, .grand-total .tvd-val {{ color: var(--nav-text); }}
  .card-delta {{ font-size: .8rem; font-weight: 600; margin-top: 6px; }}
  .delta-over {{ color: #C44040; }}
  .delta-under {{ color: #7A9B76; }}
  .delta-neutral {{ color: var(--text-dim); }}
  .grand-total .delta-over {{ color: #E8A8A8; }}
  .grand-total .delta-under {{ color: #B5CEAD; }}
  body.dark .delta-over {{ color: #E89090; }}
  body.dark .delta-under {{ color: #96C490; }}
  /* ── charts ────────────────────────────────────────────────────────────────── */
  .chart-section {{ background: var(--surface); border-radius: 10px; padding: 24px; box-shadow: 0 1px 4px var(--shadow); margin-bottom: 16px; transition: background .2s; }}
  .chart-wrap {{ position: relative; }}
  .section-title {{ font-size: .85rem; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: var(--text-muted); margin-bottom: 18px; }}
  /* ── table ─────────────────────────────────────────────────────────────────── */
  .table-section {{ background: var(--surface); border-radius: 10px; box-shadow: 0 1px 4px var(--shadow); overflow: hidden; margin-bottom: 32px; margin-top: 16px; transition: background .2s; }}
  .table-header {{ padding: 18px 24px; border-bottom: 1px solid var(--border); }}
  table {{ width: 100%; border-collapse: collapse; font-size: .82rem; }}
  thead th {{ background: var(--surface-alt); padding: 10px 12px; text-align: left; font-size: .72rem; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; color: var(--text-muted); white-space: nowrap; }}
  tbody td {{ padding: 8px 12px; border-bottom: 1px solid var(--surface-alt); vertical-align: top; }}
  tbody tr:last-child td {{ border-bottom: none; }}
  tbody tr:hover {{ background: var(--surface-hover); }}
  .cluster-header td {{ border-bottom: none !important; }}
  .zero-row {{ opacity: .5; }}
  .unmapped-row {{ background: var(--unmapped-bg); }}
  .unmapped-row td {{ font-style: italic; color: var(--unmapped-text); }}
  .unmapped-link {{ cursor: pointer; transition: background .12s; }}
  .unmapped-link:hover {{ background: var(--unmapped-hover) !important; }}
  .total-cell {{ font-weight: 600; color: var(--text); }}
  .zero-row .total-cell {{ font-weight: normal; }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  .mono {{ font-family: monospace; font-size: .78rem; }}
  .qty-src {{ font-size: .72rem; color: var(--text-dim); }}
  .note {{ font-size: .72rem; color: #C44040; }}
  body.dark .note {{ color: #E89090; }}
  .badge-fixed {{ display: inline-block; background: var(--badge-bg); color: var(--badge-text); font-size: .65rem; font-weight: 600; border-radius: 4px; padding: 1px 5px; margin-left: 4px; text-transform: uppercase; }}
  footer {{ text-align: center; font-size: .75rem; color: var(--text-dim); padding: 24px; }}
  /* ── export PDF button ───────────────────────────────────────────────────── */
  .export-pdf-btn {{ display: flex; align-items: center; gap: 6px; padding: 7px 14px; background: var(--accent); color: #fff; border: none; border-radius: 6px; font-size: .78rem; font-weight: 600; cursor: pointer; letter-spacing: .02em; transition: opacity .15s, transform .1s; white-space: nowrap; flex-shrink: 0; }}
  .export-pdf-btn:hover {{ opacity: .88; transform: translateY(-1px); }}
  .export-pdf-btn:active {{ opacity: 1; transform: none; }}
  .export-pdf-btn svg {{ flex-shrink: 0; }}
  .export-pdf-btn.loading {{ opacity: .6; pointer-events: none; }}
  /* ── toggle buttons ────────────────────────────────────────────────────────── */
  .toggle-group {{ display: flex; border-radius: 6px; overflow: hidden; border: 1px solid var(--border); }}
  .toggle-btn {{ padding: 6px 18px; font-size: .78rem; font-weight: 600; background: var(--surface); color: var(--text-muted); border: none; cursor: pointer; transition: background .15s, color .15s; letter-spacing: .02em; }}
  .toggle-btn.active {{ background: var(--text); color: var(--surface); }}
  .toggle-btn:hover:not(.active) {{ background: var(--surface-hover); color: var(--text); }}
  .chart-section-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; }}
  /* ── mode switcher bar ───────────────────────────────────────────────────── */
  .mode-bar {{ background: var(--nav-bg); border-bottom: 2px solid var(--accent); padding: 0 32px; transition: background .2s; }}
  .mode-bar-inner {{ max-width: 1400px; margin: 0 auto; display: flex; }}
  .mode-btn {{ padding: 14px 28px; background: none; border: none; border-bottom: 3px solid transparent; color: var(--nav-text); font-size: .875rem; font-weight: 600; cursor: pointer; transition: color .15s, background .15s, border-color .15s; letter-spacing: .02em; }}
  .mode-btn.active {{ color: var(--header-text); border-bottom-color: var(--accent); }}
  .mode-btn:hover:not(.active) {{ color: var(--header-text); background: rgba(255,255,255,.06); }}
  /* ── mode content areas ──────────────────────────────────────────────────── */
  .mode-controls {{ display: flex; align-items: center; gap: 20px; flex-wrap: wrap; margin-bottom: 24px; padding: 14px 20px; background: var(--surface); border-radius: 10px; box-shadow: 0 1px 4px var(--shadow); transition: background .2s; }}
  .mode-controls label {{ font-size: .78rem; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: .04em; white-space: nowrap; }}
  .mode-controls select {{ padding: 7px 14px; border-radius: 6px; border: 1px solid var(--border); font-size: .84rem; color: var(--text); background: var(--surface); cursor: pointer; min-width: 220px; }}
  .compare-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
  .compare-col-label {{ font-size: .78rem; font-weight: 700; text-transform: uppercase; color: var(--text-muted); letter-spacing: .06em; text-align: center; padding: 10px 16px; background: var(--compare-label-bg); border-radius: 8px; margin-bottom: 14px; transition: background .2s; }}
  @media (max-width: 900px) {{ .compare-grid {{ grid-template-columns: 1fr; }} }}
  /* ── dark mode toggle switch ─────────────────────────────────────────────── */
  .theme-toggle {{ display: flex; align-items: center; gap: 7px; cursor: pointer; user-select: none; flex-shrink: 0; }}
  .theme-toggle input {{ position: absolute; opacity: 0; width: 0; height: 0; pointer-events: none; }}
  .theme-toggle-track {{ width: 40px; height: 22px; background: rgba(255,255,255,.22); border-radius: 11px; position: relative; transition: background .2s; flex-shrink: 0; }}
  .theme-toggle input:checked + .theme-toggle-track {{ background: var(--accent); }}
  .theme-toggle-thumb {{ width: 16px; height: 16px; background: #fff; border-radius: 50%; position: absolute; top: 3px; left: 3px; transition: transform .2s; box-shadow: 0 1px 3px rgba(0,0,0,.25); }}
  .theme-toggle input:checked + .theme-toggle-track .theme-toggle-thumb {{ transform: translateX(18px); }}
  .theme-toggle-icon {{ display: flex; align-items: center; color: rgba(255,255,255,.85); transition: color .2s, opacity .2s; }}
  #iconMoon {{ opacity: .38; }}
  body.dark #iconSun {{ opacity: .38; }}
  body.dark #iconMoon {{ opacity: 1; }}
  /* ── inline-style colour overrides for dark mode ─────────────────────────── */
  body.dark [style*="color:#A8A8A8"] {{ color: var(--text-dim) !important; }}
  body.dark [style*="color:#6B6B6B"] {{ color: var(--text-muted) !important; }}
  body.dark [style*="color:#424242"] {{ color: var(--text) !important; }}
  body.dark [style*="background:#EDE8E3"] {{ background: var(--compare-label-bg) !important; }}
</style>
</head>
<body>
<header>
  <div>
    <h1>Target Value Design Dashboard</h1>
    <div class="header-subtitle">Island Team 2026</div>
  </div>
  <div class="header-right">
    <button class="export-pdf-btn" id="exportPdfBtn" onclick="generatePDF()" title="Export executive summary as PDF">
      <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="15" y2="17"/><line x1="9" y1="9" x2="11" y2="9"/></svg>
      Export PDF
    </button>
    <label class="theme-toggle" title="Toggle dark mode">
      <span class="theme-toggle-icon" id="iconSun"><svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/></svg></span>
      <input type="checkbox" id="darkToggleInput" onchange="toggleDark(this.checked)">
      <span class="theme-toggle-track"><span class="theme-toggle-thumb"></span></span>
      <span class="theme-toggle-icon" id="iconMoon"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg></span>
    </label>
    <div class="meta">Last updated: {ts}</div>
  </div>
</header>

<div class="mode-bar">
  <div class="mode-bar-inner">
    <button class="mode-btn active" id="modeBtn-current"  onclick="switchMode('current')">Current</button>
    <button class="mode-btn"        id="modeBtn-history"  onclick="switchMode('history')">History</button>
    <button class="mode-btn"        id="modeBtn-compare"  onclick="switchMode('compare')">Compare</button>
  </div>
</div>

<div id="modeCurrentContent" class="container">

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

  <!-- TVD Targets Chart: Target / Estimate / Delta grouped vertical bars -->
  <div class="chart-section">
    <div class="chart-wrap" style="height:420px">
      <canvas id="tvdTargetsChart"></canvas>
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

<!-- ── History Mode ──────────────────────────────────────────────────────── -->
<div id="modeHistoryContent" class="container" style="display:none">
  <div class="chart-section" style="margin-bottom:28px">
    <div class="chart-wrap" style="height:300px">
      <canvas id="trackingChart"></canvas>
    </div>
  </div>
  <div class="mode-controls">
    <label for="historySelect">Version:</label>
    <select id="historySelect" onchange="renderHistoryMode()">
      <!-- populated by initHistoryMode() -->
    </select>
  </div>
  <div id="historyCards" class="cards" style="margin-bottom:32px"></div>
  <div id="historyChartsArea"></div>
  <div id="historyTableArea"></div>
</div>

<!-- ── Compare Mode ──────────────────────────────────────────────────────── -->
<div id="modeCompareContent" class="container" style="display:none">
  <div class="mode-controls">
    <div style="display:flex;align-items:center;gap:10px">
      <label for="compareSelectA">Version A:</label>
      <select id="compareSelectA" onchange="renderCompareMode()"></select>
    </div>
    <div style="display:flex;align-items:center;gap:10px">
      <label for="compareSelectB">Version B:</label>
      <select id="compareSelectB" onchange="renderCompareMode()"></select>
    </div>
  </div>
  <div id="compareArea"></div>
</div>

<footer>
  Island Team 2026 &middot; Last updated: {ts} &middot; Data: {data_source}
</footer>

<script>
// ── Embedded data ─────────────────────────────────────────────────────────────
const versionsData       = {history_versions_js};   // historic snapshots (array)
const currentVersionData = {current_version_js};    // this run
const CLUSTER_COLORS_JS  = {cluster_colors_js};
const CLUSTER_TARGETS_JS = {cluster_targets_js};
const CLUSTER_LETTER_JS  = {cluster_letter_js};
const GROSS_SF_JS        = {gross_sf_js};
const PIE_PALETTE_JS     = ["#F94144","#F3722C","#F8961E","#F9C74F",
                             "#90BE6D","#43AA8B","#577590","#415262"];

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

// ── Chart colour helper ───────────────────────────────────────────────────────
function _chartColors() {{
  const dark = document.body.classList.contains('dark');
  return {{
    text:   dark ? '#e4e4e4' : '#424242',
    grid:   dark ? '#35353a' : '#EAE6E0',
    border: dark ? '#2a2a2e' : '#FEFFFE',
  }};
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
  const cc = _chartColors();
  pieChart = new Chart(document.getElementById('pieChart'), {{
    type: 'doughnut',
    data: {{
      labels: d.labels,
      datasets: [{{
        data: d.data,
        backgroundColor: d.colors,
        borderColor: cc.border,
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
            color: cc.text,
            generateLabels: chart => {{
              const dataset = chart.data.datasets[0];
              return chart.data.labels.map((label, i) => {{
                const val = dataset.data[i];
                let suffix;
                if (pieMode === 'dollar') {{
                  const fmtVal = val >= 1e6
                    ? '$' + (val / 1e6).toFixed(1) + 'M'
                    : '$' + Math.round(val / 1e3) + 'k';
                  suffix = '  ' + fmtVal;
                }} else {{
                  suffix = '  ' + val.toFixed(1) + '%';
                }}
                return {{
                  text: label + suffix,
                  fillStyle: dataset.backgroundColor[i],
                  strokeStyle: cc.border,
                  fontColor: cc.text,
                  color: cc.text,
                  lineWidth: 2,
                  hidden: false,
                  index: i,
                }};
              }});
            }},
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

// ── TVD Targets Chart: Target / Estimate / Delta — grouped vertical bars ──────
function renderTVDTargetsChart() {{
  _killChart('tvdTargetsChart');
  const cc = _chartColors();
  new Chart(document.getElementById('tvdTargetsChart'), {{
    type: 'bar',
    data: {{
      labels: {ct_labels},
      datasets: [
        {{
          label: 'Target',
          data: {ct_targets},
          backgroundColor: '#70AD47',
          borderRadius: 4,
          borderSkipped: false,
        }},
        {{
          label: 'Estimate',
          data: {ct_estimates},
          backgroundColor: '#4472C4',
          borderRadius: 4,
          borderSkipped: false,
        }},
        {{
          label: 'Delta',
          data: {ct_deltas},
          backgroundColor: '#C44040',
          borderRadius: 4,
          borderSkipped: false,
        }}
      ]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        title: {{
          display: true,
          text: 'TVD \u2013 TARGETS BY CLUSTER',
          font: {{ size: 14, weight: 'bold' }},
          color: cc.text,
          padding: {{ bottom: 16 }}
        }},
        legend: {{ display: true, position: 'top', labels: {{ color: cc.text }} }},
        tooltip: {{
          callbacks: {{
            label: ctx => {{
              const v = ctx.parsed.y;
              const sign = v < 0 ? '\u2212' : '';
              return ctx.dataset.label + ': ' + sign + '$' + Math.round(Math.abs(v)).toLocaleString('en-US');
            }}
          }}
        }}
      }},
      scales: {{
        y: {{
          ticks: {{
            color: cc.text,
            callback: v => {{
              if (v === 0) return '$0';
              const sign = v < 0 ? '\u2212' : '';
              return sign + '$' + (Math.abs(v) / 1e6).toFixed(1) + 'M';
            }}
          }},
          grid: {{ color: cc.grid }}
        }},
        x: {{ ticks: {{ color: cc.text }}, grid: {{ display: false }} }}
      }}
    }}
  }});
}}
renderTVDTargetsChart();

// ── Mode switcher ─────────────────────────────────────────────────────────────
function switchMode(mode) {{
  ['current', 'history', 'compare'].forEach(m => {{
    document.getElementById('modeCurrentContent').style.display  = 'none';
    document.getElementById('modeHistoryContent').style.display  = 'none';
    document.getElementById('modeCompareContent').style.display  = 'none';
  }});
  document.getElementById('mode' + mode.charAt(0).toUpperCase() + mode.slice(1) + 'Content').style.display = '';
  ['current','history','compare'].forEach(m => {{
    document.getElementById('modeBtn-' + m).classList.toggle('active', m === mode);
  }});
  if (mode === 'history') initHistoryMode();
  if (mode === 'compare') initCompareMode();
}}

// ── Chart / display helpers ───────────────────────────────────────────────────
function _hexToRgba(hex, a) {{
  const h = hex.replace('#','');
  return 'rgba(' + parseInt(h.slice(0,2),16) + ',' + parseInt(h.slice(2,4),16) + ',' + parseInt(h.slice(4,6),16) + ',' + a + ')';
}}
function _fmtUSD(n) {{
  if (!n && n !== 0) return '\u2014';
  return '$' + Math.round(n).toLocaleString('en-US');
}}
function _fmtPSF(n) {{
  return '$' + Math.round(n / GROSS_SF_JS).toLocaleString('en-US') + '/SF';
}}
function _killChart(id) {{
  const c = document.getElementById(id);
  if (c) {{ const ch = Chart.getChart(c); if (ch) ch.destroy(); }}
}}

function buildVersionCharts(summary, idChart) {{
  const rows    = summary.filter(r => r.cluster !== 'GRAND TOTAL')
                         .filter(r => CLUSTER_TARGETS_JS[r.cluster] !== undefined);
  const labels  = rows.map(r => {{
    const letter = CLUSTER_LETTER_JS[r.cluster] || '';
    return (letter ? letter + ' ' : '') + r.cluster;
  }});
  const targets  = rows.map(r => CLUSTER_TARGETS_JS[r.cluster]);
  const estimates = rows.map(r => r.total);
  const deltas   = rows.map(r => r.total - CLUSTER_TARGETS_JS[r.cluster]);
  _killChart(idChart);
  const cc = _chartColors();
  new Chart(document.getElementById(idChart), {{
    type: 'bar',
    data: {{
      labels,
      datasets: [
        {{ label: 'Target',   data: targets,   backgroundColor: '#70AD47', borderRadius: 4, borderSkipped: false }},
        {{ label: 'Estimate', data: estimates, backgroundColor: '#4472C4', borderRadius: 4, borderSkipped: false }},
        {{ label: 'Delta',    data: deltas,    backgroundColor: '#C44040', borderRadius: 4, borderSkipped: false }}
      ]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        title: {{ display: true, text: 'TVD \u2013 TARGETS BY CLUSTER',
                  font: {{ size: 13, weight: 'bold' }}, color: cc.text, padding: {{ bottom: 14 }} }},
        legend: {{ display: true, position: 'top', labels: {{ color: cc.text }} }},
        tooltip: {{ callbacks: {{ label: ctx => {{
          const v = ctx.parsed.y;
          const sign = v < 0 ? '\u2212' : '';
          return ctx.dataset.label + ': ' + sign + '$' + Math.round(Math.abs(v)).toLocaleString('en-US');
        }} }} }}
      }},
      scales: {{
        y: {{ ticks: {{ color: cc.text, callback: v => v === 0 ? '$0' : (v < 0 ? '\u2212' : '') + '$' + (Math.abs(v)/1e6).toFixed(1) + 'M' }},
              grid: {{ color: cc.grid }} }},
        x: {{ ticks: {{ color: cc.text }}, grid: {{ display: false }} }}
      }}
    }}
  }});
}}

function buildCardsHtml(summary) {{
  const rows      = summary.filter(r => r.cluster !== 'GRAND TOTAL');
  const grandRow  = summary.find(r => r.cluster === 'GRAND TOTAL') || {{}};
  const grandTot  = grandRow.total || 0;
  const grandTgt  = Object.values(CLUSTER_TARGETS_JS).reduce((a,b)=>a+b,0);
  const gDelta    = grandTot - grandTgt;
  const gDcls     = gDelta>0?'delta-over':gDelta<0?'delta-under':'delta-neutral';
  const gSign     = gDelta>=0?'+':'\u2212';
  let html = '<div class="card grand-total">'
    + '<div class="card-label">Grand Total</div>'
    + '<div class="card-est">' + _fmtUSD(grandTot) + '</div>'
    + '<div class="card-tvd-row"><span class="tvd-label">Target</span><span class="tvd-val">' + _fmtUSD(grandTgt) + '</span></div>'
    + '<div class="card-delta ' + gDcls + '">' + gSign + '$' + Math.round(Math.abs(gDelta)).toLocaleString('en-US') + ' (' + (gDelta/grandTgt*100).toFixed(1) + '%)</div>'
    + '<div class="card-tvd-row" style="margin-top:6px"><span class="tvd-label">$/SF (' + GROSS_SF_JS.toLocaleString('en-US') + ' GSF)</span><span class="tvd-val">' + _fmtPSF(grandTot) + '</span></div>'
    + '</div>';
  for (const r of rows) {{
    const color = CLUSTER_COLORS_JS[r.cluster] || '#94a3b8';
    const tgt   = CLUSTER_TARGETS_JS[r.cluster];
    let tHtml = '';
    if (tgt !== undefined) {{
      const d = r.total - tgt;
      const dcls  = d>0?'delta-over':d<0?'delta-under':'delta-neutral';
      const dsign = d>=0?'+':'\u2212';
      tHtml = '<div class="card-tvd-row"><span class="tvd-label">Target</span><span class="tvd-val">' + _fmtUSD(tgt) + '</span></div>'
            + '<div class="card-delta ' + dcls + '">' + dsign + '$' + Math.round(Math.abs(d)).toLocaleString('en-US') + ' (' + (d/tgt*100).toFixed(1) + '%)</div>';
    }} else {{
      tHtml = '<div class="card-tvd-row" style="margin-top:6px"><span class="tvd-label" style="font-style:italic">No target set</span></div>';
    }}
    html += '<div class="card" style="border-top:4px solid ' + color + '">'
          + '<div class="card-label">' + r.cluster + '</div>'
          + '<div class="card-est">' + _fmtUSD(r.total) + '</div>'
          + tHtml
          + '</div>';
  }}
  return html;
}}

function buildTableHtml(results, unmappedCount) {{
  if (!results || !results.length) return '<p style="padding:20px;color:#A8A8A8">No line-item data in this snapshot.</p>';
  const clTotals = {{}};
  for (const r of results) clTotals[r.cluster] = (clTotals[r.cluster]||0) + (r.total||0);
  let prevCluster = null;
  let rows = '';
  for (const r of results) {{
    if (r.cluster !== prevCluster) {{
      const bg    = CLUSTER_COLORS_JS[r.cluster] || '#94a3b8';
      const clTot = clTotals[r.cluster] || 0;
      const clPsf = clTot ? _fmtPSF(clTot) : '\u2014';
      const letter= r.ac ? r.ac[0].toUpperCase() : '';
      rows += '<tr class="cluster-header" style="background:' + bg + '20;border-left:4px solid ' + bg + '">'
            + '<td colspan="7" style="font-weight:600;padding:6px 12px;color:' + bg + '">'
            + '<span style="font-family:monospace;font-weight:700;margin-right:10px;opacity:.55">' + letter + '</span>' + r.cluster
            + '</td>'
            + '<td class="num" style="padding:6px 12px;font-weight:700;color:' + bg + '">' + (clTot?_fmtUSD(clTot):'\u2014') + '</td>'
            + '<td class="num" style="padding:6px 12px;font-size:.72rem;color:' + bg + ';opacity:.75">' + clPsf + '</td>'
            + '</tr>';
      prevCluster = r.cluster;
    }}
    const zc  = r.total===0?' class="zero-row"':'';
    const fb  = r.qty_src==='Fixed'?' <span class="badge-fixed">fixed</span>':'';
    const nc  = r.notes?'<span class="note">'+r.notes+'</span>':'';
    const uc  = r.unit_cost!=null?('$'+Math.round(r.unit_cost).toLocaleString('en-US')):'\u2014';
    const tot = r.total?_fmtUSD(r.total):'\u2014';
    const qty = (r.qty||0).toLocaleString('en-US',{{minimumFractionDigits:1,maximumFractionDigits:1}});
    rows += '<tr'+zc+'>'
          + '<td class="mono">'+(r.ac||'')+'</td>'
          + '<td>'+(r.group||'')+'</td>'
          + '<td>'+(r.desc||'')+'</td>'
          + '<td class="mono">'+(r.unit||'')+'</td>'
          + '<td class="num">'+uc+'</td>'
          + '<td class="num">'+qty+fb+'</td>'
          + '<td class="qty-src">'+(r.qty_src||'')+'</td>'
          + '<td class="num total-cell">'+tot+'</td>'
          + '<td>'+nc+'</td>'
          + '</tr>';
  }}
  rows += '<tr class="unmapped-row">'
        + '<td class="mono">\u2014</td>'
        + '<td colspan="5">Unmapped elements (no Assembly Code in takeoff)</td>'
        + '<td class="qty-src">Takeoff</td>'
        + '<td class="num total-cell">\u2014</td>'
        + '<td><span class="note">' + (unmappedCount||0) + ' elements \u2014 review required</span></td>'
        + '</tr>';
  return '<div class="table-section"><div class="table-header"><div class="section-title" style="margin-bottom:4px">Line Item Detail</div></div>'
       + '<table><thead><tr><th>Assy Code</th><th>Group</th><th>Description</th><th>Unit</th>'
       + '<th class="num">Unit Cost</th><th class="num">Quantity</th><th>Qty Source</th>'
       + '<th class="num">Line Total</th><th>Notes</th></tr></thead>'
       + '<tbody>' + rows + '</tbody></table></div>';
}}

// ── History mode ──────────────────────────────────────────────────────────────
function buildTrackingChart() {{
  const grandTgt   = Object.values(CLUSTER_TARGETS_JS).reduce((a,b) => a+b, 0);
  const allVersions = [...versionsData, currentVersionData];
  const labels     = allVersions.map(v => v.label);
  const estimates  = allVersions.map(v => {{
    const g = (v.summary||[]).find(r => r.cluster === 'GRAND TOTAL');
    return g ? g.total : (v.summary||[]).reduce((s,r) => s+(r.total||0), 0);
  }});
  const deltas = estimates.map(e => e - grandTgt);
  _killChart('trackingChart');
  const cc = _chartColors();
  new Chart(document.getElementById('trackingChart'), {{
    type: 'bar',
    data: {{
      labels,
      datasets: [
        {{ label: 'Estimate', data: estimates, backgroundColor: '#4472C4', borderRadius: 4, borderSkipped: false }},
        {{ label: 'Delta',    data: deltas,    backgroundColor: '#C44040', borderRadius: 4, borderSkipped: false }}
      ]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        title: {{ display: true, text: 'TVD \u2013 TRACKING TARGET OVER TIME',
                  font: {{ size: 14, weight: 'bold' }}, color: cc.text, padding: {{ bottom: 16 }} }},
        legend: {{ display: true, position: 'top', labels: {{ color: cc.text }} }},
        tooltip: {{ callbacks: {{ label: ctx => {{
          const v = ctx.parsed.y;
          const sign = v < 0 ? '\u2212' : '';
          return ctx.dataset.label + ': ' + sign + '$' + Math.round(Math.abs(v)).toLocaleString('en-US');
        }} }} }}
      }},
      scales: {{
        y: {{
          ticks: {{ color: cc.text, callback: v => v === 0 ? '$0' : (v<0?'\u2212':'') + '$' + (Math.abs(v)/1e6).toFixed(1) + 'M' }},
          grid: {{ color: cc.grid }}
        }},
        x: {{
          ticks: {{ color: cc.text, maxRotation: 30, minRotation: 0 }},
          grid: {{ display: false }}
        }}
      }}
    }}
  }});
}}

function initHistoryMode() {{
  buildTrackingChart();
  const sel = document.getElementById('historySelect');
  if (sel.options.length > 0) {{ renderHistoryMode(); return; }}
  if (versionsData.length === 0) {{
    sel.innerHTML = '<option disabled>No snapshots saved yet</option>';
    document.getElementById('historyCards').innerHTML = '<p style="padding:20px;color:#A8A8A8">Run the script at least once to create a history snapshot.</p>';
    return;
  }}
  versionsData.forEach((v,i) => {{
    const o = document.createElement('option');
    o.value = i;
    o.text  = v.label + (v.date ? '  \u2014  ' + v.date : '');
    sel.appendChild(o);
  }});
  renderHistoryMode();
}}

function renderHistoryMode() {{
  const idx = parseInt(document.getElementById('historySelect').value);
  const v   = versionsData[idx];
  if (!v) return;
  document.getElementById('historyCards').innerHTML = buildCardsHtml(v.summary || []);
  document.getElementById('historyChartsArea').innerHTML =
    '<div class="chart-section" style="margin-bottom:16px">'
    + '<div class="chart-wrap" style="height:420px"><canvas id="hTvd"></canvas></div></div>';
  buildVersionCharts(v.summary || [], 'hTvd');
  document.getElementById('historyTableArea').innerHTML = buildTableHtml(v.results || [], v.unmapped_count || 0);
}}

// ── Compare mode ──────────────────────────────────────────────────────────────
function _getVersion(selId) {{
  const val = document.getElementById(selId).value;
  return val === '__current__' ? currentVersionData : versionsData[parseInt(val)];
}}

function _populateCompareSelect(selId, defaultCurrentSelected) {{
  const sel = document.getElementById(selId);
  if (sel.options.length > 0) return;
  const curOpt = document.createElement('option');
  curOpt.value = '__current__'; curOpt.text = 'Current';
  sel.appendChild(curOpt);
  versionsData.forEach((v,i) => {{
    const o = document.createElement('option'); o.value = i;
    o.text = v.label + (v.date ? '  \u2014  ' + v.date : '');
    sel.appendChild(o);
  }});
  if (!defaultCurrentSelected && sel.options.length > 1) sel.selectedIndex = 1;
}}

function initCompareMode() {{
  _populateCompareSelect('compareSelectA', true);
  _populateCompareSelect('compareSelectB', false);
  renderCompareMode();
}}

function _compareTotalHtml(v) {{
  const grandRow = (v.summary || []).find(r => r.cluster === 'GRAND TOTAL');
  const tot = grandRow ? grandRow.total : (v.summary||[]).reduce((s,r)=>s+(r.total||0),0);
  return '<div style="text-align:center;margin-bottom:14px;padding:10px 0;background:#F5F2EE;border-radius:8px">'
    + '<div style="font-size:.68rem;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#6B6B6B;margin-bottom:4px">Total Cost</div>'
    + '<div style="font-size:1.35rem;font-weight:700;color:#424242">$' + Math.round(tot).toLocaleString('en-US') + '</div>'
    + '</div>';
}}

function renderCompareMode() {{
  const vA = _getVersion('compareSelectA');
  const vB = _getVersion('compareSelectB');
  if (!vA || !vB) return;
  const labelA = vA.label || 'Version A';
  const labelB = vB.label || 'Version B';
  document.getElementById('compareArea').innerHTML =
    '<div class="compare-grid">'
    + '<div><div class="compare-col-label">' + labelA + '</div>'
    + _compareTotalHtml(vA)
    + '<div class="chart-section"><div class="chart-wrap" style="height:420px"><canvas id="cA_tvd"></canvas></div></div>'
    + '</div>'
    + '<div><div class="compare-col-label">' + labelB + '</div>'
    + _compareTotalHtml(vB)
    + '<div class="chart-section"><div class="chart-wrap" style="height:420px"><canvas id="cB_tvd"></canvas></div></div>'
    + '</div></div>';
  buildVersionCharts(vA.summary || [], 'cA_tvd');
  buildVersionCharts(vB.summary || [], 'cB_tvd');
}}

// ── Dark mode ─────────────────────────────────────────────────────────────────
function toggleDark(isDark) {{
  document.body.classList.toggle('dark', isDark);
  localStorage.setItem('tvd-dark', isDark ? '1' : '0');
  // re-render charts so their hardcoded text/grid colours update
  renderPieChart(pieMode);
  renderTVDTargetsChart();
  // rebuild tracking chart only if history tab is currently open
  if (document.getElementById('modeHistoryContent').style.display !== 'none') {{
    buildTrackingChart();
  }}
}}

(function () {{
  const saved      = localStorage.getItem('tvd-dark');
  const preferDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  const dark       = saved !== null ? saved === '1' : preferDark;
  if (dark) {{
    document.body.classList.add('dark');
    const cb = document.getElementById('darkToggleInput');
    if (cb) cb.checked = true;
  }}
}})();

// ── Executive Summary PDF Export ───────────────────────────────────────────────
function _hexRgb(hex) {{
  return [parseInt(hex.slice(1,3),16), parseInt(hex.slice(3,5),16), parseInt(hex.slice(5,7),16)];
}}

async function generatePDF() {{
  const btn = document.getElementById('exportPdfBtn');
  btn.classList.add('loading');
  btn.textContent = 'Generating…';
  await new Promise(r => setTimeout(r, 50));

  try {{
    const {{ jsPDF }} = window.jspdf;
    const doc = new jsPDF({{ orientation: 'portrait', unit: 'mm', format: 'a4' }});

    const W = 210, H = 297, M = 14, CW = W - M * 2;
    const cDark   = [66, 66, 66];
    const cAccent = [196, 102, 38];
    const cMuted  = [107, 107, 107];
    const cOver   = [196, 64, 64];
    const cUnder  = [90, 140, 86];
    const cLight  = [245, 241, 236];
    const cBorder = [224, 219, 213];

    const fmt  = n => '$' + Math.round(n).toLocaleString('en-US');
    const fmtS = n => {{ if (n >= 1e6) return '$' + (n/1e6).toFixed(2) + 'M'; if (n >= 1e3) return '$' + (n/1e3).toFixed(0) + 'K'; return fmt(n); }};

    function pageHeader(doc, title) {{
      doc.setFillColor(...cDark);
      doc.rect(0, 0, W, 22, 'F');
      doc.setFillColor(...cAccent);
      doc.rect(0, 22, W, 2.5, 'F');
      doc.setTextColor(254, 255, 254);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(12);
      doc.text(title, M, 14.5);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7.5);
      doc.setTextColor(168, 168, 168);
      doc.text('Island Team 2026  ·  TVD Dashboard', W - M, 14.5, {{ align: 'right' }});
    }}

    function sectionLabel(doc, text, y) {{
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.5);
      doc.setTextColor(...cMuted);
      doc.text(text.toUpperCase(), M, y);
      doc.setDrawColor(...cAccent);
      doc.setLineWidth(0.4);
      doc.line(M, y + 1.5, W - M, y + 1.5);
      return y + 7;
    }}

    function addFooters(doc) {{
      const n = doc.internal.getNumberOfPages();
      for (let i = 1; i <= n; i++) {{
        doc.setPage(i);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(6.5);
        doc.setTextColor(...cMuted);
        doc.setDrawColor(...cBorder);
        doc.setLineWidth(0.3);
        doc.line(M, H - 10, W - M, H - 10);
        doc.text('TVD Executive Summary  ·  Island Team 2026  ·  Generated ' + new Date().toLocaleDateString('en-US', {{year:'numeric',month:'long',day:'numeric'}}), M, H - 6);
        doc.text('Page ' + i + ' of ' + n, W - M, H - 6, {{ align: 'right' }});
      }}
    }}

    // ── data ──
    const clusterKeys  = Object.keys(CLUSTER_TARGETS_JS);
    const grandTarget  = Object.values(CLUSTER_TARGETS_JS).reduce((a,b) => a+b, 0);
    const curSummary   = currentVersionData.summary || [];
    const grandTotal   = (curSummary.find(r => r.cluster === 'GRAND TOTAL') || {{}}).total || 0;
    const grandDelta   = grandTotal - grandTarget;
    const grandPct     = (grandDelta / grandTarget * 100).toFixed(1);
    const curDate      = (currentVersionData.date || '').replace(/\s.*/, '');

    // ── PAGE 1: Cover + Cluster Summary ──
    pageHeader(doc, 'Target Value Design — Executive Summary');
    let y = 32;

    y = sectionLabel(doc, 'Key Metrics', y);
    const mBoxW = (CW - 9) / 4;
    const metrics = [
      {{ label: 'TOTAL ESTIMATE',  value: fmtS(grandTotal),  sub: curDate,          col: cDark   }},
      {{ label: 'TVD TARGET',      value: fmtS(grandTarget), sub: '30,000 GSF',     col: cMuted  }},
      {{ label: 'DELTA',           value: (grandDelta < 0 ? '−' : '+') + fmtS(Math.abs(grandDelta)), sub: grandPct + '%', col: grandDelta < 0 ? cUnder : cOver }},
      {{ label: '$ / SF',          value: '$' + Math.round(grandTotal / 30000),     sub: 'per Gross SF',   col: cAccent }},
    ];
    metrics.forEach((m, i) => {{
      const bx = M + i * (mBoxW + 3);
      doc.setFillColor(...cLight);
      doc.roundedRect(bx, y, mBoxW, 22, 2, 2, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(6.5); doc.setTextColor(...cMuted);
      doc.text(m.label, bx + mBoxW / 2, y + 6, {{ align: 'center' }});
      doc.setFontSize(13); doc.setTextColor(...m.col);
      doc.text(m.value, bx + mBoxW / 2, y + 15, {{ align: 'center' }});
      doc.setFontSize(6.5); doc.setTextColor(...cMuted);
      doc.text(m.sub, bx + mBoxW / 2, y + 20.5, {{ align: 'center' }});
    }});
    y += 28;

    y = sectionLabel(doc, 'Cluster Summary — Current (' + curDate + ')', y);
    const clRows = clusterKeys.map(cl => {{
      const s   = curSummary.find(r => r.cluster === cl) || {{}};
      const est = s.total || 0;
      const tgt = CLUSTER_TARGETS_JS[cl] || 0;
      const d   = est - tgt;
      const p   = tgt ? (d / tgt * 100).toFixed(1) : '0.0';
      return [cl, fmt(est), fmt(tgt), (d < 0 ? '−' : '+') + fmt(Math.abs(d)) + ' (' + p + '%)'];
    }});
    clRows.push(['GRAND TOTAL', fmt(grandTotal), fmt(grandTarget),
      (grandDelta < 0 ? '−' : '+') + fmt(Math.abs(grandDelta)) + ' (' + grandPct + '%)']);

    doc.autoTable({{
      startY: y,
      head: [['Cluster', 'Estimate', 'Target', 'Delta vs. Target']],
      body: clRows,
      margin: {{ left: M, right: M }},
      styles:      {{ fontSize: 8, cellPadding: 3 }},
      headStyles:  {{ fillColor: cDark, textColor: [254,255,254], fontStyle: 'bold' }},
      columnStyles: {{ 0: {{ cellWidth: 58 }}, 1: {{ halign:'right' }}, 2: {{ halign:'right' }}, 3: {{ halign:'right' }} }},
      didParseCell(d) {{
        if (d.section !== 'body') return;
        if (d.column.index === 3) {{
          const v = String(d.cell.raw);
          d.cell.styles.textColor = v.startsWith('−') ? cUnder : cOver;
        }}
        if (d.row.index === clRows.length - 1) {{
          d.cell.styles.fontStyle = 'bold';
          d.cell.styles.fillColor = [234,230,224];
        }}
      }},
    }});
    y = doc.lastAutoTable.finalY + 6;

    const headroomPct = Math.abs(parseFloat(grandPct));
    const statusText  = grandDelta < 0
      ? 'Current estimate is ' + fmt(Math.abs(grandDelta)) + ' (' + headroomPct + '%) UNDER the TVD target — ' + fmt(Math.abs(grandDelta)) + ' of headroom available.'
      : 'Current estimate EXCEEDS the TVD target by ' + fmt(grandDelta) + ' (' + headroomPct + '%). Value engineering required.';
    doc.setFillColor(...(grandDelta < 0 ? [234,244,232] : [252,236,236]));
    doc.roundedRect(M, y, CW, 10, 2, 2, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5);
    doc.setTextColor(...(grandDelta < 0 ? cUnder : cOver));
    doc.text(statusText, M + 4, y + 6.5);
    y += 16;

    // ── PAGE 2: Milestone Comparison ──
    doc.addPage();
    pageHeader(doc, 'Project Milestone Comparison');
    y = 32;

    const milestones = [...versionsData, currentVersionData];
    const msLabels   = milestones.map(v => {{
      const raw = v.label || v.date || '';
      return raw.replace(/Current\s+\(.*\)/, 'Current').replace(/\s*\(.*\)/, '').trim();
    }});

    y = sectionLabel(doc, 'Cost Evolution by Cluster Across Milestones', y);
    const msHead = ['Cluster / Milestone', ...msLabels];
    const msBody = clusterKeys.map(cl => {{
      const row = [cl];
      milestones.forEach(v => {{
        const s = (v.summary || []).find(r => r.cluster === cl);
        row.push(s ? fmt(s.total) : '—');
      }});
      return row;
    }});
    const gtRow = ['GRAND TOTAL'];
    milestones.forEach(v => {{
      const s = (v.summary || []).find(r => r.cluster === 'GRAND TOTAL');
      gtRow.push(s ? fmt(s.total) : '—');
    }});
    msBody.push(gtRow);
    const tgtRow = ['TVD TARGET'];
    milestones.forEach(() => tgtRow.push(fmt(grandTarget)));
    msBody.push(tgtRow);

    const lastColIdx = msLabels.length;
    doc.autoTable({{
      startY: y,
      head: [msHead],
      body: msBody,
      margin: {{ left: M, right: M }},
      styles:      {{ fontSize: 7, cellPadding: 2.5, overflow: 'linebreak' }},
      headStyles:  {{ fillColor: cDark, textColor: [254,255,254], fontStyle: 'bold', halign: 'right' }},
      columnStyles: {{ 0: {{ cellWidth: 50, halign: 'left' }} }},
      didParseCell(d) {{
        if (d.column.index > 0) d.cell.styles.halign = 'right';
        if (d.column.index === lastColIdx && d.section === 'body') d.cell.styles.fontStyle = 'bold';
        if (d.section === 'body') {{
          const ri = d.row.index;
          if (ri === msBody.length - 1) {{ d.cell.styles.textColor = cAccent; d.cell.styles.fontStyle = 'italic'; d.cell.styles.fillColor = [254,248,240]; }}
          else if (ri === msBody.length - 2) {{ d.cell.styles.fontStyle = 'bold'; d.cell.styles.fillColor = [234,230,224]; }}
        }}
      }},
    }});
    y = doc.lastAutoTable.finalY + 10;

    y = sectionLabel(doc, 'Grand Total Trend', y);
    const gtVals = milestones.map(v => {{
      const s = (v.summary || []).find(r => r.cluster === 'GRAND TOTAL');
      return {{ label: msLabels[milestones.indexOf(v)], val: s ? s.total : 0 }};
    }});
    const barH2  = 5;
    const barMax = Math.max(...gtVals.map(g => g.val), grandTarget) * 1.05;
    gtVals.forEach((g, i) => {{
      const bw = CW * (g.val / barMax);
      const by = y + i * (barH2 + 4);
      const isLast = i === gtVals.length - 1;
      doc.setFillColor(...(isLast ? cAccent : [200,195,188]));
      doc.roundedRect(M, by, bw, barH2, 1, 1, 'F');
      doc.setFont('helvetica', isLast ? 'bold' : 'normal');
      doc.setFontSize(7);
      doc.setTextColor(...(isLast ? cAccent : cMuted));
      doc.text(g.label, M - 1, by + barH2 - 0.5, {{ align: 'right' }});
      doc.text(fmt(g.val), M + bw + 2, by + barH2 - 0.5);
    }});
    const tgtX = M + CW * (grandTarget / barMax);
    doc.setDrawColor(...cAccent);
    doc.setLineWidth(0.4);
    doc.setLineDashPattern([2, 1.5], 0);
    doc.line(tgtX, y - 2, tgtX, y + gtVals.length * (barH2 + 4));
    doc.setLineDashPattern([], 0);
    doc.setFont('helvetica', 'italic'); doc.setFontSize(6.5); doc.setTextColor(...cAccent);
    doc.text('Target ' + fmt(grandTarget), tgtX + 1.5, y - 3);

    // ── PAGE 3: Charts ──
    doc.addPage();
    pageHeader(doc, 'Cost Visualization');
    y = 32;

    y = sectionLabel(doc, 'Cost Distribution by Cluster', y);
    try {{
      const pieCanvas = document.getElementById('pieChart');
      const pieImg    = pieCanvas.toDataURL('image/png');
      const pieRatio  = pieCanvas.height / pieCanvas.width;
      const pieW = 72, pieH = pieW * pieRatio;
      doc.addImage(pieImg, 'PNG', M, y, pieW, pieH);

      let legY = y + 4;
      const legX = M + pieW + 10;
      clusterKeys.forEach(cl => {{
        const col = _hexRgb(CLUSTER_COLORS_JS[cl] || '#888');
        const s   = curSummary.find(r => r.cluster === cl) || {{}};
        const pct = grandTotal > 0 ? ((s.total || 0) / grandTotal * 100).toFixed(1) : '0.0';
        doc.setFillColor(...col);
        doc.roundedRect(legX, legY - 2.5, 4, 4, 0.5, 0.5, 'F');
        doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(...cDark);
        doc.text(cl, legX + 6, legY);
        doc.setFont('helvetica', 'bold'); doc.setTextColor(...cMuted);
        doc.text(fmt(s.total || 0) + '  (' + pct + '%)', W - M, legY, {{ align: 'right' }});
        legY += 7.5;
      }});
      y = Math.max(y + pieH, legY) + 8;
    }} catch(e) {{ y += 5; }}

    y = sectionLabel(doc, 'TVD Targets vs. Estimates by Cluster', y);
    try {{
      const barCanvas = document.getElementById('tvdTargetsChart');
      const barImg    = barCanvas.toDataURL('image/png');
      const bRatio    = barCanvas.height / barCanvas.width;
      const bW = CW, bH = Math.min(bW * bRatio, H - y - 20);
      doc.addImage(barImg, 'PNG', M, y, bW, bH);
      y += bH + 8;
    }} catch(e) {{}}

    // ── PAGE 4+: Line Item Detail ──
    doc.addPage();
    pageHeader(doc, 'Line Item Detail');
    y = 32;
    y = sectionLabel(doc, 'All Line Items Grouped by Cluster (Current Estimate)', y);

    const results = currentVersionData.results || [];
    const grouped = {{}};
    results.forEach(r => {{ (grouped[r.cluster] = grouped[r.cluster] || []).push(r); }});

    const liHead = [['Code', 'Description', 'Unit', 'Unit Cost', 'Qty', 'Line Total']];
    const liBody = [];
    clusterKeys.forEach(cl => {{
      const items   = grouped[cl] || [];
      const clTotal = (curSummary.find(s => s.cluster === cl) || {{}}).total || 0;
      const clColor = _hexRgb(CLUSTER_COLORS_JS[cl] || '#888');
      liBody.push([
        {{ content: cl, colSpan: 5, styles: {{ fontStyle: 'bold', fillColor: [234,230,224], textColor: cDark, fontSize: 8 }} }},
        {{ content: fmt(clTotal), styles: {{ fontStyle: 'bold', fillColor: [234,230,224], textColor: clColor, halign: 'right', fontSize: 8 }} }},
      ]);
      items.filter(r => r.total > 0).forEach(r => {{
        liBody.push([
          r.ac || '',
          (r.desc || r.group || '').substring(0, 70),
          r.unit || '',
          r.unit_cost != null ? '$' + Number(r.unit_cost).toLocaleString('en-US') : '',
          r.qty != null ? Number(r.qty).toLocaleString('en-US', {{maximumFractionDigits:1}}) : '',
          fmt(r.total),
        ]);
      }});
    }});

    doc.autoTable({{
      startY: y,
      head: liHead,
      body: liBody,
      margin: {{ left: M, right: M }},
      styles:      {{ fontSize: 6.8, cellPadding: 2, overflow: 'linebreak' }},
      headStyles:  {{ fillColor: cDark, textColor: [254,255,254], fontStyle: 'bold' }},
      columnStyles: {{
        0: {{ cellWidth: 17 }},
        1: {{ cellWidth: 72 }},
        2: {{ cellWidth: 12, halign: 'center' }},
        3: {{ cellWidth: 22, halign: 'right' }},
        4: {{ cellWidth: 18, halign: 'right' }},
        5: {{ cellWidth: 21, halign: 'right', fontStyle: 'bold' }},
      }},
    }});

    addFooters(doc);
    doc.save('TVD_Executive_Summary_' + new Date().toISOString().slice(0,10) + '.pdf');

  }} finally {{
    btn.classList.remove('loading');
    btn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="15" y2="17"/><line x1="9" y1="9" x2="11" y2="9"/></svg> Export PDF';
  }}
}}
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
    parser.add_argument(
        "--snapshot",
        metavar="LABEL",
        help='Save a named snapshot of this run to history/ before generating the dashboard. '
             'Example: --snapshot "Scheme A – Week 12"',
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

    # 3. Aggregate takeoff quantities (excluding furnishings and DNC elements)
    code_qtys, unmapped_count, all_ac_counts, unmapped_rows, dnc_count = aggregate_quantities(
        all_elements, EXCLUDE_CATEGORIES
    )
    toilet_count = sum(all_ac_counts.get(t, 0) for t in TOILET_ACS)
    print(f"   Assembly codes in takeoff: {len(code_qtys)}")
    print(f"   Unmapped elements (no AC): {unmapped_count}")
    print(f"   Skipped (DNC marker):      {dnc_count}")
    print(f"   Toilet elements found (C1030): {toilet_count}")

    # 4. Load cost data
    cost_data = load_cost_data(cost_csv_rows)
    print(f"   Cost line items: {len(cost_data)}")

    # 5. Calculate line item costs
    results = calculate_costs(cost_data, code_qtys, all_ac_counts)

    # 6. Build cluster summary
    summary = build_cluster_summary(results)

    # 6b. Fire n8n webhook if grand total exceeds target
    grand_total = next((r["total"] for r in summary if r["cluster"] == "GRAND TOTAL"), 0.0)
    if grand_total > TOTAL_TARGET:
        fire_budget_webhook(summary, grand_total, TOTAL_TARGET)

    # 7. Print summary to console
    print()
    print(f"  {'Cluster':<35} {'Total':>14}")
    print("  " + "─" * 51)
    for r in summary:
        marker = " ◄" if r["cluster"] == "GRAND TOTAL" else ""
        print(f"  {r['cluster']:<35} {fmt_usd(r['total']):>14}{marker}")
    print(f"\n  Unmapped elements: {unmapped_count} (no Assembly Code)")

    # 8. Save explicit snapshot if requested, then load history
    if args.snapshot:
        snap_path = save_snapshot(args.snapshot, results, summary, unmapped_count)
        print(f"\n   Snapshot saved: {snap_path}")

    history = load_history()
    if not history and not ci_mode:
        print("\n   No history snapshots found — creating demo snapshot...")
        _make_demo_snapshot(results, unmapped_count)
        history = load_history()

    # 9. Generate HTML dashboard
    if ci_mode:
        out_dir = os.path.join(LOCAL_BASE, "docs")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, "index.html")
    else:
        out_path = OUTPUT_HTML

    html = generate_html(results, summary, unmapped_count, source,
                         CLUSTER_TARGETS, TOTAL_TARGET, GROSS_SF, unmapped_rows,
                         history)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\nDashboard saved: {out_path}")
    print(f"History snapshots loaded: {len(history)}")

    # 10. Open in browser (local mode only)
    if not ci_mode:
        webbrowser.open(f"file:///{out_path.replace(os.sep, '/')}")


if __name__ == "__main__":
    main()

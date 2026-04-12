# AutoTVD

**Island Team 2026 · Stanford AEC Global Teamwork**

AutoTVD is a Target Value Design (TVD) cost tracking tool for the ISLAND 2026 project. It reads Revit quantity take-off (QTO) schedules exported as CSV files, applies unit costs from a cost database, and generates an interactive HTML dashboard to compare the current cost estimate against TVD targets — by cluster and line item.

The dashboard is automatically deployed to GitHub Pages on every push to `main` via GitHub Actions.

---

## How it works

1. **QTO CSVs** (`qto/`) — Revit schedules exported as CSV, one per discipline. Each row is one Revit element and must carry an `Assembly Code` (CSI Uniformat) column.
2. **Cost database** (`cost_data.csv`) — maps every Assembly Code to a unit cost, cluster, and unit type. Fixed quantities can override the takeoff.
3. **`tvd_analysis.py`** — merges the takeoffs, aggregates quantities per Assembly Code, multiplies by unit cost, and renders the HTML dashboard.
4. **`docs/index.html`** — the generated dashboard, served via GitHub Pages.

### Quantity rules (in priority order)

| Priority | Rule | When it applies |
|---|---|---|
| 1 | **Fixed Quantity** | `cost_data.csv` has a value in the `Fixed Quantity` column — always wins |
| 2 | **Toilet stall count** | Assembly Code `C1030` — counts elements with codes in `TOILET_ACS` |
| 3 | **Quantity mirror** | Code is in `QUANTITY_MIRRORS` — borrows area/length from another AC |
| 4 | **Takeoff lookup** | Clusters A–C (Substructure, Shell, Interiors) — quantity from the QTO |
| 5 | **Zero** | All other clusters without a Fixed Quantity |

---

## Repository structure

```
AutoTVD/
├── qto/
│   ├── Architecture_TakeOff.csv   # Architectural Revit QTO schedule
│   └── Structural_Schedule.csv    # Structural Revit QTO schedule
├── cost_data.csv                  # Unit costs, clusters, fixed quantities
├── tvd_analysis.py                # Main script — generates the dashboard
├── .github/workflows/deploy.yml   # CI: runs script on push, deploys to Pages
├── history/                       # Named snapshots — committed to repo so they appear on the deployed dashboard
└── docs/
    └── index.html                 # Generated dashboard (deployed to Pages)
```

---

## Updating a take-off to a newer version

Use this workflow whenever you export a fresh QTO from Revit and want the old version to be preserved and accessible in the History / Compare modes of the dashboard.

### Step 1 — Checkpoint the current version

With the **old** take-off still in `qto/`, run:

```bash
python tvd_analysis.py --snapshot "Scheme A – Week 12"
```

This calculates costs from the existing files, saves a snapshot to `history/`, and opens the dashboard as usual. Pick a label that identifies this design iteration clearly — it will appear in the History and Compare dropdowns.

### Step 2 — Commit the snapshot

```bash
git add history/
git commit -m "Snapshot: Scheme A – Week 12"
```

Committing the snapshot file is what makes it available on the deployed GitHub Pages dashboard. Snapshots in `history/` that are not committed only exist on your machine.

### Step 3 — Replace the take-off file

Overwrite the file in `qto/` with the new Revit export, keeping the **same filename**:

```
qto/Architecture_TakeOff.csv   ← replace with new export
qto/Structural_Schedule.csv    ← replace with new export (if updated)
```

Do not rename the file — the script looks for these exact paths.

### Step 4 — Verify locally

```bash
python tvd_analysis.py
```

The dashboard opens with the new data as **Current**. Switch to **History** or **Compare** in the mode bar to confirm the old snapshot appears.

### Step 5 — Commit and push

```bash
git add qto/Architecture_TakeOff.csv
git commit -m "Update arch take-off – Scheme B Week 14"
git push
```

Pushing a change to any file in `qto/` triggers the GitHub Actions workflow, which re-runs the script, embeds all committed snapshots from `history/`, and redeploys the dashboard to GitHub Pages.

---

## Required columns in a take-off CSV

Every QTO file loaded by the script must contain these columns (names are case-sensitive):

| Column | Notes |
|---|---|
| `ElementId` | Unique Revit element ID — used to deduplicate across discipline files |
| `Category` | Revit category (e.g. `Walls`, `Floors`, `Structural Columns`) |
| `Family` | Revit family name |
| `Type` | Revit type name |
| `Assembly Code` | CSI Uniformat code (e.g. `B2010`) — must be set on every element |
| `Level` | Level the element belongs to |
| `Area` | Area in SF (walls, floors) |
| `Length` | Length in LF (linear elements) |
| `Volume` | Volume in CF (concrete, mass elements) |

Other columns (`Mark`, `Material`, `Comments`, etc.) are preserved and included in the unmapped-elements download for Revit review.

The `Unit` column in `cost_data.csv` controls how the quantity is extracted from the take-off:

| Unit | Quantity used |
|---|---|
| `SF` / `GSF` | Area |
| `LF` | Length |
| `EA` / `Flight` | Element count |
| `CY` | Volume ÷ 27 |
| `CF` | Volume |
| `MSF` | Area ÷ 1000 |

---

## Keyword-based Assembly Code splitting

Some Revit families share the same Assembly Code but need to map to different cost line items (e.g. curtain walls and plaster walls both tagged `B2010`). The `AC_KEYWORD_SPLIT` config in `tvd_analysis.py` handles this by inspecting the element's Category + Family + Type name:

```python
AC_KEYWORD_SPLIT = {
    "B2010": [
        (["storefront", "curtain wall", "curtain", "glazing"], "B2010.CW"),
        ([], "B2010.PW"),   # catch-all fallback
    ],
}
```

Add a matching pair of rows to `cost_data.csv` using the synthetic sub-codes (`B2010.CW`, `B2010.PW`).

---

## Configuration reference (`tvd_analysis.py`)

| Constant | Purpose |
|---|---|
| `CLUSTER_TARGETS` | TVD budget targets per cluster — update each design iteration |
| `TOTAL_TARGET` | Overall project TVD target |
| `GROSS_SF` | Gross square footage used for $/SF indices |
| `TAKEOFF_CLUSTERS` | Clusters that use QTO quantities (all others require Fixed Quantity) |
| `QUANTITY_MIRRORS` | Maps one AC to borrow its quantity from another AC's takeoff |
| `AC_KEYWORD_SPLIT` | Splits a single AC into sub-codes based on element name keywords |
| `EXCLUDE_CATEGORIES` | Revit categories excluded from quantity aggregation (e.g. Furniture) |
| `TOILET_ACS` | Assembly codes used to count toilet stalls for C1030 |

---

## Dashboard modes

| Mode | Description |
|---|---|
| **Current** | Live view of the most recent CI run — cards, donut chart, TVD targets chart, line item table |
| **History** | Full dashboard view for any committed snapshot in `history/` |
| **Compare** | Side-by-side chart comparison of any two versions, including Current |

Snapshots are stored as JSON files in `history/`. Committed snapshots are embedded in the deployed dashboard. Only snapshots committed to the repo appear on GitHub Pages — local-only snapshots are visible when running the dashboard locally.

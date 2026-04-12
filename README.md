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
├── history/                       # Auto-generated run snapshots (gitignored)
└── docs/
    └── index.html                 # Generated dashboard (deployed to Pages)
```

---

## Adding a new take-off file

### Step 1 — Export from Revit

Export the schedule as a **CSV (comma-separated)** file. The schedule must include these columns (column names are case-sensitive):

| Column | Notes |
|---|---|
| `ElementId` | Unique Revit element ID — used to deduplicate across files |
| `Category` | Revit category (e.g. `Walls`, `Floors`, `Structural Columns`) |
| `Family` | Revit family name |
| `Type` | Revit type name |
| `Assembly Code` | CSI Uniformat code (e.g. `B2010`) — **must be set on every element** |
| `Level` | Level the element belongs to |
| `Area` | Area in SF (for wall/floor elements) |
| `Length` | Length in LF (for linear elements) |
| `Volume` | Volume in CF (for concrete/mass elements) |

Other columns (`Mark`, `Material`, `Comments`, etc.) are carried through to the unmapped-elements export and can be left blank.

### Step 2 — Place the file in `qto/`

Save the exported CSV to the `qto/` folder with a descriptive filename, e.g. `qto/MEP_TakeOff.csv`.

### Step 3 — Register it in `tvd_analysis.py`

Open `tvd_analysis.py` and find the `QTO_FILES` and `LOCAL_QTO` dicts near the top of the file:

```python
QTO_FILES = {
    "arch":   "qto/Architecture_TakeOff.csv",
    "struct": "qto/Structural_Schedule.csv",
}

LOCAL_QTO = {
    "arch":   os.path.join(LOCAL_BASE, "qto", "Architecture_TakeOff.csv"),
    "struct": os.path.join(LOCAL_BASE, "qto", "Structural_Schedule.csv"),
}
```

Add a new entry for your file in both dicts (use any unique key):

```python
QTO_FILES = {
    "arch":   "qto/Architecture_TakeOff.csv",
    "struct": "qto/Structural_Schedule.csv",
    "mep":    "qto/MEP_TakeOff.csv",          # <-- new
}

LOCAL_QTO = {
    "arch":   os.path.join(LOCAL_BASE, "qto", "Architecture_TakeOff.csv"),
    "struct": os.path.join(LOCAL_BASE, "qto", "Structural_Schedule.csv"),
    "mep":    os.path.join(LOCAL_BASE, "qto", "MEP_TakeOff.csv"),          # <-- new
}
```

Then find the `fetch_qto_and_cost` and `main` functions and add your file to the merge call. Look for the `merge_takeoffs` call in `main()`:

```python
all_elements = merge_takeoffs(arch_rows, struct_rows)
```

Replace it with:

```python
all_elements = merge_takeoffs(arch_rows, struct_rows, mep_rows)
```

And update `merge_takeoffs` to accept the additional argument, or simply concatenate before passing:

```python
all_elements = merge_takeoffs([*arch_rows, *mep_rows], struct_rows)
```

> **Note:** `merge_takeoffs` deduplicates by `ElementId` — structural rows override architectural rows on overlap. If two files share ElementIds, the last file in the merge wins. Keep discipline schedules free of cross-discipline duplication where possible.

### Step 4 — Map the Assembly Codes in `cost_data.csv`

Every Assembly Code that appears in the new take-off must have at least one matching row in `cost_data.csv`, otherwise the element will be counted as **unmapped** (shown in red at the bottom of the Line Item Detail table).

Add rows for any new codes following the existing format:

```
Cluster Name,Assembly Code,Assembly Group Name,Description,Unit,Total O&P,Fixed Quantity
Services,D2020,Domestic Water,Hot water system,GSF,"$12,00",30000
```

- **Unit** determines how the quantity is extracted: `SF`/`GSF` → area, `LF` → length, `EA`/`Flight` → count, `CY` → volume÷27, `CF` → volume, `MSF` → area÷1000.
- Leave **Fixed Quantity** blank to use the takeoff quantity; fill it in to override.

### Step 5 — Run the script locally to verify

```bash
python tvd_analysis.py
```

The dashboard opens in your browser automatically. Check the **Line Item Detail** table — any elements from the new file without a matching Assembly Code will appear in the orange "Unmapped elements" row at the bottom. Click that row to download a CSV for review in Revit.

### Step 6 — Commit and push

```bash
git add qto/MEP_TakeOff.csv cost_data.csv tvd_analysis.py
git commit -m "Add MEP take-off and cost mapping"
git push
```

Pushing a change to `tvd_analysis.py`, `cost_data.csv`, or any file inside `qto/` triggers the GitHub Actions workflow, which re-runs the script and redeploys the dashboard to GitHub Pages.

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
| **Current** | Live view of the most recent run — cards, donut chart, TVD chart, line item table |
| **History** | Replay any saved snapshot with the full dashboard view |
| **Compare** | Side-by-side chart comparison of any two versions (including Current) |

History snapshots are saved automatically to `history/` (gitignored) each time the script runs locally for the first time. To save additional snapshots manually, call `save_snapshot(label, results, summary, unmapped_count)` from within the script.

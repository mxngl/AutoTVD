# AutoTVD Results — JSON Schema

Each run of `tvd_analysis.py` writes two files here:
- `YYYYMMDD_HHMMSS.json` — immutable timestamped snapshot
- `latest.json` — always overwritten with the most recent run

## Top-level keys

| Key | Type | Description |
|-----|------|-------------|
| `meta` | object | Run provenance |
| `financials` | object | Grand total, target, delta, $/SF, status |
| `cluster_targets` | object | `{ "Cluster Name": target_value, … }` |
| `cluster_summary` | array | Per-cluster rollup (see below) |
| `line_items` | object | `{ "Cluster Name": [ line_item, … ], … }` |

## `meta`

| Field | Description |
|-------|-------------|
| `generated_at` | ISO-8601 timestamp of this run |
| `date` | YYYY-MM-DD |
| `label` | Human-readable run label |
| `data_source` | Path or URL of QTO/cost files used |
| `gross_sf` | Building gross square footage (default 30,000) |
| `total_elements` | Combined arch + structural takeoff element count |
| `duplicates_removed` | Elements deduplicated across sources |
| `unmapped_count` | Elements with no Assembly Code |
| `dnc_count` | Elements marked DNC (Do Not Count) |

## `financials`

| Field | Description |
|-------|-------------|
| `grand_total` | Sum of all cluster estimates ($) |
| `tvd_target` | Total TVD budget target ($) |
| `delta` | `grand_total − tvd_target` (negative = under budget) |
| `delta_pct` | Delta as % of target |
| `cost_per_sf` | `grand_total / gross_sf` |
| `status` | `"under_target"` \| `"over_target"` \| `"on_target"` |

## `cluster_summary` items

```json
{
  "cluster":   "Shell",
  "estimate":  4430372.0,
  "target":    3826446.0,
  "delta":     603926.0,
  "delta_pct": 15.78,
  "per_sf":    147.68
}
```

## `line_items` rows

```json
{
  "ac":        "B1010",
  "group":     "Floor Construction",
  "desc":      "Structural Engineered Bamboo",
  "unit":      "SF",
  "unit_cost": 100.0,
  "qty":       29428.0,
  "qty_src":   "Area (SF)",
  "total":     2942800.0,
  "notes":     ""
}
```

`qty_src` values: `"Area (SF)"`, `"Volume (CY)"`, `"Length (LF)"`, `"Count (EA)"`,
`"Fixed"`, `"Mirror: <AC>"`, `"No takeoff match"`, `"Toilet elements (D2010)"`

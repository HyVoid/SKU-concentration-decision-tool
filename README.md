
<!-- Language Navigation -->
<div align="center">

[English](#english)

</div>

---

<a name="english"></a>

# Run SKU Concentration & Rationalization Analysis Before the Category Review

<div align="center">

![License](https://img.shields.io/badge/license-Apache%202.0-blue)
![Platform](https://img.shields.io/badge/platform-Excel%20365%20%2B%20Browser%20HTML-orange)
![Type](https://img.shields.io/badge/type-Retail%20Category%20Analytics-004B87)

**Stop guessing which SKUs carry the category. Get a repeatable, auditable concentration analysis that works on your live weekly data — no database, no BI team, no pivot table rebuilds.**

[▶ Live Demo](#) | [📥 Download Workbook](#) | [📝 Full Methodology](./docs/methodology.md)

</div>

---

<!-- Add real dashboard screenshot here -->
<div align="center">
  <img src="./docs/dashboard-hero.png" alt="SKU Concentration Dashboard" width="800">
  <br/>
  <em>Active SKUs: 124 · SKUs for 80% revenue: 17 · Concentration Ratio: 13.7%</em>
</div>

---

## What Decision Does This Help You Make?

Before you remove a single SKU from the assortment, you need answers that a normal pivot table can't give you in one click:

- **Which SKUs actually carry the category?** (not the ones that are just listed)
- **Which items are dormant but still holding inventory?** (dead stock tying up working capital)
- **Does the concentration profile change by price tier?** (a Budget hero can look like a Tail SKU in a mixed-band view)
- **How many SKUs drive 80% of sales?** (and is that number acceptable for your range?)

This workbook answers all four questions **in seconds, on any filter combination**, by applying a rigorous Pareto engine that properly isolates the active population, separates price bands, and classifies dead stock before cumulative thresholds.

---

## About The Builder

I build **lightweight, productized decision tools** for operational business teams — mostly in Excel, occasionally in browser simulations. The philosophy is simple: turn recurring analytical questions into repeatable workflows that don't require ERP projects, BI teams, or custom software.

This project is part of a broader portfolio that applies the same pattern to cash flow stress testing, inventory aging, crew compliance tracking, and retail margin waterfalls.

> **I don't sell advanced Excel. I sell the ability to make a better business decision without waiting for IT.**

[See all projects →](#)

---

## Who This Is For

| User | How They Use It |
|---|---|
| **Category managers** | Pressure-test SKU concentration logic before a range review or rationalization recommendation |
| **Buying teams** | Run price-band-isolated Pareto on weekly data without rebuilding pivot tables for each store |
| **Inventory planners** | Flag dead stock and long-tail SKUs with positive stock positions for clearance or supplier return |
| **Merchandising analysts** | Compare Core / Main / Tail distribution across stores, groups, and weeks with a single filter change |
| **Range planning teams** | Audit classification logic and denominator selection before committing to a delist |
| **Junior analysts** | Learn how availability filtering, price band assignment, and IFS priority interact in a real Pareto engine |
| **Commercial directors** | Review SKU health and rationalization outputs in a browser simulation without opening the workbook |

Best suited for teams running weekly SKU-level analysis where concentration logic needs to be **communicated, audited, or stress-tested** outside a protected planning file.

---

## Why Most SKU Rationalization Errors Aren't Judgment Errors

They happen because the **denominator is wrong**, the **price band is mixed**, or the **dead stock gets misclassified** — all before any human decision is made.

Here's the typical risk path inside a merchandise planning file:

```text
Pivot table opened on full SKU list
→ Range filtered by category
→ Dead stock included in denominator
→ Price bands mixed in single Pareto run
→ Classification applied without priority order
→ Core SKU count overstated, Tail understated
→ Rationalization list misses actual dead stock
→ Inventory commitment made on a contaminated concentration curve
```

This workbook breaks that path and replaces it with a **verification workflow**:

```text
Weekly data pasted into Data_Input
→ 4-filter mask applied: Store × Group × Week × Price Band
→ Active population isolated (Sell_Qty > 0 OR Stock_Qty > 0)
→ Sales sorted, cumulative % calculated via SCAN
→ Dead stock checked BEFORE cumulative thresholds
→ Core / Main / Tail / Dead classified in a single LET pass
→ KPIs, Pareto chart, and Rationalization Matrix refreshed instantly
```

No pivot rebuild. No manual formula edits. No "what if" guesswork.

---

## Three Traps That Catch Even Experienced Category Managers

### Trap 1 – Pareto run on the listed range, not the active population

80 listed SKUs. 25 have zero sales and zero stock in the analysis week. They contribute nothing, but they inflate the denominator.

| Measurement basis | SKUs generating 80% revenue | Concentration ratio |
|---|---|---|
| All listed (80 SKUs) | 10 | 10/80 = **12.5%** |
| Active only (55 SKUs) | 10 | 10/55 = **18.2%** |

The same 10 SKUs drive the same 80% of revenue. The perceived concentration differs by **45%** simply because of what's in the denominator. Every rationalization built on the listed range is optimizing against the wrong baseline.

**How the workbook fixes it:**
```text
is_available = (Sell_Qty > 0) + (Stock_Qty > 0)   → 1 if active, 0 if dormant
mask = (Store = SelStore) × (Group = SelGroup) × (Week = SelWeek) × (PriceBand = SelRange) × (is_available > 0)
```
Dormant SKUs are excluded before any cumulative calculation. They never enter the Pareto.

---

### Trap 2 – Price band contamination in a mixed-range Pareto

Budget ($0–50) and Mid-Range ($50–200) SKUs compete on revenue in a single Pareto. A volume-priced Budget hero can rank below the 80% threshold and get flagged as "Tail" — wrong for its price architecture.

**Real example:** Footwear, Week 202621, SKU-A003 ($49.99, Budget).

- Mixed-band Pareto (11 SKUs): revenue $1,399.72, rank 5th, cumulative 89.8% → **Tail**
- Budget-only Pareto (3 SKUs): same revenue, rank 1st, cumulative 73.3% → **Main**

Same SKU, same week, same revenue. Different classification. If your replenishment algorithm acts on Tail, you'll understock the most productive SKU in the Budget band.

**How the workbook fixes it:**
The price band is isolated **before** the Pareto runs, using a staircase INDEX/MATCH lookup (not XLOOKUP, which shifts SKUs into the wrong tier). Concentration is measured within each band's competitive set.

```text
// Correct: INDEX/MATCH type=1 on ascending Min_Price
Price_Band = INDEX(Band_Name, MATCH(Unit_Price, Min_Price_ascending, 1))
// $75 → largest Min_Price ≤ $75 is $50 → "Mid-Range (50-200)" ✓

// Wrong: XLOOKUP match_mode=1 returns first Min_Price ≥ Unit_Price
// $75 → first Min_Price ≥ $75 is $200 → "Premium (200-500)" ✗
```

---

### Trap 3 – Dead stock misclassified because IFS checks thresholds first

When `Sales_Amount = 0`, the item must be flagged as Dead **before** any cumulative percentage is evaluated. The wrong order creates a nasty edge case during blackout weeks (total category sales = 0).

```excel
// Wrong: cumulative thresholds checked first
=IFS(cum_pct <= 0.6, "Core", cum_pct <= 0.8, "Main", Sales_Amount = 0, "Dead", TRUE, "Tail")
// In a zero-sales week, cum_pct = 0 for every SKU → everything becomes "Core"

// Correct: dead stock check first
=IFS(Sales_Amount = 0, "Dead", cum_pct <= 0.6, "Core", cum_pct <= 0.8, "Main", TRUE, "Tail")
```

One blackout week included in a rolling review can misclassify the entire population as Core, hiding the real tail and dead stock. The workbook's classification order is **non-negotiable**: Dead → Core → Main → Tail.

---

## Example Scenario

**Store S001 · FOOTWEAR · Week 202621 · Mid-Range ($50–200)**

7 active SKUs, one with zero sales and positive stock (Dead).

| Rank | SKU | Unit Price | Sell Qty | Sales | Cumul. % | Class |
|---|---|---|---|---|---|---|
| 1 | SKU-A002 | $129.99 | 32 | $4,159.68 | 33.7% | Core |
| 2 | SKU-A001 |  $89.99 | 45 | $4,049.55 | 66.4% | Main |
| 3 | SKU-A004 | $199.99 | 15 | $2,999.85 | 90.7% | Tail |
| 4 | SKU-A007 |  $59.99 |  7 |   $419.93 | 94.1% | Tail |
| 5 | SKU-A008 |  $79.99 |  5 |   $399.95 | 97.3% | Tail |
| 6 | SKU-A009 | $109.99 |  3 |   $329.97 |100.0% | Tail |
| 7 | SKU-A011 | $149.99 |  0 |       $0 |    — | Dead |

| Metric | Mid-Range filter | All Ranges (mixed) |
|---|---|---|
| Active SKUs | 7 | 11 |
| SKUs for 80% | **3 (42.9%)** | **4 (36.4%)** |
| Core / Main / Tail / Dead | 1 / 1 / 4 / 1 | 2 / 1 / 7 / 1 |
| Concentration ratio | **42.9%** | **36.4%** |

Switching to All Ranges adds four Budget and Premium SKUs. The concentration ratio drops not because the Mid-Range band became less concentrated, but because the denominator grew. **Range comparisons must isolate price bands before measuring concentration.**

---

## What The Workbook Actually Delivers (Not How It Calculates)

- **Active population filter** — Four-dimensional boolean mask removes dormant SKUs before any calculation. Price band assigned via configurable boundary table.
- **Instant SKU classification** — Every active SKU gets exactly one tier: Dead, Core, Main, or Tail, based on cumulative revenue contribution.
- **Three live KPI tiles** — Total Active SKUs, SKUs to reach 80% revenue, and the resulting Concentration Ratio. All refresh on filter change.
- **Rationalization Matrix** — Filters the classification output to show only Dead and Tail SKUs, with recommended actions (Return/Liquidate, Markdown/Delist).
- **Dynamic dropdowns** — Store, Group, Week, and Price Band lists auto-populate from the data table. No named range maintenance needed.
- **Configurable thresholds** — Core/Main/Tail cutoffs (default 60%/80%/90%) editable on the Readme & Config sheet.

---

## Workbook Logic (For the Skeptical Reviewer)

The workbook contains four dedicated sheets, segregated by concern. No sheet does the job of another.

1. **Readme & Config** — Price band boundaries, concentration thresholds, reset procedure. Safe to edit.
2. **Data_Input** — Paste weekly SKU data into `tbl_RawData`. Seven fixed columns. Accepts 100k+ rows.
3. **Calc_Engine (hidden)** — Five formula cells run the entire engine. Hidden to prevent accidental edits.
4. **Dashboard** — Four dropdown filters, three KPI tiles, classification table, rationalization matrix, Pareto chart checklist.

All calculations are linked through **13 workbook-level named ranges**. Changing a filter on the Dashboard propagates through `rng_CalcArray` into every KPI tile without breaking a single cell reference.

---

## Implementation Notes (For Those Who Want to Look Under the Hood)

<details>
<summary>Click to expand: Classification engine pipeline & key formulas</summary>

The full calculation is a single LET formula in `Calc_Engine!A11`, executed in one forward pass with no iteration.

```text
Stage 1: FILTER  → apply 4D mask (Store × Group × Week × Price Band × is_available)
Stage 2: SORTBY  → reorder SKU IDs by Sales descending (parallel array)
Stage 3: SORT    → reorder Sales array descending
Stage 4: SCAN    → SCAN(0, sorted_sales, LAMBDA(a, c, a + c))
Stage 5: IF      → cumulative % = running_total / total_sales (guard against zero)
Stage 6: IFS     → Dead → Core → Main → Tail (fixed priority)
Stage 7: HSTACK  → assemble output array
```

**Classification outcome table:**

| Priority | Condition | Class | Threshold | Action signal |
|---|---|---|---|---|
| 1 | `Sales_Amount = 0` | Dead | — | Return / Liquidate |
| 2 | `cum_pct ≤ 0.60` | Core | 60% | Protect, prioritise |
| 3 | `cum_pct ≤ 0.80` | Main | 80% | Monitor |
| 4 | `TRUE` | Tail | >80% | Mark down / Delist |

**SKU count for 80% threshold:**

```excel
= IFERROR(MATCH(0.8, Cumulative_Pct_array, 1), 0) + 1
```

`MATCH(…,1)` on ascending array returns the largest value ≤ 0.8; +1 gives the first SKU pushing cumulative past 80%. Edge case for single dominant SKU or zero-sales blackout is handled.

**Named range architecture (key ranges only):**

| Named range | Purpose |
|---|---|
| `SelStore` / `SelGroup` / `SelRange` / `SelWeek` | Dashboard filter values |
| `rng_CalcArray` | Main LET formula spill anchor |
| `rng_KPI_TotalSKUs` / `rng_KPI_SKUs80` / `rng_KPI_Ratio80` | KPI tile formulas |
| `list_Stores` / `list_Groups` / `list_Weeks` / `list_PriceBands` | Dynamic dropdown sources |

All ranges are workbook-level. No cell addresses in formulas. Restructuring sheets won't break them.

</details>

---

## Get The Workbook

The complete Excel workbook includes the full formula engine, configurable tables, dynamic dropdowns, and a Pareto chart build checklist. Ready to paste your own weekly data.

[📥 Download Workbook](#)

For a quick interactive look without downloading:

[▶ Open Browser Simulation](#)

---

## Limitations

- **Pareto chart requires manual build** — A step-by-step checklist inside the workbook covers the combo chart setup (dual Y-axis, thresholds at 60%/80%/90%, brand colors).
- **Price band filter is exact-match** — Cross-band analysis requires two separate runs and manual comparison.
- **Single week per run** — Multi-week or period-to-date analysis needs upstream aggregation before pasting into Data_Input.
- **No side-by-side store comparison** — Two stores require two filter runs or a dashboard extension.
- **Excel 365 required** — Uses LET, FILTER, SORT, SORTBY, SCAN, UNIQUE, HSTACK. Older versions show `#NAME?`.
- **Table auto-extension** — Pasting beyond the existing table boundary needs a manual resize first.
- **Browser simulation rounding** — Sales rounded to two decimals; negligible impact on classification.

---

## Other Decision Tools You Might Need

- **Cash Flow Stress Testing** — Multi-scenario liquidity projection for SMEs
- **Flight Duty Compliance Tracker** — Crew fatigue rule monitoring for small aviation operators
- **Inventory Aging Analysis** — Aging buckets, slow-mover flagging, and provision calculation
- **Retail Margin Waterfall** — From ticket price to net margin, broken down by discount layer

[Explore all tools →](#)

---

## License

Distributed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0). Use it, adapt it, build on it — just keep the attribution.
```

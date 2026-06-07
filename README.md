
<!-- Language Navigation -->
<div align="center">

[English](#english)

</div>

---

<a name="english"></a>

# Run SKU Concentration and Rationalization Analysis Before the Category Review

<div align="center">

![License](https://img.shields.io/badge/license-Apache%202.0-blue)
![Platform](https://img.shields.io/badge/platform-Excel%20365%20%2B%20Browser%20HTML-orange)
![Type](https://img.shields.io/badge/type-Retail%20Category%20Analytics-004B87)

**Built for category managers and buying teams who need to pressure-test SKU concentration logic — Pareto thresholds, price band isolation, dead stock classification, and rationalization outputs — without rebuilding a pivot table for every filter combination.**

[Live Demo](#) | [Full Excel Workbook](#)

</div>

A self-contained Excel workbook and browser simulation that runs the complete Pareto classification engine for retail SKU concentration analysis — including multi-dimensional filtering, staircase price band lookup, dead stock separation, and live KPI derivation — across up to 100,000 rows of weekly transactional data with no database, no pivot tables, and no manual formula edits required between analyses.

---

## Quick Preview

<!-- Add screenshot here -->
> View the interactive simulation: [SKU Concentration Dashboard Live](#)

---

## Why This Exists

Most SKU rationalization errors are not caused by bad category judgment.

They happen because availability filtering, price band isolation, classification priority order, and concentration threshold logic are difficult to isolate and verify inside a merchandise planning file or a general-purpose pivot table.

The risk follows a predictable path:

```
Pivot table opened on full SKU list
-> Range filtered by category
-> Dead stock included in denominator
-> Price bands mixed in single Pareto run
-> Classification applied without priority order
-> Core SKU count overstated, Tail understated
-> Rationalization list misses the actual dead stock
-> Inventory commitment made on a contaminated concentration curve
```

This workbook is designed to change the verification workflow:

```
Weekly transactional data pasted into Data_Input
-> 4-filter mask applied: Store × Group × Week × Price Band
-> Active population isolated: Sell_Qty > 0 OR Stock_Qty > 0
-> Sales sorted descending, cumulative calculated via SCAN
-> Dead stock checked before cumulative thresholds
-> Core / Main / Tail / Dead classification computed in single LET pass
-> KPIs, Pareto chart, and Rationalization Matrix rendered on Dashboard
-> All outputs refresh on filter change — no pivot rebuild required
```

A rationalization recommendation that requires rebuilding a pivot table for each store, group, and week combination is a bottleneck.  
A workbook that reruns the full Pareto on every filter change — in one recalculation — converts that bottleneck into a shareable, auditable analysis artifact.

---

## Three SKU Concentration Traps That Catch Category Managers

### Trap 1 — Pareto run on the listed range instead of the available population

A category of 80 listed SKUs may contain 25 with zero sales and zero stock in any given analysis week. Those 25 items contribute zero revenue and zero available inventory, but they inflate the denominator of every concentration ratio calculated against the listed range.

The active population is 55 SKUs.

| Measurement basis | SKUs generating 80% of revenue | Concentration ratio |
|---|---|---|
| All listed (80 SKUs) | 10 SKUs | 10 / 80 = **12.5%** |
| Active only (55 SKUs) | 10 SKUs | 10 / 55 = **18.2%** |

The same 10 SKUs drive the same 80% of revenue. The concentration ratio differs by 45% depending on whether the denominator is correct.

A buying team measuring concentration on the listed range consistently underestimates how few SKUs are actually carrying the category. Every range rationalization built on a listed-range denominator is optimizing against the wrong baseline.

The availability filter in this workbook:

```
is_available = (Sell_Qty > 0) + (Stock_Qty > 0)   → 1 if active, 0 if dormant

mask = (Store_ID = SelStore)
     × (Group_ID = SelGroup)
     × (Week = SelWeek)
     × (Price_Band_Label = SelRange)
     × (is_available > 0)
```

Dormant SKUs are excluded before the FILTER runs. They do not enter the cumulative calculation. They do not inflate the denominator or suppress the concentration ratio.

### Trap 2 — Price band contamination in a mixed-range Pareto

Running a single Pareto across Budget ($0–50) and Mid-Range ($50–200) SKUs in the same category produces a concentration curve that is arithmetically correct and operationally misleading.

Volume-priced Budget SKUs compete against margin-priced Mid-Range SKUs on a revenue basis. The SKUs that dominate the Budget band may rank well below the 80% threshold in the mixed-band analysis — and receive a rationalization action signal that is wrong for their price architecture.

**Example:** Footwear category, Week 202621.

SKU-A003 ($49.99, Budget): generates $1,399.72 in weekly revenue.

- In the mixed-band Pareto (11 active SKUs, total revenue $16,268): ranks 5th, cumulative 89.8% → classified **Tail**.
- In a Budget-only Pareto (3 Budget SKUs, total $1,909): ranks 1st, cumulative 73.3% → classified **Main**.

Same SKU. Same week. Same revenue. Different classification depending on which population it competes within.

A replenishment algorithm acting on Tail classification will understock the Budget band's most productive SKU.  
A markdown trigger set at Tail status will clear inventory that is Main-equivalent within its own price architecture.

The staircase price band lookup that correctly assigns each SKU to its tier:

```
// Correct — INDEX/MATCH match_type=1, ascending Min_Price array
Price_Band_Label = INDEX(Band_Name, MATCH(Unit_Price, Min_Price_ascending, 1))
// MATCH type=1 returns position of largest Min_Price ≤ Unit_Price
// $75 → largest Min_Price ≤ $75 is $50 → "Mid-Range (50-200)" ✓

// Wrong — XLOOKUP match_mode=1 (returns next LARGER value)
Price_Band_Label = XLOOKUP(Unit_Price, Min_Price, Band_Name, "No Match", 1)
// match_mode=1 returns first Min_Price ≥ Unit_Price
// $75 → first Min_Price ≥ $75 is $200 → "Premium (200-500)" ✗
```

The XLOOKUP error moves every non-boundary SKU to a higher price tier. A $75 product classified as Premium receives wrong replenishment depth targets, wrong markdown timing, and wrong space allocation — for every week the classification runs on the incorrect formula.

### Trap 3 — Dead stock misclassified by IFS evaluation order

SKU classification requires one structural decision that is not obvious from the formula syntax: the zero-sales condition must be evaluated before cumulative percentage thresholds, not after.

The evaluation order determines whether dead stock appears in the rationalization matrix or disappears into a misassigned tier.

```
// Wrong order — cumulative thresholds checked before zero-sales test
= IFS(cum_pct <= 0.6,  "Core",
      cum_pct <= 0.8,  "Main",
      Sales_Amount = 0, "Dead",   ← unreachable if prior conditions fire
      TRUE,            "Tail")

// Correct order — zero-sales checked first regardless of cumulative position
= IFS(Sales_Amount = 0, "Dead",   ← captured before any threshold comparison
      cum_pct <= 0.6,   "Core",
      cum_pct <= 0.8,   "Main",
      TRUE,             "Tail")
```

The wrong order matters most in one specific scenario: when the entire filter selection returns zero total sales for the week — a blackout period caused by a store closure, system outage, or post-event clearance gap.

When `total_sales = 0`, the cumulative percentage formula returns 0 for every SKU. The wrong IFS order evaluates `cum_pct = 0 ≤ 60%` first and classifies every item in the population as Core.

An analyst running a rolling 4-week rationalization review who includes one blackout week will have that week's full population — including dead stock — incorrectly classified as Core. Any multi-week rank aggregation or average tier assignment that includes that period will suppress the true Tail and Dead count and overstate Core representation across the review window.

The correct order classifies all zero-sales SKUs as Dead regardless of total sales structure, and routes them directly into the rationalization matrix.

---

## Who This Tool Is For

| User | Practical Use Case |
|---|---|
| Category managers | Pressure-test concentration logic before a range review or assortment rationalization recommendation |
| Buying teams | Run price-band-isolated Pareto on weekly data without rebuilding pivot tables between stores |
| Inventory planners | Identify dead stock and long-tail SKUs with positive stock positions for clearance or supplier return |
| Merchandising analysts | Compare Core / Main / Tail distribution across stores, groups, and weeks on a single filter change |
| Range planning teams | Audit classification logic and denominator selection before submitting a rationalization commitment |
| Junior analysts | Learn how availability filtering, staircase price band lookup, and IFS priority interact in a Pareto engine |
| Commercial directors | Review SKU health and rationalization outputs in a browser simulation without distributing the workbook |

Best suited for teams running weekly SKU-level analysis where concentration logic needs to be communicated, audited, or stress-tested outside a merchandising system or protected planning file.

---

## What The Workbook Calculates

### Active Population Filter

Four-dimensional boolean mask: `(Store_ID = sel) × (Group_ID = sel) × (Week = sel) × (Price_Band_Label = sel) × (is_available > 0)`.  
Availability defined as `Sell_Qty > 0 OR Stock_Qty > 0`.  
Dormant SKUs excluded before any cumulative calculation runs.  
Price band assigned via staircase INDEX/MATCH type=1 lookup against the configurable Min_Price boundary table (`tbl_Config_PriceBands`).

### Sales Ranking and SCAN Accumulation

Filtered population sorted by Sales_Amount descending: SORTBY for SKU ID parallel array; SORT for the sales value array.  
Running cumulative computed via `SCAN(0, sorted_sales, LAMBDA(acc, curr, acc + curr))`.  
Cumulative percentage: `IF(total_sales = 0, 0, running_cumulative / total_sales)`.  
Produces a monotonically increasing percentage array from the highest-revenue SKU (first) to the lowest or zero-sales SKU (last).

### SKU Classification Engine

Single IFS pass over the cumulative array. Evaluation priority is fixed and non-negotiable:

1. **Dead** — `Sales_Amount = 0`. Evaluated first. Routes directly to the rationalization matrix.
2. **Core** — `cum_pct ≤ 60%`. First tier. Protect and prioritise.
3. **Main** — `cum_pct ≤ 80%`. Second tier. Monitor.
4. **Tail** — all remaining active SKUs with sales above zero. Mark down or delist.

Tier thresholds configurable in `tbl_Config_Thresholds` on the Readme & Config sheet.

### KPI Derivation

**Total Active SKUs** — `ROWS(spill_array) − IF(no-data sentinel detected, 1, 0)`. Excludes the placeholder row that fires when no SKUs match the filter selection.  
**SKUs for 80%** — `IFERROR(MATCH(0.8, Cumulative_Pct, type=1), 0) + 1`. Returns the minimum SKU count required to cross the 80% revenue threshold.  
**Concentration Ratio** — `[SKUs for 80%] ÷ Total Active SKUs`. Measures how few SKUs carry the category in the selected filter slice.

### Dynamic Dropdown Sources

Store, Group, and Week lists derived from `SORT(UNIQUE(FILTER(...)))` on the live data table.  
Price Band list derived from `tbl_Config_PriceBands[Price_Band_Name]`.  
Adding new stores, groups, or weeks to Data_Input automatically populates the filter dropdowns. No named range edits, no list maintenance.

### SKU Rationalization Matrix

FILTER pass on the computed classification array returning Dead and Tail SKUs only.  
Recommended action appended via IFS: `"⚠ Return / Liquidate"` for Dead; `"▼ Markdown / Delist"` for Tail.  
Output refreshes on every filter change. No separate query or pivot refresh required.

---

## Example Scenario

**Store S001 · FOOTWEAR · Week 202621 · Price Range: Mid-Range ($50–200)**

Seven active SKUs in the Mid-Range band. One with zero sales and positive stock (Dead).

| Rank | SKU | Unit price | Sell qty | Sales | Cumulative % | Class |
|---|---|---|---|---|---|---|
| 1 | SKU-A002 | $129.99 | 32 | $4,159.68 | 33.7% | Core |
| 2 | SKU-A001 |  $89.99 | 45 | $4,049.55 | 66.4% | Main |
| 3 | SKU-A004 | $199.99 | 15 | $2,999.85 | 90.7% | Tail |
| 4 | SKU-A007 |  $59.99 |  7 |   $419.93 | 94.1% | Tail |
| 5 | SKU-A008 |  $79.99 |  5 |   $399.95 | 97.3% | Tail |
| 6 | SKU-A009 | $109.99 |  3 |   $329.97 |100.0% | Tail |
| 7 | SKU-A011 | $149.99 |  0 |       $0 |    — | Dead |

| | Mid-Range filter | All Ranges (mixed band) |
|---|---|---|
| Active SKUs | 7 | 11 |
| SKUs for 80% | **3 (42.9%)** | **4 (36.4%)** |
| Core | 1 | 2 |
| Main | 1 | 1 |
| Tail | 4 | 7 |
| Dead | 1 | 1 |
| Concentration ratio | **42.9%** | **36.4%** |

Switching to All Ranges adds four Budget and Premium SKUs to the population. The concentration ratio drops from 42.9% to 36.4% — not because the Mid-Range band became less concentrated, but because the denominator increased by four SKUs contributing negligible revenue relative to their price tier.

A range review comparing store concentration ratios on an All Ranges basis will consistently understate concentration in the revenue-driving price tiers. The correct comparison isolates each band before measuring.

---

## Why SKU Concentration Works This Way

The concentration curve is constructed by sorting the active SKU population by revenue descending and computing a running cumulative as a percentage of total category revenue for the selected filter slice.

Three structural decisions determine whether the output is analytically valid.

**The denominator is the active population, not the listed range.**  
Including dormant SKUs in the denominator reports lower concentration than the revenue structure implies. The active filter removes this distortion before a single cumulative value is calculated. The concentration ratio then reflects how few of the working, available SKUs are actually carrying the category — which is the commercially useful measurement.

**The price band is isolated before the Pareto runs, not applied as a label after.**  
Concentration analysis across multiple price architectures produces a single curve that correctly reflects total category revenue and incorrectly positions individual SKUs relative to their own competitive set. Replenishment depth, markdown triggers, and space allocation decisions are made at the price band level. The analysis needs to match the decision unit. Running a mixed-band Pareto and then segmenting the output by band is not equivalent to running separate Paretos per band — the cumulative denominators are different.

**The classification priority checks dead stock status before cumulative thresholds.**  
A SKU with zero sales cannot meaningfully occupy a Core or Main tier. Checking thresholds first exposes the classification to edge cases in cumulative arithmetic — including the zero-total-sales blackout scenario — that override the zero-sales fact. Checking dead stock first ensures every item with positive inventory and no revenue appears in the rationalization matrix, regardless of the week's overall sales structure.

Each decision appears minor in isolation. Together, they determine whether a rationalization recommendation is based on a valid concentration analysis or a contaminated one.

---

## How The Classification Engine Works

The full calculation runs as a single LET formula anchored at `Calc_Engine!A11`. The formula completes in one forward pass through the filtered population with no iteration and no circular references.

### Six-Stage Pipeline

```
Stage 1: FILTER    — apply 4D mask (Store × Group × Week × Price_Band × is_available)
Stage 2: SORTBY    — reorder SKU IDs by Sales_Amount descending (parallel array sort)
Stage 3: SORT      — reorder Sales_Amount array descending (matching sort for SCAN input)
Stage 4: SCAN      — running cumulative: SCAN(0, sorted_sales, LAMBDA(a, c, a + c))
Stage 5: IF        — cumulative %: IF(total_sales = 0, 0, cumulative / total_sales)
Stage 6: IFS       — classify: Dead → Core → Main → Tail
Stage 7: HSTACK    — assemble: [SKU_ID | Sales_Amount | Cum_Sales | Cum_Pct | SKU_Class]
```

### Classification Outcome Table

| Priority | Condition checked | Class | Threshold | Action signal |
|---|---|---|---|---|
| 1 (first) | `Sales_Amount = 0` | Dead | — | Return / Liquidate |
| 2 | `cum_pct ≤ 0.60` | Core | 60% | Protect, prioritise |
| 3 | `cum_pct ≤ 0.80` | Main | 80% | Monitor |
| 4 (default) | `TRUE` | Tail | >80% | Mark down / Delist |

### KPI_SKUs_80 Formula Mechanics

The minimum SKU count required to cross the 80% revenue threshold is derived as:

```
= IFERROR(MATCH(0.8, Cumulative_Pct_array, match_type=1), 0) + 1
```

`MATCH` with `match_type=1` operates on the ascending cumulative percentage array and returns the 1-indexed position of the largest value ≤ 0.8. Adding 1 gives the first SKU that pushes cumulative past 80%.

| Scenario | Cumulative % array (first four) | MATCH result | + 1 | Interpretation |
|---|---|---|---|---|
| Standard 7-SKU set | [33.7%, 66.4%, 90.7%, ...] | 2 (position of 66.4%) | **3** | 3 SKUs to reach 80% |
| Single dominant SKU | [91.0%, 97.0%, 100%] | #N/A → 0 via IFERROR | **1** | 1 SKU alone exceeds 80% |
| Zero total sales (blackout) | [0%, 0%, 0%, ...] | 0 | **1** | Guarded by active SKU count check |

### Named Range Architecture

Thirteen workbook-level named ranges link the formula chain without cell address dependencies.

| Named range | Points to | Purpose |
|---|---|---|
| `SelStore` | `Dashboard!$B$5` | Active store filter value |
| `SelGroup` | `Dashboard!$B$6` | Active group filter value |
| `SelRange` | `Dashboard!$B$7` | Active price band filter value |
| `SelWeek` | `Dashboard!$B$8` | Active week filter value |
| `rng_CalcArray` | `Calc_Engine!$A$11` | Main LET formula spill anchor |
| `rng_CountSKUs` | `Calc_Engine!$G$11` | Active SKU counter (excludes no-data sentinel row) |
| `rng_KPI_TotalSKUs` | `Dashboard!$H$6` | KPI tile anchor — Total Active SKUs |
| `rng_KPI_SKUs80` | `Dashboard!$L$6` | KPI tile anchor — SKUs for 80% |
| `rng_KPI_Ratio80` | `Dashboard!$P$6` | KPI tile anchor — Concentration Ratio |
| `list_Stores` | `OFFSET(Calc_Engine!$H$11, ...)` | Dynamic dropdown source — stores |
| `list_Groups` | `OFFSET(Calc_Engine!$I$11, ...)` | Dynamic dropdown source — groups |
| `list_Weeks` | `OFFSET(Calc_Engine!$J$11, ...)` | Dynamic dropdown source — weeks |
| `list_PriceBands` | `OFFSET(Calc_Engine!$K$11, ...)` | Dynamic dropdown source — price bands |

Changing `SelStore` on the Dashboard propagates immediately through `rng_CalcArray`, through `rng_CountSKUs`, and into all three KPI tiles. No formula references a cell address. No formula breaks if the sheet is restructured.

---

## Workbook Logic

The workbook contains four dedicated sheets, segregated by concern. No sheet performs the function of another.

1. **Readme & Config** — Configure price band boundaries (four tiers, editable Min_Price values), set concentration thresholds (Core 60% / Main 80% / Tail 90%), and review the data reset procedure. No formula logic present. Safe to edit without affecting the calculation chain.

2. **Data_Input** — Paste weekly SKU transactional data into `tbl_RawData`. Seven required columns in fixed order: Week, Store_ID, Group_ID, SKU_ID, Unit_Price, Sell_Qty, Stock_Qty. Accepts up to 100,000 rows per diagnostic run. No calculated columns permitted. The structured table reference in Calc_Engine reads the live table regardless of row count.

3. **Calc_Engine (hidden)** — Five formula cells run the entire engine. `A11`: the main LET formula (all seven pipeline stages). `G11`: active SKU counter. `H11`, `I11`, `J11`, `K11`: SORT/UNIQUE dropdown source arrays for Store, Group, Week, and Price Band respectively. No manual edits required or permitted. Hidden to prevent accidental modification.

4. **Dashboard** — Four dropdown filter controls (Store, Group, Price Range, Week). Three KPI tiles with named range formulas. Tier distribution bar. Pareto chart placeholder with step-by-step manual build checklist. Full SKU classification table (all active SKUs, ranked). SKU Rationalization Matrix (Dead and Tail only, with recommended actions). All outputs refresh on filter change.

---

## Get The Workbook

The complete Excel workbook contains 13 workbook-level named ranges, four FAST-architecture sheets, a full LET/FILTER/SORT/SORTBY/SCAN/HSTACK formula engine, configurable price band and threshold tables, dynamic dropdown source lists, a Pareto chart build checklist, and a SKU Rationalization Matrix with recommended action outputs.

[Download the complete Excel workbook →](#)

---

## Limitations

- **Pareto chart requires manual build** — The combo chart (column bars + cumulative percentage line on a secondary axis) cannot be inserted programmatically; a step-by-step checklist in the workbook covers the full setup procedure including dual Y-axis configuration, brand color application, and reference line placement at the 60%, 80%, and 90% thresholds
- **Price band filter is exact-match by design** — Selecting a price band returns only SKUs assigned to that band; cross-band concentration analysis requires running the workbook twice with different filter selections and comparing outputs manually
- **Single week per run** — The engine analyzes one fiscal week per filter selection; rolling 4-week or period-to-date concentration analysis requires upstream aggregation of transactional data before pasting into Data_Input
- **No inter-store simultaneous comparison** — Two stores cannot be compared side-by-side in a single workbook session; comparison requires two separate filter runs or a dashboard extension not present in the base workbook
- **Excel 365 required** — LET, FILTER, SORT, SORTBY, SCAN, UNIQUE, and HSTACK are Microsoft 365 functions (Excel 2024 or later); the workbook will not recalculate in Excel 2019 or earlier; affected cells will display `#NAME?`
- **Table auto-extension requires paste discipline** — The pre-allocated table covers the sample data rows at initialization; pasting a dataset that exceeds the table boundary requires a manual table resize before pasting, or pasting within the table boundary to trigger Excel's auto-extension
- **No multi-tranche or hierarchical structure** — The workbook handles one Store × Group × Week × Price Band slice per run; multi-level category hierarchies, cluster aggregation, and portfolio-level concentration are outside scope
- **Browser simulation floating-point parity** — The browser simulation rounds Sales_Amount to two decimal places; minor differences from the Excel workbook on edge price inputs are expected and do not affect classification outcomes for standard inputs

---

## About This Project

This workbook is part of a broader effort to make category analysis logic visible, auditable, and independent of the original analyst's file.

Most retail SKU concentration models live in a single analyst's pivot table or a protected planning spreadsheet.  
The logic may be correct. The data may be current. But the analysis cannot be reviewed without the file, cannot be demonstrated without screen-sharing, and cannot be reproduced by anyone who did not build it.

The goal here is different: take the same calculation logic — active population filtering, staircase price band lookup, SCAN accumulation, IFS priority classification — and make it auditable in an open, self-contained workbook with a browser simulation that runs the same engine without a file upload.

Not as a replacement for merchandising systems. As a layer that makes the concentration analysis reviewable, presentable, and interrogable before a rationalization decision is committed.

If the work involves assortment planning, inventory rationalization, or category concentration analysis that currently lives in a protected file or requires the original builder to run it,  
[see what else is available →](#)

---

## License

Distributed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0).
```

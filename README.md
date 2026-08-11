# Flipkart Fulfillment & Seller Performance Analytics Platform

**An end-to-end seller performance and fulfillment analytics system for a multi-seller e-commerce marketplace — built in Excel, SQL, and Power BI, with one consistent data model and one consistent set of KPI definitions validated across all three.**

![Excel](https://img.shields.io/badge/Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![SQL](https://img.shields.io/badge/SQL-SQLite-4479A1?style=flat&logo=sqlite&logoColor=white)
![Power BI](https://img.shields.io/badge/Power%20BI-F2C811?style=flat&logo=powerbi&logoColor=black)
![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)

---

## Table of Contents
- [Overview](#overview)
- [The Business Problem](#the-business-problem)
- [Tech Stack](#tech-stack)
- [Repository Structure](#repository-structure)
- [Data Model (Star Schema)](#data-model-star-schema)
- [Dataset](#dataset)
- [KPI Dictionary](#kpi-dictionary)
- [The Seller Scorecard](#the-seller-scorecard)
- [Excel Workbook](#excel-workbook)
- [SQL Layer](#sql-layer)
- [Power BI Report](#power-bi-report)
- [Cross-Validation](#cross-validation-proving-all-three-tools-agree)
- [Key Results](#key-results)
- [How to Run This Project](#how-to-run-this-project)
- [Known Limitations](#known-limitations)
- [Roadmap](#roadmap--future-improvements)
- [What I Learned](#what-i-learned)
- [Author](#author)

---

## Overview

Flipkart's marketplace runs on 150+ third-party sellers, not Flipkart-owned inventory alone. The customer promise — order today, delivered on time — breaks at six points: seller dispatch, warehouse pick/pack, inventory availability, last-mile delivery, product quality, and address/payment issues. Each failure point has a different internal owner, and a single blended "SLA = 91%" number can't tell a manager which of the six is responsible in a given week, city, or seller cohort.

This project builds the **decision-support system** a real Supply Chain / Business Analyst would use to answer that question — not just a dashboard, but a full data model that lets you drill from a network-level SLA number down to the specific seller, warehouse, or SKU causing it.

It's built **three times, deliberately** — in Excel, SQL, and Power BI — using the same underlying data, the same schema, and the same KPI formulas, so the same business logic is provably consistent no matter which tool you're looking at it in.

---

## The Business Problem

| Failure point | Symptom the customer sees | Who owns it internally |
|---|---|---|
| Seller doesn't dispatch on time | "Shipped late" | Seller Management |
| Stock not available at the right warehouse | Long ETA / out of stock | Inventory Planning |
| Warehouse picks/packs slowly or wrongly | Wrong item, late dispatch | FC Operations |
| Last-mile rider overloaded | Delivery delay | Logistics / Hub Manager |
| Product quality or description mismatch | Return | Category + Seller Mgmt |
| Anything else | Exception ticket | Ops Excellence |

The system is designed so that any one of these six failure points is independently traceable — every stage of an order (order date, dispatch date, expected delivery, actual delivery) is captured separately, so a delay can be attributed to the correct owner instead of dumped into one blended "late" bucket.

---

## Tech Stack

| Layer | Tool | What it's used for |
|---|---|---|
| Data generation | Python (pandas, NumPy) | Synthetic dataset generation with realistic business logic |
| Spreadsheet analytics | Excel (openpyxl-built) | 14-sheet workbook, 134,000+ live formulas, named parameters, conditional formatting, charts |
| Database layer | SQLite | Schema, indexes, foreign keys, CTEs, views |
| BI layer | Power BI Desktop | Star-schema data model, DAX measures, interactive dashboard |
| Documentation | Markdown | This README + supplementary study guides |

---

## Repository Structure

```
flipkart-fulfillment-analytics/
│
├── excel/
│   └── Flipkart_Fulfillment_Seller_Performance_Analytics_Platform.xlsx
│
├── sql/
│   ├── schema.sql              # Table + index DDL
│   ├── queries.sql             # Seller scorecard, warehouse perf, SLA, inventory, dashboard KPIs, views
│   └── flipkart_analytics.db   # Populated SQLite database
│
├── powerbi/
│   ├── data/                   # CSV exports (10 tables) for import
│   ├── PowerBI_Build_Guide.md  # Full DAX + data-model build guide
│   └── Flipkart_Seller_Analysis.pbix   # (build locally — see Power BI section)
│
├── docs/
│   ├── Project_Study_Guide_and_Interview_Prep.md
│   ├── Project_Deep_Dive_Rationale_Build_Log_and_Roadmap.md
│   └── screenshots/
│
└── README.md
```

---

## Data Model (Star Schema)

Dimension tables (slow-changing) radiate out to fact tables (fast-growing transactional data). A derived analytics layer sits on top of the raw layer, and the dashboard/report sits on top of that — no raw data and no original calculations live in the presentation layer.

```
DIMENSIONS                    FACTS                          DERIVED
-----------                   -----                          -------
Sellers (150)     ─┐
Products (500)     ├─────►  Orders (15,000 lines) ────┐
Warehouses (4)     │         Inventory Txns (3,000)    ├──►  Inventory Summary (500 SKUs)
Riders (40)       ─┘         Deliveries (2,000)        ├──►  Seller Scorecard (150 sellers)
                              Returns (1,500)           ├──►  Warehouse Performance (4 WH)
                              Exceptions (800)          └──►  Delivery SLA (city/WH/month)
                                                                      │
                                                                      ▼
                                                            Dashboard / Report
```

**Grain matters.** `Orders` is at **order-LINE grain**, not order grain — one 3-item basket produces 3 rows. 15,000 order-line rows represent **8,362 distinct orders**, verified identically across all three tools (`SUMPRODUCT` in Excel, `COUNT(DISTINCT order_id)` in SQL, `DISTINCTCOUNT` in DAX). Line-level KPIs (units sold) use the raw row count; order-level KPIs (AOV, distinct order count) must dedupe.

---

## Dataset

Synthetic, generated with controlled business logic (not random noise) over a 6-month window:

| Entity | Volume |
|---|---|
| Sellers | 150 |
| Products (SKUs) | 500 |
| Warehouses (fulfillment centres) | 4 |
| Cities served | 8 |
| Delivery riders | 40 |
| Order lines | 15,000 (≈8,362 distinct orders) |
| Inventory transactions | 3,000 |
| Deliveries | 2,000 |
| Customer returns | 1,500 |
| Operational exception tickets | 800 |

Delivery expectations are built from a **zone-based rule** (intra-city = 2 days, intra-zone = 3 days, inter-zone = 5 days), with independent dispatch-delay and transit-delay simulation layered on top — so a late delivery can genuinely be traced to either a slow dispatch (seller/FC-side) or a slow transit (logistics-side).

> **Note on data realism:** Inbound and Outbound inventory transaction volumes were generated independently rather than demand-linked, which produces some overstocked and some understocked SKUs (occasionally net-negative stock) in the synthetic set. This is a known, disclosed limitation — see [Known Limitations](#known-limitations) — not a calculation error.

---

## KPI Dictionary

### Weighted into the Seller Scorecard

| KPI | Weight | Formula | Why this weight |
|---|---|---|---|
| On-Time Delivery Rate | 30% | On-time deliveries ÷ delivered orders | Highest-weighted — the failure a customer actually notices |
| Cancellation Rate | 20% (inverted) | Cancelled lines ÷ total lines | Pre-fulfillment failure |
| Return Rate | 20% (inverted) | Returns ÷ delivered orders | Post-fulfillment failure |
| Inventory Accuracy | 15% | Audited system-vs-physical match | Lowest of the "controllable" tier — largely warehouse-driven |
| Customer Rating | 15% | Seller's average rating (1–5) | The only subjective/holistic signal |

### Tracked separately (not weighted into the score)

| KPI | Formula | Why it's separate |
|---|---|---|
| On-Time Dispatch Rate | Dispatched within target ÷ total lines | Isolates seller/FC-caused delay from transit-caused delay |
| Inventory Availability (Stockout Rate) | SKUs below reorder point ÷ total SKUs | Distinct from "accuracy" — availability is seller-specific demand planning, not warehouse cycle-counting |

### Supporting KPIs (power the rest of the system)

SLA Adherence %, Avg Dispatch/Delivery Time, Warehouse Utilization %, Days of Inventory, Reorder Point, Safety Stock (√lead-time formula), Stockout Risk Flag, Dead Stock, Rider Utilization %, Perfect Order Rate (the compounding on-time × no-return × no-cancel check).

Full formulas, targets, and owners for all 17 KPIs are documented in the Excel workbook's README sheet and in `docs/Project_Study_Guide_and_Interview_Prep.md`.

---

## The Seller Scorecard

```
Seller_Score = OnTime_Delivery% × 30%
             + (1 − Cancellation%) × 20%
             + (1 − Return%) × 20%
             + Inventory_Accuracy × 15%
             + (Rating / 5) × 15%
```

Tiered into **Gold** (≥85) / **Silver** (≥70) / **Bronze** (≥55) / **Watchlist** (below), with all thresholds and weights stored as named parameters — not hardcoded — so a manager can re-run "what if we raised the SLA bar" scenarios by changing one cell/measure instead of hunting through formulas.

---

## Excel Workbook

**14 sheets, ~22,000 rows of data, 134,000+ formulas, zero recalculation errors.**

| Sheet | Purpose |
|---|---|
| README | Data model, 12 named parameters, KPI dictionary, assumptions |
| Sellers, Products, Warehouses, Riders | Dimension tables + computed performance columns |
| Inventory | Raw transaction log + derived per-SKU summary (safety stock, reorder point, DOI) |
| Orders, Deliveries, Customer Returns, Operational Exceptions | Fact tables |
| Seller Scorecard | All 150 sellers graded and tiered |
| Warehouse Performance | Utilization, accuracy, throughput per FC |
| Delivery SLA | SLA% by city / warehouse / month |
| Dashboard | 12 KPI cards + 9 charts |

Built with Excel Tables (structured references), named ranges as parameters, `SUMIFS`/`COUNTIFS`/`AVERAGEIFS`/`INDEX-MATCH` (deliberately avoiding `XLOOKUP`/`UNIQUE`/`FILTER` for cross-environment recalculation compatibility), and conditional formatting throughout.

---

## SQL Layer

A real, populated **SQLite database** (`sql/flipkart_analytics.db`) — not just a schema on paper.

- **`schema.sql`** — 9 tables, primary/foreign keys, indexes on every join/filter column
- **`queries.sql`** — 7 sections: Seller Scorecard (CTEs + `CASE`-based tiering), Warehouse Performance, Delivery SLA, Inventory Summary (same safety-stock formula as Excel), Dashboard KPIs, and two reusable **views** — `vw_seller_scorecard` and `vw_seller_availability`

```sql
-- Example: sellers who'd drop to Watchlist if the SLA bar were tightened
SELECT seller_id, seller_name, ontime_rate, seller_score
FROM vw_seller_scorecard
WHERE ontime_rate < 0.95 AND seller_score >= 55
ORDER BY ontime_rate ASC
LIMIT 10;
```

---

## Power BI Report

A star-schema data model imported from the same source data, with DAX measures and calculated columns replicating every Excel/SQL KPI, across 5 report pages:

1. **Executive Dashboard** — KPI cards, monthly trend, warehouse comparison, category revenue, ABC analysis, return reasons, with slicers for city/warehouse/category/month
2. **Seller Scorecard** — full 150-seller table with conditional formatting and a tier slicer
3. **Warehouse Performance**
4. **Delivery SLA**
5. **Inventory** — per-SKU stockout risk table

See `powerbi/PowerBI_Build_Guide.md` for the complete, copy-paste DAX and step-by-step build instructions (the `.pbix` itself is built locally in Power BI Desktop from the provided CSVs, since `.pbix` is a proprietary format that can't be generated outside the application).

---

## Cross-Validation: proving all three tools agree

Every headline number was checked across all three tools, not just computed once and assumed correct:

| Metric | Excel | SQL | Power BI |
|---|---|---|---|
| Distinct orders | 8,362 | 8,362 | 8,362 |
| Revenue (delivered) | ₹111,035,250.05 | ₹111,035,250.05 | ₹111.04M |
| Network SLA % | 49.5% | 49.5% | 49.5% |
| Cancellation rate | 4.12% | 4.12% | 4.12% |
| Network warehouse utilization | 82.4% | — | 82.4% |
| Seller SEL0001 — score / tier | 81.5 / Silver | 81.5 / Silver | (validated during build) |

---

## Key Results

- **Network SLA is 49.5%**, well below the 95% target — the system is built to diagnose *why*, not just report the shortfall
- **150 sellers graded and tiered**, surfacing which specific sellers need intervention instead of a single blended average
- **Warehouse Performance flags capacity status** (Overloaded / Healthy / Underutilized) per FC, so infrastructure bottlenecks are visible before they cause a network-wide SLA drop
- **Per-SKU stockout risk** identifies at-risk inventory before a customer-facing outage
- Delay is attributable to the correct stage (dispatch vs. transit) and the correct owner (seller vs. logistics vs. warehouse), rather than one undifferentiated "late" bucket

---

## How to Run This Project

**Excel:** Open `excel/Flipkart_Fulfillment_Seller_Performance_Analytics_Platform.xlsx`. Start with the README sheet.

**SQL:**
```bash
sqlite3 sql/flipkart_analytics.db < sql/queries.sql
# or open flipkart_analytics.db directly in DB Browser for SQLite
```

**Power BI:**
1. Unzip `powerbi/data/` (10 CSVs)
2. Open Power BI Desktop → Get Data → Text/CSV → import all 10
3. Follow `powerbi/PowerBI_Build_Guide.md` for relationships, DAX measures, and report pages

---

## Known Limitations

Documented deliberately, not hidden:

1. **On-time dispatch and stockout rate are tracked but not weighted into the Seller Score** — changing what feeds a live scoring formula changes what "85 = Gold" means for every seller retroactively; that's a decision for stakeholders, not a silent formula change.
2. **`par_SLATarget` / `par_ReturnMax` are reference targets**, not yet wired into conditional pass/fail flags.
3. **Top-10-sellers ranking can mis-attribute exact score ties** (`LARGE` + `MATCH` in Excel).
4. **Warehouse-level accuracy metrics are modeled as audited inputs**, not derived from the transaction log — realistic (real cycle-count data isn't reconstructable from order history alone), but means those figures are identical for every seller at a given warehouse.
5. **Inbound/Outbound inventory transactions were generated independently of order demand**, producing some overstocked and some net-negative-stock SKUs in the synthetic dataset — a data-generation limitation, not a calculation error.
6. **All data is synthetic.**

---

## Roadmap / Future Improvements

- Confidence-adjust the scorecard for order volume (a 10-order seller and a 1,000-order seller shouldn't carry equal statistical weight)
- Calibrate scorecard weights and tier cutoffs against a real outcome metric (repeat-purchase rate, complaint volume) instead of judgment-based weights
- Warehouse-adjust on-time delivery, separating seller-caused delay from infrastructure-caused delay
- Move from a static 6-month average to a rolling 30-day score
- Tie inventory transactions to actual order demand to eliminate the negative-stock artifact
- Add seasonality (festive-season demand spikes)
- Migrate the SQL layer to a client-server database (Postgres/MySQL) with dbt-orchestrated transforms at scale

---

## What I Learned

A clean recalculation with zero formula errors proves internal **consistency**, not **correctness** — real validation means spot-checking specific, named entities (a specific seller, a specific warehouse) against expectation every time something new is added. Three real bugs surfaced during the build (an Excel Table corruption issue from mismatched metadata, a recalculation-engine quirk with structured self-references, and a data-realism bug producing >100% warehouse utilization) and all three were caught this way, not by the absence of red error cells alone.

Full build log, every design decision and the alternative that was rejected, and 20+ "why not X instead" interview questions with model answers are in `docs/Project_Deep_Dive_Rationale_Build_Log_and_Roadmap.md`.

---

## Author

Built by [Your Name] — Production & Industrial Engineering, targeting Supply Chain Analyst / Business Analyst roles.

📧 [your email] · 🔗 [LinkedIn] · 💼 [Portfolio]

---

*If you're reviewing this for a hiring process and have questions about any specific formula, KPI choice, or design decision, everything is documented in `/docs` — including the honest limitations and what I'd change with more time.*

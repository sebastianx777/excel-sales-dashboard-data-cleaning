# Anvik Trail Co. — Sales Performance Dashboard

## Table of Contents
- [Tools](#tools)
- [What's here](#whats-here)
- [Key decisions](#key-decisions)
- [Stakeholder Questions](#stakeholder-questions)
- [Findings](#findings)
- [Recommendations](#recommendations)

An interactive Excel sales dashboard for a fictional outdoor gear retailer,
built entirely with formulas — no Power Query, no macros, no add-ins.
Covers data cleaning of a ~823-row messy export and filter-driven analysis
across region, category, channel, and time.

## Tools
Excel (formulas only — SUMIFS, COUNTIFS, INDEX/MATCH-family text cleaning,
data validation). No Power Query, VBA, or external libraries required to
open or use the file.

## What's here
- `Anvik_Trail_Co_Dashboard.xlsx` — the full workbook (4 sheets: README,
  Raw Data, Dashboard, Data Cleaning Log)
- `generate_raw_data.py` — script that generated the synthetic messy source
  data (seeded random, fully reproducible)
- `build_workbook.py` — script that builds the workbook from that data,
  including every cleaning and dashboard formula

## Key decisions
- **Cleaning done with formulas, not Power Query.** Power Query transformations
  live inside the workbook's internal XML in a way that isn't reliably
  reproducible outside Excel itself. Instead, every cleaning step — trimming
  whitespace, standardizing casing, parsing 3 different date formats, and
  stripping "$" from prices — is a plain formula in columns J:Q of Raw Data,
  so the logic is fully visible and auditable in the cells themselves.
- **Dropdown filters instead of native Slicers.** Slicers require a PivotTable
  bound to the Data Model, which has to be assembled inside Excel and doesn't
  travel reliably in a generated file. Data-validation dropdown cells wired to
  SUMIFS formulas give the same filter-and-recalculate experience and work
  identically in Excel, LibreOffice, and Google Sheets.
- **Duplicates flagged, not deleted.** A `COUNTIF` against an expanding range
  marks the *second and later* occurrence of each Order ID as a duplicate,
  keeping the first. The raw rows stay in the sheet for auditability; the
  duplicate flag simply excludes them from every downstream total.
- **Each breakdown table ignores its own filter dimension.** The Region
  breakdown responds to the Category/Channel/Date filters but not the Region
  filter itself — otherwise selecting "West" would leave only one bar on the
  Region chart. This mirrors how cross-filtering works in real BI tools.

## Stakeholder Questions
- How much of the raw export was actually unusable, and does it change the
  headline revenue number enough to matter?
- Which channel is really driving the most revenue once messy/duplicate rows
  are excluded — is Online outperforming In-Store, or does that only look
  true in the raw numbers?
- Is any one category so small that it's not worth a dedicated promotion,
  versus one that's underperforming *relative to* the others?

## Findings

**1. Roughly 13% of the raw export was unusable, but it didn't change which
channel leads — it did meaningfully shrink the sales total.** Of 823 raw
rows, 8 were fully blank, 35 were duplicate order entries, 32 had unparseable
dates, 20 had invalid (zero/negative) quantities, and 16 were missing a unit
price — leaving 713 clean, valid orders. Total revenue on the clean data is
**$142,039**, from **713 orders** and **1,462 units** — an average order
value of **$199.21**.

| Issue | Rows affected |
|---|---|
| Duplicate order rows | 35 |
| Unparseable dates | 32 |
| Invalid quantity (≤0) | 20 |
| Missing/blank price | 16 |
| Fully blank rows | 8 |
| **Clean, valid rows used** | **713 of 823** |

**2. Online is the top channel, but only by a narrow margin — not the
runaway leader raw totals might suggest.** Online: $49,797 · In-Store:
$47,786 · Wholesale: $44,456. The three channels are within about 12% of
each other, meaning channel strategy shouldn't be built around one clearly
dominant option.

**3. Camping Gear is the strongest category; Accessories is a clear
laggard, not just the smallest by nature.** Camping Gear leads at $39,637,
followed by Footwear ($32,760), Apparel ($29,070), and Climbing Gear
($27,925). Accessories trails at $12,648 — less than half of the next-lowest
category, a bigger gap than "just the smallest line" would suggest.

**4. Regional performance is close to even, with the South slightly ahead.**
South: $37,650 · Midwest: $36,182 · West: $35,597 · East: $32,611 — none of
the four regions is dramatically over- or under-performing the others.

## Recommendations
- **Don't over-trust "total revenue" figures pulled straight from a raw
  export.** 13% of this dataset was unusable, and if that fraction is
  consistent across future exports, any quick top-line number quoted without
  a cleaning pass will be inflated (from duplicates) and simultaneously
  incomplete (from dropped invalid rows) in ways that don't cancel out.
- **Treat channel strategy as a three-way balance, not an "Online-first"
  story.** With all three channels within 12% of each other, budget and
  staffing decisions shouldn't over-index on Online just because it happens
  to rank first this period.
- **Investigate Accessories before writing it off as inherently small.**
  A less-than-half-of-next-lowest gap is a bigger shortfall than category
  breadth alone would explain — worth checking whether it's an assortment
  issue (too few SKUs), a merchandising issue (buried on the site/in-store),
  or genuinely lower demand before deciding whether it's worth promotional
  investment.
- **Use the Region breakdown as a baseline, not an alarm.** With regions
  this close together, don't chase small period-to-period swings in any one
  region as if they were a trend — the current data doesn't show a region
  that's structurally underperforming.

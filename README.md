# Zepto Product & Inventory Analysis

An Excel-based exploratory data analysis project examining product
pricing, discount patterns, and recorded stock availability in a Zepto
product dataset. The project uses data-quality checks, Excel formulas,
PivotTables, PivotCharts, slicers, and a KPI dashboard to turn raw
product records into interpretable findings.

> **Project outcome:** The project improves visibility into the product
> dataset by summarizing pricing and stock patterns, comparing stockout
> rates across discount bands, and flagging data-quality issues that can
> affect price comparisons. It does not establish the causes of
> stockouts or measure actual sales performance because transaction,
> demand, and replenishment data were not available.

## 1. Business Context and Problem Statement

A product catalogue can contain many records with different listed
prices, discount values, weights, and stock statuses. Without structured
analysis, it can be difficult to answer basic business questions
consistently.

This project addresses the following analytical problems:

1.  **Limited visibility into pricing and discounts:** How are products
    distributed across price and discount ranges?
2.  **Limited visibility into recorded stock availability:** What share
    of products is marked out of stock, and does the observed rate vary
    across discount bands?
3.  **Difficult product-level comparisons:** Which products have the
    largest recorded discount amounts or the highest price per gram?
4.  **Data reliability concerns:** Are important fields missing or
    unsuitable for calculations, and do recorded unit prices agree with
    recalculated values?
5.  **Fragmented reporting:** Can these results be brought together in
    an Excel dashboard with KPIs, charts, and interactive filters?

### What problem does the project solve?

The project provides a structured, interactive view of product pricing,
recorded discounts, and stock status. It helps users identify patterns
and records that may deserve further investigation.

It does **not** solve stockouts operationally or prove that discounts
cause stockouts. Those decisions would require additional evidence, such
as sales volume, demand history, replenishment lead times, and inventory
movement over time.

## 2. Dataset and Preparation

-   **Records analysed:** 1,830 product records
-   **Tool:** Microsoft Excel
-   **Relevant fields:** Product Name, Original MRP, MRP, Discount %,
    Available Quantity, Discounted Selling Price, Selling Price, Weight
    (g), Price per gram, Out Of Stock, Discount Amount, Price Band, and
    Discount Band.
-   **Currency preparation:** Two price fields were converted from paise
    to rupees for analysis.
-   **Stock-status encoding:** `FALSE` = In Stock; `TRUE` = Out of
    Stock.

The source dataset does not contain a Category field, so the project
does not claim category-level analysis. The Low, Medium, and High price
bands are project-defined groupings based on Selling Price.

### Price-band definitions

  Band     Rule
  -------- --------------------------------------
  Low      Selling Price below ₹50
  Medium   Selling Price from ₹50 to below ₹200
  High     Selling Price of ₹200 or more

These thresholds are analytical choices for this project, not official
retail classifications.

## 3. Tools and Techniques

-   Microsoft Excel
-   Data cleaning and data-quality checks
-   `IF`, `IFERROR`, `COUNTIF`, `COUNTBLANK`, `INDEX`, `MATCH`, and
    `VLOOKUP`
-   Price and percentage calculations
-   PivotTables and PivotCharts
-   Slicers and interactive filtering
-   KPI summaries
-   Dashboard layout and visual formatting
-   Comparison of calculated values against source-recorded values

## 4. Analysis Performed: Business Questions, Method, and Findings

### Analysis 1 --- Product and stock-status counts

**Business question:** How many product records are present, and how are
they split between in-stock and out-of-stock status?

**Method:** Counted non-empty product names and used the Boolean
`Out Of Stock` field to count `TRUE` and `FALSE` records. Checked that
the stock-status counts reconcile with the total record count.

**Results:**

  Measure                   Records
  ----------------------- ---------
  Total product records       1,830
  In stock (`FALSE`)          1,612
  Out of stock (`TRUE`)         218

The counts reconcile: 1,612 + 218 = 1,830. This is a consistency check;
it does not independently prove that each source stock status is
correct.

**Business value:** Establishes the overall stock-status snapshot and
provides a baseline for more detailed comparisons.

### Analysis 2 --- Discount-band distribution

**Business question:** Are recorded discounts concentrated in particular
ranges?

**Method:** Grouped the source-recorded discount percentages into five
discount bands and counted product records in each band.

  Discount band     Product records      Share
  --------------- ----------------- ----------
  0--5%                         851     46.50%
  6--10%                        551     30.11%
  11--15%                       181      9.89%
  16--20%                       137      7.49%
  21% and above                 110      6.01%
  **Total**               **1,830**   **100%**

**Finding:** Approximately 76.6% of records fall in the 0--10% discount
bands, and the largest single band is 0--5%.

**Business value:** Helps describe the recorded discount mix. It does
not establish whether a discount is commercially effective, because
transaction and sales data are not available.

### Analysis 3 --- Stockout rate by discount band

**Business question:** Does the observed stockout rate differ across
discount bands?

**Method:** Built a PivotTable with Discount Band in Rows, Out Of Stock
status in Columns, and product count in Values. Displayed values as a
percentage of each row total and used the `TRUE` percentage as the
stockout rate.

  Discount band     Out-of-stock rate
  --------------- -------------------
  0--5%                        15.51%
  6--10%                        6.08%
  11--15%                       5.84%
  16--20%                       8.18%
  21% and above                10.53%

**Findings:** - The 0--5% band has the highest observed stockout rate at
15.51%. - The 11--15% band has the lowest observed rate at 5.84%. - The
rate rises again in the 16--20% and 21%+ bands, so the pattern is not
linear.

**Business interpretation:** The differences identify discount bands
that may merit further inventory investigation. They do not show that
discount levels cause stockouts. Product demand, replenishment, product
mix, and other factors could affect the pattern.

### Analysis 4 --- Price-band distribution

**Business question:** How is the product assortment distributed across
the project's price bands?

**Method:** Used the existing Price Band field, based on Selling Price
thresholds, and summarized the record share in each band.

  Price band                     Share of records
  ---------------------------- ------------------
  Low (below ₹50)                          23.66%
  Medium (₹50 to below ₹200)               55.96%
  High (₹200 or more)                      20.38%
  **Total**                              **100%**

**Finding:** The Medium band contains the largest share of product
records at 55.96%.

**Business value:** Provides a quick description of the listed price
mix. It does not measure sales contribution or customer preference.

### Analysis 5 --- Average selling price by price band

**Business question:** What is the average recorded selling price within
each price band?

**Method:** Summarized Selling Price by Price Band using the Average
aggregation in a PivotTable.

  Price band     Average selling price
  ------------ -----------------------
  Low                           ₹30.63
  Medium                       ₹109.14
  High                         ₹339.49

**Interpretation:** The averages are consistent with the band
definitions. Since Selling Price was also used to create the bands, this
analysis is descriptive and largely confirms the grouping logic rather
than providing an independent performance insight.

### Analysis 6 --- Average recorded discount by stock status

**Business question:** Does the average source-recorded discount differ
between products marked in stock and out of stock?

**Method:** Summarized Discount % by Out Of Stock status using Average
in a PivotTable.

  Stock status              Average recorded discount
  ----------------------- ---------------------------
  In stock (`FALSE`)                            8.10%
  Out of stock (`TRUE`)                         5.65%
  Overall average                               7.81%

**Finding:** The average recorded discount is 2.45 percentage points
higher for in-stock products than for out-of-stock products.

**Business interpretation:** This is a descriptive difference between
two groups. It does not prove that discounts improve availability or
that lower discounts cause stockouts. The source discount field also has
validation discrepancies that should be considered when interpreting
this comparison.

### Analysis 7 --- Top 10 products by price per gram

**Business question:** Which products have the highest recorded price
per gram?

**Method:** Used the existing Price per gram field to rank products and
display the top 10 in a PivotTable/chart.

**Data-quality consideration:** A zero product weight makes
price-per-gram calculation invalid. Two records were identified with
zero weight and `#DIV/0!` in the existing unit-price field. These
records were flagged rather than deleted, because other fields may still
be useful for other analyses.

**Business value:** Supports unit-price comparisons for products with
valid weights. Such comparisons should be made cautiously where products
differ in size, packaging, or product type.

### Analysis 8 --- Products ranked by discount amount

**Business question:** Which products have the largest recorded absolute
discount amounts?

**Method:** Summed Discount Amount by Product Name and sorted from
largest to smallest.

**Interpretation:** This ranks products by absolute discount amount, not
by discount percentage. A larger currency reduction does not necessarily
mean a higher percentage discount. The measure also does not represent
realized revenue loss or profit impact without sales quantities and
margin information.

## 5. Data Quality and Validation

Data-quality checks were included to assess whether important fields
were usable for the intended analysis.

### Checks performed

-   Counted product records using non-empty Product Name cells.
-   Checked for missing Product Names.
-   Checked zero MRP, zero product weight, and zero selling-price
    records.
-   Counted in-stock and out-of-stock records and reconciled the counts
    to the total.
-   Checked for missing Selling Price values.
-   Recalculated Price per gram for comparison with the existing field.
-   Compared source-recorded discount percentages against a discount
    recalculated from MRP and Selling Price.

### Price-per-gram validation

The validation found: - **1,828 records** matched the recalculated price
per gram to two decimal places. - **0 mismatches** among records
eligible for comparison. - **2 records** required investigation due to
invalid weight.

The two records were: - Maybelline New York Colossal Kajal Black -
Maybelline New York Colossal Kajal Super Black

Both had a recorded selling price, zero weight, and `#DIV/0!` for price
per gram.

**Decision:** The records were retained in the source dataset, but their
unit prices were not treated as valid comparisons. Their weights should
be verified from a reliable source before calculating a replacement
value.

### Discount validation

The source Discount % field contains fixed whole-number values.
Recalculated discounts did not always match the source-recorded
discounts. A comparison using whole-percentage rounding yielded 1,353
matches, 476 mismatches, and 1 record requiring investigation.

These mismatches were **not** treated as confirmed source-data errors.
The exact source calculation and rounding convention has not been
established, so the original discount values were retained unchanged.
The mismatch count is a validation limitation to document, not a target
to eliminate by changing formulas until more evidence is available.

## 6. Dashboard

The Excel dashboard brings key measures and comparisons into one view.
It includes:

-   Total Products
-   Average MRP
-   Average Selling Price
-   Average Discount
-   Top 10 Products by Savings
-   Product Distribution by Discount Band
-   Top 10 Products by Price per Gram
-   Product Distribution by Price Band
-   Stock Availability
-   Stockout Rate by Discount Band

The dashboard uses a Zepto-inspired purple and lavender palette with
green as an inventory-status accent. PivotTables, PivotCharts, and
slicers support interactive exploration.

## 7. Main Findings and Business Implications

1.  **The recorded discount mix is concentrated at lower discount
    levels.** About 76.6% of product records fall in the 0--10% bands.
2.  **Stockout rates vary across discount bands.** The highest observed
    rate is 15.51% in the 0--5% band, while the lowest is 5.84% in the
    11--15% band.
3.  **The listed assortment is concentrated in the Medium price band.**
    This band accounts for 55.96% of records under the project's
    thresholds.
4.  **Two zero-weight records cannot support valid price-per-gram
    comparisons.** They were flagged rather than automatically removed.
5.  **Source discount values do not always agree with price-based
    recalculations.** These differences remain documented for further
    investigation.

These findings help identify where further review may be worthwhile.
They should not be interpreted as evidence of sales performance,
profitability, customer behaviour, or causal effects on inventory.

## 8. Limitations

-   The dataset is product-level and does not establish actual sales
    volume, revenue, profit, or customer demand.
-   Stock status is a recorded snapshot; the data does not explain when
    or why stockouts occurred.
-   Discount-band and stock-status comparisons show association, not
    causation.
-   The exact source convention for calculating and rounding discounts
    has not been established.
-   Two zero-weight records cannot be used for reliable price-per-gram
    comparison until weights are verified.
-   Price-band thresholds are project-defined.
-   No category-level analysis was performed because the source dataset
    has no Category field.

## 9. Future Enhancements

-   Verify the source discount calculation convention and investigate
    the flagged differences using documented rules.
-   Verify zero weights from a reliable source before calculating
    replacement values.
-   Add sales volume, demand history, and replenishment information to
    investigate stockouts more directly.
-   Develop business recommendations once additional evidence supports
    them.
-   Improve chart-label readability and dashboard accessibility as the
    workbook evolves.

## 10. How to Use

1.  Clone or download this repository.
2.  Open the Excel workbook in Microsoft Excel.
3.  Review the `Full_Data`, `Data_Quality_Audit`, `Analysis`, and
    `Dashboard` sheets.
4.  Explore PivotTables and use available slicers to filter the
    analysis.
5.  Refresh PivotTables after any legitimate source-data update.

## Project Status

**Exploratory analysis and dashboard completed.** The project improves
visibility into product pricing, recorded discounts, and stock
availability. Validation limitations are documented, and deeper
operational recommendations are reserved for future work when more
evidence is available.

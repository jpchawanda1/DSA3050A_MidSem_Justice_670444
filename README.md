# DSA 3050A – Business Intelligence & Visualization
## Mid-Semester Practical Examination

| Field | Detail |
|-------|--------|
| **Student Name** | Justice Chawanda |
| **Student ID** | 670444 |
| **Date** | June 26–27, 2026 |
| **Tool** | Microsoft Power BI Desktop |
| **Source Dataset** | v1-legacy-carapi-datafeed.xlsx |

---

## Folder Structure

```
DSA3050A_MidSem_Justice_670444/
├── Dataset/
│   ├── v1-legacy-carapi-datafeed.xlsx        # Original source dataset (raw)
│   └── cleaned_carapi_dataset.xlsx           # Cleaned & transformed dataset (output)
├── PBIX/
│   └── MidSemExam.pbix                       # Power BI report file with all visuals
├── Screenshots/
│   ├── raw_dataset.png                       # Navigator – raw data preview
│   ├── power_query_editor.png                # Power Query Editor overview
│   ├── column_profile.png                    # Column profiling view
│   ├── applied_steps_overview.png            # All 25 applied steps panel
│   ├── step01_source.png                     # Step 01 – Source
│   ├── step02_navigation.png                 # Step 02 – Navigation
│   ├── step03_promoted_headers.png           # Step 03 – Promoted Headers
│   ├── step04_changed_data_types.png         # Step 04 – Changed Data Types
│   ├── step05_removed_columns.png            # Step 05 – Removed Unused Columns
│   ├── step06_renamed_columns.png            # Step 06 – Renamed Columns
│   ├── step07_removed_duplicates.png         # Step 07 – Removed Duplicates (Trim ID)
│   ├── step08_filtered_year.png              # Step 08 – Filtered Model Year ≥ 2018
│   ├── step09_filtered_msrp.png              # Step 09 – Filtered Price MSRP > 0
│   ├── step10_trimmed_text.png               # Step 10 – Trimmed Whitespace
│   ├── step11_full_vehicle_name.png          # Step 11 – Added Full Vehicle Name
│   ├── step12_market_segment.png             # Step 12 – Added Market Segment
│   ├── step13_dealer_margin.png              # Step 13 – Added Dealer Margin
│   ├── step14_drive_config.png               # Step 14 – Added Drive Config
│   ├── step15_date_parts.png                 # Step 15 – Added Date Parts
│   ├── step16_final_value_classification.png # Step 16 – Added Value Classification
│   ├── chart01_kpi_avg_msrp.png              # KPI Card – Average MSRP
│   ├── chart02_kpi_count_trimid.png          # KPI Card – Count of Trim ID
│   ├── chart03_kpi_avg_dealer_margin.png     # KPI Card – Average Dealer Margin
│   ├── chart04_bar_trimid_brand.png          # Clustered Bar – Count by Brand
│   ├── chart05_col_msrp_segment.png          # Clustered Column – Avg MSRP by Segment
│   ├── chart06_line_count_year.png           # Line Chart – Count by Model Year
│   ├── chart07_donut_count_segment.png       # Donut Chart – Count by Segment
│   ├── chart08_table_brand_metrics.png       # Table – Brand Financial Metrics
│   ├── chart09_matrix_brand_segment.png      # Matrix – Brand × Market Segment
│   ├── chart10_treemap_segment_count.png     # Treemap – Count by Segment
│   ├── chart11_area_avgmsrp_year.png         # Area Chart – Avg MSRP Trend by Year
│   ├── slicer01_brand.png                    # Slicer – Brand
│   ├── slicer02_market_segment.png           # Slicer – Market Segment
│   ├── slicer03_model_year.png               # Slicer – Model Year
│   ├── interactivity01_cross_filter_chevrolet.png  # Cross-filtering demo
│   ├── drilldown01_mode_enabled.png          # Drill-down mode enabled
│   ├── drilldown02_midrange_brands.png       # Drill-down into Mid-Range segment
│   ├── slicer_filter01_year2019.png          # Slicer filter – Year 2019
│   └── pbix_saved.png                        # PBIX file save confirmation
├── Insights.pdf                              # Business insights report (7 sections)
└── README.md                                 # This documentation file
```

---

## Q1: Power Query Transformations

### Dataset Overview – Before Transformation

The raw source file (`v1-legacy-carapi-datafeed.xlsx`) was loaded into Power BI Desktop via Power Query Editor.

| Property | Value |
|----------|-------|
| Source file | v1-legacy-carapi-datafeed.xlsx |
| Sheet loaded | v1-legacy-carapi-datafeed-sampl |
| Original records | ~10,800 rows |
| Original columns | 57 columns |
| Year range (raw) | 2015 – 2020 |
| Duplicate Trim IDs | Present |
| Missing/zero MSRP rows | Present |

### A. Data Cleaning Steps

All 25 transformation steps were applied inside the Power Query Editor:

| Step | Transformation | Description |
|------|---------------|-------------|
| 01 | Source | Connected to Excel workbook |
| 02 | Navigation | Selected sheet `v1-legacy-carapi-datafeed-sampl` |
| 03 | Promoted Headers | Used first row as column headers |
| 04 | Changed Data Types | Price MSRP → Decimal, Model Year → Integer, Invoice Price → Decimal |
| 05 | Removed Unused Columns | Dropped EV-specific sparse columns (battery, range, charge time) |
| 06 | Renamed Columns | Make Id → Manufacturer ID, Make Name → Brand, Model Name → Vehicle Model, etc. |
| 07 | Removed Duplicates | Removed duplicate rows based on Trim ID (kept first occurrence) |
| 08 | Filtered – Model Year | Kept only rows where Model Year ≥ 2018 |
| 09 | Filtered – Price MSRP | Kept only rows where Price MSRP > 0 |
| 10 | Trimmed Text | Removed leading/trailing whitespace from Brand, Vehicle Model, Engine Type |

### B. Calculated Columns Added

| Step | Column Name | Formula / Logic |
|------|-------------|----------------|
| 11 | Full Vehicle Name | `Brand & " " & Vehicle Model` |
| 12 | Market Segment | MSRP-based: Economy (<$20K) / Entry-Level (<$35K) / Mid-Range (<$60K) / Luxury (<$100K) / Ultra-Luxury (≥$100K) |
| 13 | Dealer Margin | `Price MSRP − Invoice Price` |
| 14 | Drive Config | Derived from Drive Type: AWD / FWD / RWD / Other |
| 15 | Created Year / Month / Quarter / Day | Date part extractions from Model Year |
| 16 | Value Classification | Segment-based: Budget / Mid / Premium / Luxury |

Also added: **Body Style**, **Specs Detail** (Horsepower + Engine Type), **Row Index**.

### C. Dataset After Power Query Transformation

| Property | Value |
|----------|-------|
| Records after transformation | **8,939 rows** |
| Columns after transformation | **63 columns** (original 57 + 9 calculated − 3 removed) |
| Transformation steps applied | **25 steps** |
| Records removed (duplicates) | ~200 |
| Records removed (Year < 2018) | ~1,400 |
| Records removed (MSRP ≤ 0) | ~200 |
| Model years in final dataset | 2018, 2019, 2020 |
| Unique brands | 31 |
| Market segments | 5 (Economy, Entry-Level, Mid-Range, Luxury, Ultra-Luxury) |

The cleaned and exported dataset (`cleaned_carapi_dataset.xlsx`) contains **8,939 rows × 57 columns** after additionally removing empty/sparse columns and rows with missing values in key fields.

---

## Q2: Dashboard Visuals

The Power BI report (`MidSemExam.pbix`, Page 1) contains **11 visuals** arranged on a single dashboard page:

| # | Visual Type | Title / Measure | Fields Used |
|---|-------------|-----------------|-------------|
| 1 | KPI Card (New) | Average of Price MSRP | Price MSRP (Average) |
| 2 | KPI Card (New) | Count of Trim ID | Trim ID (Count) |
| 3 | KPI Card (New) | Average of Dealer Margin | Dealer Margin (Average) |
| 4 | Clustered Bar Chart | Count of Trim ID by Brand | Y-axis: Brand, X-axis: Count of Trim ID |
| 5 | Clustered Column Chart | Average Price MSRP by Market Segment | X-axis: Market Segment, Y-axis: Avg MSRP |
| 6 | Line Chart | Count of Trim ID by Model Year | X-axis: Model Year, Y-axis: Count of Trim ID |
| 7 | Donut Chart | Count of Trim ID by Market Segment | Legend: Market Segment, Values: Count of Trim ID |
| 8 | Table | Brand Financial Metrics | Brand, Dealer Margin, Market Segment, Model Year, Price MSRP |
| 9 | Matrix | Average MSRP by Brand × Segment | Rows: Brand, Columns: Market Segment, Values: Avg MSRP |
| 10 | Treemap | Count of Trim ID by Market Segment | Category: Market Segment, Values: Count of Trim ID |
| 11 | Area Chart | Average Price MSRP by Model Year | X-axis: Model Year, Y-axis: Avg Price MSRP |

---

## Q3: Interactivity

### Slicers (3 added)

| Slicer | Field | Type | Values |
|--------|-------|------|--------|
| Slicer 1 | Brand | List (multi-select) | Acura, Alfa Romeo, Aston Martin, Audi, BMW … (31 brands) |
| Slicer 2 | Market Segment | List (multi-select) | Economy, Entry-Level, Luxury, Mid-Range, Ultra-Luxury |
| Slicer 3 | Model Year | List (single-select) | 2018, 2019, 2020 |

All three slicers filter every visual on the page simultaneously.

### Drill-Down

The **Clustered Bar Chart** (Count of Trim ID by Brand) uses a two-level hierarchy:

```
Market Segment  →  Brand
```

- Drill-down mode enabled via the drill-down icon on the visual header
- Clicking any segment bar (e.g., "Mid-Range") drills down to show individual brands within that segment
- Clicking the up-arrow returns to the segment level
- Example: Mid-Range drill-down shows Ford, Chevrolet, Ram, GMC, Toyota, Nissan among top brands

### Cross-Filtering

Power BI default cross-filtering is active across all visuals on the page:

- Selecting any data point on any visual filters all other visuals simultaneously
- Example demonstrated: Clicking "Chevrolet" in the bar chart filtered the KPI cards, donut chart, treemap, table, matrix, line chart, and area chart to show only Chevrolet data
- No manual cross-filter configuration was needed — Power BI applies this automatically

---

## Key Insights

Full analysis is documented in `Insights.pdf`. Summary findings:

| Insight | Detail |
|---------|--------|
| Top brand by volume | **Ford** (highest trim count), followed by Chevrolet and Toyota |
| Highest MSRP segment | **Ultra-Luxury** (~$200K average MSRP) |
| Most common segments | Mid-Range and Economy together represent >60% of all trims |
| MSRP trend | Average MSRP trended **upward** from 2018 to 2020 (~$49K → ~$51K) |
| Volume trend | Trim count **declined** year-over-year from 2018 to 2020 |
| Dealer margin leader | **Ultra-Luxury** segment has the highest average dealer margin (~$9K+) |
| Drive type distribution | AWD trims dominate the Luxury and Ultra-Luxury segments |

# Part 1: Business Data Cleaning, Validation & Excel Reporting

## Problem Summary
This project involves cleaning, validating, and analyzing a raw retail order dataset exported from multiple internal systems. The dataset contained 
numerous data quality issues including inconsistent text formatting, mixed date formats, duplicate records, missing values, invalid discounts, and 
order status inconsistencies. The goal was to produce a clean, validated, analysis-ready dataset along with summary reports for business review.

## Dataset Description
- **File:** raw_orders.xlsx
- **Total raw rows:** 932
- **Total columns:** 21
- **Fields include:** order_id, order_date, ship_date, customer_id, customer_name, segment, region, state, city, category, sub_category, product_name, ship_mode, quantity, unit_price, discount, sales, cost, profit, payment_status, order_status
- **Final cleaned rows:** 912 (after removing 20 exact duplicate rows)

## Tools Used
- Microsoft Excel
  - TRIM(), PROPER(), IF(), COUNTIF(), ABS(), IFERROR(), MONTH(), YEAR(), ROUND(), Duration.Days()
  - Power Query (Data → From Table/Range) for mixed date format handling
  - PivotTables for summary reporting
  - Remove Duplicates feature
  - Helper column method (formula → Paste Special Values Only → delete helper)

## Cleaning Steps Performed
1. **Task 2 - Text Cleaning:** Applied TRIM and PROPER functions to all text fields to remove extra spaces, fix case inconsistencies, and standardize values. Missing region and ship_mode values filled with Unknown.
2. **Task 3 - Date Cleaning:** Used Power Query with format-aware locale logic to convert 4 mixed date formats into a consistent DD-MM-YYYY format. Created shipping_delay_days column. Flagged 21 rows where ship date was before order date.
3. **Task 4 - Duplicates:** Removed 20 exact duplicate rows using Excel's Remove Duplicates. Flagged 24 rows with conflicting duplicate order_ids for manual review.
4. **Task 5 - Business Rules:** Filled 18 missing discounts with 0. Flagged 
15 negative and 15 outlier discounts. Flagged 145 cancelled orders, 37 failed payments, and 34 refunded orders.
5. **Task 6 - Calculated Columns:** Created cleaned_discount, 
calculated_sales, calculated_profit, profit_margin, order_month, order_year, and data_quality_flag columns.

## Business Rules Applied
- Missing region → Filled with Unknown
- Missing ship_mode → Filled with Unknown
- Missing discount → Filled with 0 (if other sales fields valid)
- Negative discount → Flagged as invalid; corrected using absolute value
- Discount above 0.5 (50%) → Flagged as invalid outlier; capped at 0.5
- Cancelled orders → Excluded from completed sales summary
- Failed payments → Excluded from completed sales summary
- Refunded orders → Separately summarized
- Ship date before order date → Flagged as invalid shipping record

## Summary of Data Quality Issues Found
| Issue | Count |
|-------|-------|
| Exact duplicate rows removed | 20 |
| Conflicting duplicate order_ids flagged | 24 |
| Missing region values | 26 |
| Missing ship_mode values | 22 |
| Missing discount values | 18 |
| Negative discount values | 15 |
| Discount outlier values (above 50%) | 15 |
| Ship date before order date | 21 |
| Cancelled orders | 145 |
| Failed payments | 37 |
| Refunded orders | 34 |
| Sales calculation mismatches | 93 |
| Profit calculation mismatches | 94 |
| **Final Clean rows** | **643** |
| **Final Warning rows** | **218** |
| **Final Invalid rows** | **51** |

## Summary of Final Pivot Reports
All pivot reports are in outputs/pivot_summary.xlsx:
1. **Sales and Profit by Region** — South had highest sales (2,266,376.24), Unknown region had lowest (292,047.64)
2. **Sales and Profit by Category and Sub-category** — Technology and Furniture categories led in total sales
3. **Order Count by Ship Mode** — Standard Class was most used (242 orders), Same Day least used (204 orders)
4. **Average Profit Margin by Segment** — Small Business had highest average profit margin (0.30), others at 0.29
5. **Refunded/Cancelled/Failed Orders by Region** — North had highest cancelled order count (48)
6. **Monthly Sales Trend** — 2024 total sales (4,794,539.69) higher than 2025 partial year (4,216,747.87)

## Key Business Insights
- South region generated the highest sales but East region had the highest profit, suggesting better cost management in the East
- Standard Class is the most preferred shipping mode, while Same Day delivery had surprisingly similar order counts to other modes
- Small Business segment showed slightly higher profit margins than other 
segments, making it a potentially more profitable customer type
- 2024 data is complete (12 months) while 2025 only covers 10 months, so year-over-year comparison should account for this
- 30 discount-related issues (15 negative + 15 outliers) suggest a 
systematic data entry problem in the discount field worth investigating
- North region had the highest cancellation rate (48 cancelled out of 214 total orders = 22.4%), which may warrant further investigation

## Assumptions and Limitations
- All slash-format (/) dates assumed to be in US format (MM/DD/YYYY) based on evidence from unambiguous values in the dataset
- Discount threshold of 0.5 chosen based on data distribution and real-world retail norms, since no fixed threshold was provided
- Negative discounts assumed to be sign-entry errors and corrected using 
absolute value
- Missing region/ship_mode filled with Unknown per business rules, which may slightly affect regional/shipping analysis
- 24 conflicting duplicate order_ids could not be resolved without access to original source system — retained and flagged for review
- 2025 data is incomplete (only 10 months available), affecting 
year-over-year comparisons

## Screenshots Included
| File | Shows |
|------|-------|
| screenshots/raw_data_preview.png | Raw dataset before any cleaning |
| screenshots/cleaned_data_preview.png | Cleaned dataset with all calculated columns |
| screenshots/pivot_summary_1.png | Sales and Profit by Region (sorted) |
| screenshots/pivot_summary_2.png | Refunded/Cancelled/Failed Orders by Region |

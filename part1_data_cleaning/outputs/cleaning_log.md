# Cleaning Log — Part 1: Business Data Cleaning


## Task 2: Clean Text Fields

### Issues Found
- customer_name: Trailing spaces, leading spaces, double spaces between words
- segment: Leading/trailing spaces, double spaces, wrong case (e.g. SMALL BUSINESS, consumer)
- region: Leading/trailing spaces, 2 blank values found
- state: No issues found
- city: No issues found
- category: Leading/trailing spaces, double spaces, wrong case (e.g. OFFICE SUPPLIES, FURNITURE)
- sub_category: Leading/trailing spaces (e.g. Accessories, Machines)
- ship_mode: Wrong case (STANDARD CLASS, FIRST CLASS), double spaces, leading/trailing spaces, 2 blank values found
- payment_status: Trailing spaces, wrong case (PENDING, failed)
- order_status: Leading/trailing spaces (Completed, Cancelled)

### Cleaning Actions Performed
- Applied TRIM and PROPER functions to all text fields
- Used helper column method: formula applied → Paste Special Values Only → helper column deleted

### Business Rules Applied
- Missing region values filled with Unknown
- Missing ship_mode values filled with Unknown

### Assumptions Made
- state and city had no detected issues but were cleaned for consistency
- All text fields standardized to Title Case


## Task 3: Clean and Validate Dates

### Issues Found
- Both order_date and ship_date had 4 mixed date formats:
  - DD Mon YYYY (e.g. 21 Jul 2024)
  - MM/DD/YYYY (e.g. 08/31/2024) — confirmed US format based on unambiguous values where the second number exceeds 12
  - DD-MM-YYYY (e.g. 28-11-2024) — confirmed day-first format based on unambiguous values where the first number exceeds 12
  - YYYY-MM-DD (e.g. 2024-05-24)
- Initial attempt using Power Query's en-IN locale first, en-US fallback produced incorrect results for ambiguous slash dates (where both day and month values were 12 or below), since it silently defaulted to Indian interpretation without verifying which locale the dataset actually used. This was caught by checking unambiguous dates in the same column, which 
proved all slash-format dates were actually US format.
- Negative shipping_delay_days found after correction (small values, -4 to -1 range) — ship date before order date in 21 rows

### Cleaning Actions Performed
- Used Power Query (Data → From Table/Range) to handle mixed date formats
- Built custom columns using format-aware parsing logic instead of a single locale guess:
  - Slash-format dates (/) parsed using en-US locale consistently
  - Dash-format dates (-) manually parsed as day-first using Text.Start/Text.Middle/Text.End extraction
  - DD Mon YYYY and YYYY-MM-DD formats parsed directly since they are unambiguous regardless of locale
- Verified correctness by spot-checking multiple unambiguous raw values against their parsed output before trusting the results
- Both order_date and ship_date converted to consistent Date format (DD-MM-YYYY display)
- Created new column shipping_delay_days = ship_date minus order_date using Duration.Days()

### Business Rules Applied
- Ship dates earlier than order dates flagged as invalid shipping records using shipping_validity_flag column

### Assumptions Made
- All slash-format (/) dates assumed to follow US format (MM/DD/YYYY) consistently, based on evidence from unambiguous values in the same column
- All dash-format (-) dates assumed to follow day-first format (DD-MM-YYYY), based on evidence from unambiguous values in the same column

### Records Flagged
- 21 rows flagged as "Invalid - Ship Before Order" in shipping_validity_flag column

### Tools Used
- Power Query (built into Microsoft Excel)
- M Language functions: Date.FromText(), Text.Contains(), Text.Start(), Text.Middle(), Text.End(), Duration.Days()


## Task 4: Handle Duplicates

### Issues Found
- 20 exact duplicate rows found (every column identical)
- 24 rows with duplicate order_id but conflicting information found after exact duplicates were removed

### Cleaning Actions Performed
- Used Excel Remove Duplicates (Data → Remove Duplicates) with all columns selected to identify and remove exact duplicate rows
- 20 exact duplicate rows removed
- 912 unique rows remain after removal

### Business Rules Applied
- Exact duplicate rows removed as they add no analytical value
- Conflicting duplicate order_id rows NOT deleted — flagged using duplicate_flag column in cleaned_orders.xlsx

### Records Removed
- 20 exact duplicate rows removed

### Records Flagged
- 24 rows flagged as "Duplicate order_id - Review" in duplicate_flag column in cleaned_orders.xlsx

### Limitations
- For the 24 duplicate order_id rows, both records had different data and it was not possible to determine which entry was correct at this stage
- These rows were kept and flagged for review


## Task 5: Apply Business Rules

### Issues Found
- 18 missing discount values
- 15 negative discount values
- Discount values of 0.55, 0.65, 0.7, 0.85 identified as outliers, separated from the normal cluster (0 to 0.25) by a clear gap in the distribution
- 145 cancelled orders, 37 failed payments, 34 refunded orders found
- 21 rows with ship_date earlier than order_date (already flagged in Task 3)

### Cleaning Actions Performed
- Filled 18 missing discount values with 0
- Created discount_flag column: flags negative discounts and discounts above 0.5 as invalid high discount
- Created business_rule_flag column: flags cancelled orders and failed payments as excluded, refunded orders as separate

### Business Rules Applied
- Discount threshold of 0.5 selected as the cutoff for invalid/outlier discounts, since the business rules did not specify an exact number. This was based on real-world retail norms (discounts above 50% are uncommon) and also aligned with a natural gap in this dataset's own discount distribution (0 to 0.25 normal cluster, 0.55 to 0.85 outlier cluster)
- Cancelled orders and failed payments excluded from completed sales via flag, not deletion
- Refunded orders flagged for separate summarization in Task 8

### Assumptions Made
- 0.5 chosen as discount threshold in absence of a specified business rule

### Records Flagged
- 15 negative discount rows
- 15 discount outlier rows
- 145 cancelled order rows
- 37 failed payment rows
- 34 refunded order rows


## Task 6: Create Calculated Columns

### Columns Created
- cleaned_discount: corrects discount_flag issues — negative discounts converted to absolute value (assumed sign entry error), discounts above 0.5 capped at 0.5 (assumed outlier/data entry error)
- calculated_sales: quantity x unit_price x (1 - cleaned_discount)
- calculated_profit: calculated_sales - cost
- profit_margin: profit / sales, rounded to 2 decimal places
- calculated_profit_margin: calculated_profit / calculated_sales (using calculated values) added as a bonus column to compare against the raw profit_margin and observe the impact of data cleaning
- order_month: month extracted from order_date using MONTH()
- order_year: year extracted from order_date using YEAR()
- data_quality_flag: combines duplicate_flag, discount_flag, business_rule_flag, and shipping_validity_flag into one overall status per row (Clean, Warning, or Invalid)

### Business Rules Applied
- A row is marked Invalid if discount_flag or shipping_validity_flag contains "Invalid"
- A row is marked Warning if duplicate_flag contains "Duplicate" or business_rule_flag contains "Excluded" or "Separate"
- All other rows marked Clean

### Assumptions Made
- Negative discounts assumed to be sign entry errors and corrected using absolute value rather than treating them as missing/unknown values
- Discounts above 0.5 capped at 0.5 rather than replaced with mean/median, since the true intended value could not be determined and capping at the valid boundary was considered safer than guessing
- IFERROR used in profit_margin calculations to handle any case where calculated_sales is 0, defaulting to 0 instead of showing a division error

### Records Flagged (data_quality_flag summary)
- Clean: 643 rows
- Warning: 218 rows
- Invalid: 51 rows


## Task 8: Create Pivot Summary Report

### Pivots Created
- Sales and profit by region (pivot_sales_profit_region) — sorted descending by calculated_sales
- Sales and profit by category and sub-category (pivot_sales_profit_category)
- Order count by ship mode (pivot_order_count_ship_mode)
- Average profit margin by customer segment(pivot_profit_margin_segment) 
  — uses cleaned profit margin (calculated_profit / calculated_sales), not raw profit/sales
- Refunded/cancelled/failed orders by region (pivot_refund_cancel_failed_region) 
  — includes payment_status filter
- Monthly sales trend (pivot_monthly_sales_trend) — nested by order_year and order_month

### Method Used
- Copied cleaned data from cleaned_orders.xlsx (Table1 sheet) into a new source_data sheet inside pivot_summary.xlsx
- Built each pivot table from this source_data using Excel's native PivotTable feature
- Sum aggregation used for sales/profit totals; Average aggregation used for profit margin by segment, since margins should not be summed across rows
- All values formatted to 2 decimal places using Value Field Settings

### Sorting and Filtering Applied
- pivot_sales_profit_region sorted descending by calculated_sales
- pivot_refund_cancel_failed_region includes a payment_status filter field

### Assumptions Made
- Pivot calculations use cleaned/calculated values (calculated_sales, calculated_profit, profit_margin) rather than original raw sales/profit columns, to reflect corrected data
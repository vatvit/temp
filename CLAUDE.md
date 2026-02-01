# Father - Excel Analysis Project

## Overview
Script to analyse and compare data in two Excel files (tovar vs база).
Finds discrepancies in sales data between the two systems.

## Files
- `tovar2025_PRN.xlsx` - 2025 product data (monthly sales sheets)
- `База товара_PRN.xls` - Product database (legacy format)

## Running Commands

**Always use Docker** - do not run Python/pip directly on host:

```bash
# Build container
docker compose build

# Run analysis script
docker compose run --rm app python src/main.py

# Interactive shell
docker compose run --rm app bash
```

## Project Structure
```
/src/main.py  - Main analysis script
/src/debug.py - Debug script for specific row comparison
/output       - Analysis results (not committed)
```

## Output Files
After running the script:
- `output/comparison_report.xlsx` - All months in one Excel (sheets: Summary, Gap Items, All Issues, per-type sheets)
- `output/comparison_summary.csv` - Monthly summary with gaps
- `output/comparison_details.csv` - All issues from all months

## Algorithm

### Key Formulas
- **expected_diff** = sold_tovar - sold_база (from КупиПродай sheet, rows 9-10)
- **computed_diff** = NOT_FOUND_sum - БАЗА_EXTRA_sum
- **gap** = expected_diff - computed_diff (unexplained discrepancy)
- **check_items** = sum of WRONG_DATE + DATE_ANOMALY + БАЗА_NO_DATE (items that may explain gap)

### Matching Logic
1. Match tovar rows to база by columns A-H with tolerance
2. Prioritize matches by date quality: exact > same month > valid date > empty date
3. Track unmatched items on both sides

### Tolerance Rules (columns A-H)
- Allow max 1 hard mismatch in columns A-H
- Year typos in column B (date) are tolerated
- Dash/empty equivalence (e.g., "-" = empty)
- F+H price difference treated as single tolerance (same item, updated price)

### Date Matching Priority
1. Exact date match (score 4)
2. Same month, different day (score 3)
3. Valid date but different month (score 2)
4. Empty база date (score 1) - worst priority

## Issue Types

| Type | Description | Impact on Gap |
|------|-------------|---------------|
| TOLERANCE | Matched with tolerance (year typos, dash/empty, 1-col diff) | None - successfully matched |
| WRONG_DATE | Matched by A-H but sell dates differ (within 1 month) | Usually none |
| DATE_ANOMALY | ⚠️ Matched but dates differ by MULTIPLE months | May indicate wrong match |
| БАЗА_NO_DATE | ★ Matched but база has EMPTY sell date | **Primary cause of gap** |
| POTENTIAL | Same name/date but A-H differs too much | May indicate data error |
| NOT_FOUND | Not found in база | Adds to computed_diff |
| БАЗА_EXTRA | In база but not matched to tovar | Subtracts from computed_diff |

### Gap Investigation Priority
1. **БАЗА_NO_DATE** - most likely cause (база item not counted in sold_база)
2. **DATE_ANOMALY** - suspicious matches, may be wrong pairs
3. **POTENTIAL** - possible data entry errors
4. **WRONG_DATE** - usually OK but check if POTENTIAL exists for same row

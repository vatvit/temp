# Father - Excel Analysis Project

## Overview
Script to analyse and compare data in two Excel files.

## Files
- `tovar2025_PRN.xlsx` - 2025 product data
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
/src          - Python source code
/output       - Analysis results
```

## Output Files
After running `./run.sh`:
- `output/comparison_report.xlsx` - All months in one Excel (sheets: Summary, All Issues, per-type sheets)
- `output/comparison_summary.csv` - Monthly summary with gaps
- `output/comparison_details.csv` - All issues from all months

### Issue Types
- `TOLERANCE` - Matched with tolerance (year typos, dash/empty, 1-col diff)
- `WRONG_DATE` - Matched but sell dates differ
- `POTENTIAL` - Same name/date but A-H differs too much (possible data error)
- `NOT_FOUND` - Not found in база
- `БАЗА_EXTRA` - In база but not matched to tovar

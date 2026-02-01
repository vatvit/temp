#!/usr/bin/env python3
"""
Excel Analysis - Compare tovar vs база for all months
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

pd.set_option('display.max_columns', 8)
pd.set_option('display.max_rows', 20)
pd.set_option('display.width', 150)

BASE_PATH = Path("/app")
OUTPUT_PATH = BASE_PATH / "output"
FILE_TOVAR = BASE_PATH / "tovar2025_PRN.xlsx"
FILE_БАЗА = BASE_PATH / "База товара_PRN.xls"

# Month configuration: (sheet_name_sales, column_index_in_КупиПродай, month_number)
MONTHS = [
    ("Январь продажи", 1, 1, "Январь"),
    ("Февраль продажи", 2, 2, "Февраль"),
    ("Март продажи", 3, 3, "Март"),
    ("Апрель продажи", 4, 4, "Апрель"),
    ("Май продажи", 5, 5, "Май"),
    ("Июнь продажи", 6, 6, "Июнь"),
    ("Июль продажи", 7, 7, "Июль"),
    ("Август продажи", 8, 8, "Август"),
    ("Сентябрь продажи", 9, 9, "Сентябрь"),
    ("Октябрь продажи", 10, 10, "Октябрь"),
    ("Ноябрь продажи", 11, 11, "Ноябрь"),
    ("Декабрь продажи", 12, 12, "Декабрь"),
]


def get_month_difference(col_idx):
    """Get expected difference from КупиПродай sheet for a month."""
    df = pd.read_excel(FILE_TOVAR, sheet_name="КупиПродай", header=None)
    try:
        sold_tovar = float(df.iloc[8, col_idx])
        sold_база = float(df.iloc[9, col_idx])
        if pd.isna(sold_tovar) or pd.isna(sold_база):
            return None, None, None
        diff = sold_tovar - sold_база
        return sold_tovar, sold_база, diff
    except:
        return None, None, None


def load_month_sales(sheet_name):
    """Load sales from a month sheet."""
    try:
        df = pd.read_excel(FILE_TOVAR, sheet_name=sheet_name, header=None)
    except:
        return []

    sales = []
    current_day = None

    for idx, row in df.iterrows():
        val_a = row[0]
        val_v = row[21]

        if pd.isna(val_a) or val_a == "Наименование товара":
            continue

        if isinstance(val_a, datetime):
            current_day = val_a
            continue

        sales.append({
            'row': idx + 1,
            'name': val_a,
            'cols_a_h': tuple(row[0:8].tolist()),
            'price': row[7],
            'sell_date': val_v,
            'day_marker': current_day
        })

    return sales


def load_база_sheets():
    """Load all product sheets from База товара."""
    xl = pd.ExcelFile(FILE_БАЗА, engine="xlrd")
    sheets = {}
    skip = ['КупиПродай', 'Курс']

    for name in xl.sheet_names:
        if name in skip:
            continue
        df = pd.read_excel(FILE_БАЗА, sheet_name=name, header=None)
        sheets[name] = df.iloc[2:].reset_index(drop=True)

    return sheets


def is_empty_equivalent(val):
    """Check if value is empty or dash placeholder."""
    if val is None or pd.isna(val):
        return True
    if isinstance(val, str) and val.strip() in ('', '-', '—', '–'):
        return True
    return False


def normalize_tuple(t):
    """Normalize tuple for comparison."""
    result = []
    for v in t:
        if pd.isna(v):
            result.append(None)
        elif isinstance(v, float):
            result.append(round(v, 4))
        else:
            result.append(v)
    return tuple(result)


def match_with_tolerance(sale_key, row_key):
    """Compare A-H tuples with tolerance. Allow max 1 hard mismatch.

    Special handling for F+H price differences:
    - If only F (unit price) and H (total) differ, treat as single "price_diff" tolerance
    - This handles cases where prices were updated but it's the same item
    """
    if sale_key == row_key:
        return True, []

    tolerances = []
    hard_mismatches = 0
    mismatch_cols = []  # Track which columns mismatch

    for i, (s, r) in enumerate(zip(sale_key, row_key)):
        if s == r:
            continue
        if s is None and r is None:
            continue

        if i == 1 and isinstance(s, datetime) and isinstance(r, datetime):
            if s.month == r.month and s.day == r.day and s.year != r.year:
                tolerances.append(f"col_B_year:{s.year}vs{r.year}")
                continue

        if is_empty_equivalent(s) and is_empty_equivalent(r):
            tolerances.append(f"col{i}_dash_vs_empty")
            continue

        hard_mismatches += 1
        mismatch_cols.append(i)
        tolerances.append(f"col{i}_diff:{fmt_val(s)}vs{fmt_val(r)}")

    # Special case: if only F (col 5) and H (col 7) differ, treat as single price difference
    # This handles price updates where the item is the same
    if hard_mismatches == 2 and set(mismatch_cols) == {5, 7}:
        # Replace the two tolerances with a single price_diff tolerance
        tolerances = [t for t in tolerances if not t.startswith('col5_diff') and not t.startswith('col7_diff')]
        tolerances.append(f"price_diff:F={fmt_val(sale_key[5])}vs{fmt_val(row_key[5])},H={fmt_val(sale_key[7])}vs{fmt_val(row_key[7])}")
        return True, tolerances

    if hard_mismatches > 1:
        return False, []

    return True, tolerances


def check_date_match(date1, date2):
    """Check if two dates match, with tolerance for year typos and same month.

    Returns:
        (is_match, tolerances, date_gap_months)
        - is_match: True if dates should be considered matching
        - tolerances: List of tolerance descriptions applied
        - date_gap_months: How many months apart the dates are (for flagging anomalies)
    """
    if date1 == date2:
        return True, [], 0
    if pd.isna(date1) or pd.isna(date2):
        return False, [], None
    if isinstance(date1, datetime) and isinstance(date2, datetime):
        # Calculate month gap
        month_gap = abs((date1.year - date2.year) * 12 + (date1.month - date2.month))

        # Exact day match but different year (typo)
        if date1.month == date2.month and date1.day == date2.day and date1.year != date2.year:
            return True, [f"col_V_year:{date1.year}vs{date2.year}"], month_gap

        # Same month tolerance (same year and month, different day)
        if date1.year == date2.year and date1.month == date2.month:
            return True, [f"same_month:{date1.day}vs{date2.day}"], 0

        return False, [], month_gap
    return False, [], None


def fmt_val(v):
    """Format value for display."""
    if pd.isna(v):
        return "empty"
    if isinstance(v, datetime):
        return v.strftime('%d.%m.%y')
    return str(v)[:12]


def fmt_date(d):
    """Format date for output."""
    if pd.isna(d):
        return ""
    if hasattr(d, 'strftime'):
        return d.strftime('%d.%m.%Y')
    return str(d)


def get_mismatches(sale_key, row_key):
    """Get list of mismatched columns."""
    cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    mismatches = []
    for i, (s, r) in enumerate(zip(sale_key, row_key)):
        if s == r:
            continue
        if s is None and r is None:
            continue
        if is_empty_equivalent(s) and is_empty_equivalent(r):
            continue
        mismatches.append(f"{cols[i]}:{fmt_val(s)}vs{fmt_val(r)}")
    return mismatches


def score_date_match(tovar_date, база_date):
    """Score how well база date matches tovar date. Higher is better.

    Returns:
        4: Exact match
        3: Same month (different day)
        2: Valid date but different month
        1: Empty база date (worst - might not be in sold_база totals)
    """
    if pd.isna(база_date):
        return 1  # Empty date - worst priority

    if not isinstance(база_date, datetime):
        return 1

    if pd.isna(tovar_date) or not isinstance(tovar_date, datetime):
        return 2  # Can't compare, but база has a date

    if tovar_date == база_date:
        return 4  # Exact match

    if tovar_date.year == база_date.year and tovar_date.month == база_date.month:
        return 3  # Same month

    return 2  # Different month but has a date


def find_in_база(sale, база_sheets):
    """Find matching record in База товара.

    Returns matches sorted by date priority:
    1. Exact date match
    2. Same month
    3. Different month but has date
    4. Empty date (last resort)
    """
    name = sale['name']
    sale_key = normalize_tuple(sale['cols_a_h'])
    tovar_date = sale['sell_date']
    matches = []

    for sheet_name, df in база_sheets.items():
        mask = df[0] == name
        if not mask.any():
            continue

        for idx in df[mask].index:
            row = df.loc[idx]
            row_key = normalize_tuple(tuple(row[0:8].tolist()))
            is_match, tolerances = match_with_tolerance(sale_key, row_key)

            if is_match:
                база_date = row[21]
                date_score = score_date_match(tovar_date, база_date)
                matches.append({
                    'sheet': sheet_name,
                    'row': idx + 3,
                    'cols_a_h': row_key,
                    'price': row[7],
                    'sell_date': база_date,
                    'tolerances': tolerances,
                    'date_score': date_score
                })

    # Sort by date_score descending (best matches first)
    matches.sort(key=lambda m: m['date_score'], reverse=True)

    return matches


def find_potential_matches(sale, база_sheets):
    """Find база rows with same name and sell date, but A-H too different."""
    name = sale['name']
    sell_date = sale['sell_date']
    sale_key = normalize_tuple(sale['cols_a_h'])
    potentials = []

    for sheet_name, df in база_sheets.items():
        mask = df[0] == name
        if not mask.any():
            continue

        for idx in df[mask].index:
            row = df.loc[idx]
            row_sell_date = row[21]

            date_match, _, _ = check_date_match(sell_date, row_sell_date)
            if not date_match:
                continue

            row_key = normalize_tuple(tuple(row[0:8].tolist()))
            is_match, _ = match_with_tolerance(sale_key, row_key)
            if is_match:
                continue

            mismatches = get_mismatches(sale_key, row_key)
            potentials.append({
                'sheet': sheet_name,
                'row': idx + 3,
                'cols_a_h': row_key,
                'price': row[7],
                'sell_date': row_sell_date,
                'mismatches': mismatches
            })

    return potentials


def process_month(month_name, sheet_name, col_idx, month_num, база_sheets):
    """Process a single month and return results."""
    sold_tovar, sold_база, expected_diff = get_month_difference(col_idx)

    if expected_diff is None:
        return None  # No data for this month

    sales = load_month_sales(sheet_name)
    if not sales:
        return None

    matched = []
    wrong_date = []
    date_anomaly = []  # NEW: matches where dates are many months apart (suspicious)
    not_found = []
    tolerance_applied = []
    potential_matches = []
    база_no_date = []  # matches where база has empty sell date
    used_база = set()

    for sale in sales:
        matches = find_in_база(sale, база_sheets)

        if not matches:
            potential = find_potential_matches(sale, база_sheets)
            if potential:
                not_found.append(sale)
                for p in potential:
                    potential_matches.append({
                        'sale': sale, 'база': p, 'mismatches': p['mismatches']
                    })
            else:
                not_found.append(sale)
            continue

        found_match = False
        for m in matches:
            key = (m['sheet'], m['row'])
            if key in used_база:
                continue

            sell_match, sell_tol, date_gap = check_date_match(sale['sell_date'], m['sell_date'])
            if sell_match:
                used_база.add(key)
                all_tol = m['tolerances'] + sell_tol
                matched.append({'sale': sale, 'база': m})
                if all_tol:
                    tolerance_applied.append({
                        'sale': sale, 'база': m, 'tolerances': all_tol
                    })
                found_match = True
                break

        if not found_match:
            for m in matches:
                key = (m['sheet'], m['row'])
                if key not in used_база:
                    used_база.add(key)

                    # Check if база has empty sell date - this is a special case
                    база_date_empty = pd.isna(m['sell_date'])

                    if база_date_empty:
                        # Track as БАЗА_NO_DATE - matched but база has no sell date
                        # This often causes gaps because база item might not be in sold_база totals
                        база_no_date.append({
                            'sale': sale, 'база': m,
                            'tovar_date': sale['sell_date'],
                            'база_date': m['sell_date']
                        })
                    else:
                        # Calculate date gap to detect anomalies
                        _, _, date_gap = check_date_match(sale['sell_date'], m['sell_date'])
                        if date_gap is not None and date_gap > 1:
                            # More than 1 month apart - this is suspicious!
                            date_anomaly.append({
                                'sale': sale, 'база': m,
                                'tovar_date': sale['sell_date'],
                                'база_date': m['sell_date'],
                                'months_apart': date_gap
                            })
                        else:
                            wrong_date.append({
                                'sale': sale, 'база': m,
                                'tovar_date': sale['sell_date'],
                                'база_date': m['sell_date']
                            })

                    if m['tolerances']:
                        tolerance_applied.append({
                            'sale': sale, 'база': m, 'tolerances': m['tolerances']
                        })
                    potential = find_potential_matches(sale, база_sheets)
                    for p in potential:
                        if (p['sheet'], p['row']) not in used_база:
                            potential_matches.append({
                                'sale': sale, 'база': p, 'mismatches': p['mismatches']
                            })
                    found_match = True
                    break
            if not found_match:
                wrong_date.append({
                    'sale': sale, 'база': matches[0],
                    'tovar_date': sale['sell_date'], 'база_date': 'ALL_USED'
                })

    # Find база items with this month's dates NOT matched
    база_unmatched = []
    for sheet_name_b, df in база_sheets.items():
        for idx, row in df.iterrows():
            sell_date = row[21]
            if pd.isna(sell_date) or not isinstance(sell_date, datetime):
                continue
            if sell_date.year == 2025 and sell_date.month == month_num:
                key = (sheet_name_b, idx + 3)
                if key not in used_база:
                    база_unmatched.append({
                        'sheet': sheet_name_b,
                        'row': idx + 3,
                        'name': row[0],
                        'price': row[7],
                        'date': sell_date
                    })

    def safe_float(v):
        try:
            return float(v) if pd.notna(v) else 0
        except:
            return 0

    not_found_sum = sum(safe_float(s['price']) for s in not_found)
    unmatched_sum = sum(safe_float(b['price']) for b in база_unmatched)
    computed_diff = not_found_sum - unmatched_sum

    return {
        'month': month_name,
        'month_num': month_num,
        'sold_tovar': sold_tovar,
        'sold_база': sold_база,
        'expected_diff': expected_diff,
        'computed_diff': computed_diff,
        'matched': matched,
        'wrong_date': wrong_date,
        'date_anomaly': date_anomaly,  # NEW: matches where dates are months apart
        'база_no_date': база_no_date,  # matches where база has no sell date
        'not_found': not_found,
        'tolerance_applied': tolerance_applied,
        'potential_matches': potential_matches,
        'база_unmatched': база_unmatched,
        'used_база': used_база
    }


def save_combined_results(all_results):
    """Save all results to combined CSV and XLS files."""
    OUTPUT_PATH.mkdir(exist_ok=True)

    all_rows = []

    for r in all_results:
        month = r['month']

        # Add tolerance items
        for t in r['tolerance_applied']:
            all_rows.append({
                'month': month,
                'type': 'TOLERANCE',
                'tovar_row': t['sale']['row'],
                'tovar_name': t['sale']['name'],
                'tovar_price_H': t['sale']['price'],
                'tovar_sell_date': fmt_date(t['sale']['sell_date']),
                'база_sheet': t['база']['sheet'],
                'база_row': t['база']['row'],
                'база_sell_date': fmt_date(t['база'].get('sell_date', '')),
                'issue': ', '.join(t['tolerances'])
            })

        # Add wrong date items
        for d in r['wrong_date']:
            all_rows.append({
                'month': month,
                'type': 'WRONG_DATE',
                'tovar_row': d['sale']['row'],
                'tovar_name': d['sale']['name'],
                'tovar_price_H': d['sale']['price'],
                'tovar_sell_date': fmt_date(d['tovar_date']),
                'база_sheet': d['база']['sheet'],
                'база_row': d['база']['row'],
                'база_sell_date': fmt_date(d['база_date']) if d['база_date'] != 'ALL_USED' else 'ALL_USED',
                'issue': f"tovar:{fmt_date(d['tovar_date'])} vs база:{fmt_date(d['база_date']) if d['база_date'] != 'ALL_USED' else 'ALL_USED'}"
            })

        # Add база_no_date items (matched but база has empty sell date)
        for d in r['база_no_date']:
            all_rows.append({
                'month': month,
                'type': 'БАЗА_NO_DATE',
                'tovar_row': d['sale']['row'],
                'tovar_name': d['sale']['name'],
                'tovar_price_H': d['sale']['price'],
                'tovar_sell_date': fmt_date(d['tovar_date']),
                'база_sheet': d['база']['sheet'],
                'база_row': d['база']['row'],
                'база_sell_date': '(пусто)',
                'issue': f"tovar:{fmt_date(d['tovar_date'])} matched to база with EMPTY date - likely causing gap!"
            })

        # Add date_anomaly items (matched but dates are many months apart)
        for d in r['date_anomaly']:
            all_rows.append({
                'month': month,
                'type': 'DATE_ANOMALY',
                'tovar_row': d['sale']['row'],
                'tovar_name': d['sale']['name'],
                'tovar_price_H': d['sale']['price'],
                'tovar_sell_date': fmt_date(d['tovar_date']),
                'база_sheet': d['база']['sheet'],
                'база_row': d['база']['row'],
                'база_sell_date': fmt_date(d['база_date']),
                'issue': f"⚠️ {d['months_apart']} мес. разница! tovar:{fmt_date(d['tovar_date'])} vs база:{fmt_date(d['база_date'])}"
            })

        # Add potential matches
        for p in r['potential_matches']:
            all_rows.append({
                'month': month,
                'type': 'POTENTIAL',
                'tovar_row': p['sale']['row'],
                'tovar_name': p['sale']['name'],
                'tovar_price_H': p['sale']['price'],
                'tovar_sell_date': fmt_date(p['sale']['sell_date']),
                'база_sheet': p['база']['sheet'],
                'база_row': p['база']['row'],
                'база_sell_date': fmt_date(p['база'].get('sell_date', '')),
                'issue': ', '.join(p['mismatches'][:4])
            })

        # Add not found items
        for s in r['not_found']:
            # Skip if already in potential matches
            if any(p['sale']['row'] == s['row'] for p in r['potential_matches']):
                continue
            all_rows.append({
                'month': month,
                'type': 'NOT_FOUND',
                'tovar_row': s['row'],
                'tovar_name': s['name'],
                'tovar_price_H': s['price'],
                'tovar_sell_date': fmt_date(s['sell_date']),
                'база_sheet': '',
                'база_row': '',
                'база_sell_date': '',
                'issue': 'Not found in база'
            })

        # Add база unmatched items
        for b in r['база_unmatched']:
            all_rows.append({
                'month': month,
                'type': 'БАЗА_EXTRA',
                'tovar_row': '',
                'tovar_name': '',
                'tovar_price_H': '',
                'tovar_sell_date': '',
                'база_sheet': b['sheet'],
                'база_row': b['row'],
                'база_sell_date': fmt_date(b['date']),
                'issue': f"{b['name'][:40]} H={b['price']}"
            })

    df_all = pd.DataFrame(all_rows)

    # Helper for safe float conversion
    def safe_float(v):
        try:
            return float(v) if pd.notna(v) else 0
        except:
            return 0

    # Summary per month
    summary_rows = []
    for r in all_results:
        gap = r['expected_diff'] - r['computed_diff']
        wrong_date_sum = sum(safe_float(d['sale']['price']) for d in r['wrong_date'])
        база_no_date_sum = sum(safe_float(d['sale']['price']) for d in r['база_no_date'])
        date_anomaly_sum = sum(safe_float(d['sale']['price']) for d in r['date_anomaly'])

        # Sum of all items that might explain the gap
        # Priority: date issues > tolerance (if gap exists)
        tolerance_sum = sum(safe_float(t['sale']['price']) for t in r['tolerance_applied'])
        issues_sum = wrong_date_sum + база_no_date_sum + date_anomaly_sum

        # Build check_items - just a number to compare with gap
        if issues_sum > 0:
            check_hint = round(issues_sum, 2)
        elif abs(gap) > 0.01 and tolerance_sum > 0:
            check_hint = round(tolerance_sum, 2)
        else:
            check_hint = ''

        summary_rows.append({
            'month': r['month'],
            'sold_tovar': r['sold_tovar'],
            'sold_база': r['sold_база'],
            'expected_diff': r['expected_diff'],
            'computed_diff': r['computed_diff'],
            'gap': gap,
            'check_items': check_hint,
            'matched': len(r['matched']),
            'tolerance': len(r['tolerance_applied']),
            'wrong_date': len(r['wrong_date']),
            'date_anomaly': len(r['date_anomaly']),
            'база_no_date': len(r['база_no_date']),
            'potential': len(r['potential_matches']),
            'not_found': len(r['not_found']),
            'база_extra': len(r['база_unmatched'])
        })
    df_summary = pd.DataFrame(summary_rows)

    # Save CSV
    df_all.to_csv(OUTPUT_PATH / 'comparison_details.csv', index=False)
    df_summary.to_csv(OUTPUT_PATH / 'comparison_summary.csv', index=False)

    # Build gap items sheet - simple list of items creating the gap
    gap_rows = []

    def to_float(v):
        try:
            return float(v) if pd.notna(v) else 0
        except:
            return 0

    for r in all_results:
        month = r['month']
        expected = r['expected_diff']
        computed = r['computed_diff']
        gap = expected - computed

        # Month header
        gap_rows.append({
            'Месяц': f'=== {month} ===',
            'Тип': '',
            'Лист': '',
            'Строка': '',
            'Название': f'Ожидаемая разница: {expected:.2f}',
            'Сумма H': '',
            'Влияние': '',
            'Остаток': ''
        })

        running_total = 0

        # NOT_FOUND items (add to tovar side - positive)
        for s in r['not_found']:
            price = to_float(s['price'])
            running_total += price
            gap_rows.append({
                'Месяц': month,
                'Тип': 'tovar (не найдено)',
                'Лист': 'tovar ' + month,
                'Строка': s['row'],
                'Название': s['name'][:45] if s['name'] else '',
                'Сумма H': round(price, 2),
                'Влияние': f'+{price:.2f}',
                'Остаток': round(running_total, 2)
            })

        # БАЗА_EXTRA items (add to база side - negative)
        for b in r['база_unmatched']:
            price = to_float(b['price'])
            running_total -= price
            gap_rows.append({
                'Месяц': month,
                'Тип': 'База (лишнее)',
                'Лист': b['sheet'],
                'Строка': b['row'],
                'Название': str(b['name'])[:45] if b['name'] else '',
                'Сумма H': round(price, 2),
                'Влияние': f'-{price:.2f}',
                'Остаток': round(running_total, 2)
            })

        not_found_sum = sum(to_float(s['price']) for s in r['not_found'])
        база_sum = sum(to_float(b['price']) for b in r['база_unmatched'])

        # Month summary
        gap_rows.append({
            'Месяц': month,
            'Тип': '>>> ИТОГО',
            'Лист': '',
            'Строка': '',
            'Название': f'tovar: +{not_found_sum:.2f}, База: -{база_sum:.2f}',
            'Сумма H': '',
            'Влияние': f'= {running_total:.2f}',
            'Остаток': ''
        })
        gap_rows.append({
            'Месяц': '',
            'Тип': '>>> СВЕРКА',
            'Лист': '',
            'Строка': '',
            'Название': f'Ожидалось: {expected:.2f}, Вычислено: {running_total:.2f}',
            'Сумма H': '',
            'Влияние': f'Разница (gap): {gap:.2f}',
            'Остаток': ''
        })

        # Add tip if there are issue items
        база_no_date_sum = sum(to_float(d['sale']['price']) for d in r['база_no_date'])
        date_anomaly_sum = sum(to_float(d['sale']['price']) for d in r['date_anomaly'])
        wrong_date_sum = sum(to_float(d['sale']['price']) for d in r['wrong_date'])
        potential_count = len(r['potential_matches'])

        if abs(gap) > 0.01 and (база_no_date_sum > 0 or date_anomaly_sum > 0 or wrong_date_sum > 0 or potential_count > 0):
            tip_parts = []
            explain_sum = 0

            # БАЗА_NO_DATE is the primary cause of gaps
            if база_no_date_sum > 0:
                tip_parts.append(f'★ БАЗА_NO_DATE: {база_no_date_sum:.2f}')
                explain_sum += база_no_date_sum

            # DATE_ANOMALY is suspicious - might be wrong matches
            if date_anomaly_sum > 0:
                tip_parts.append(f'⚠️ DATE_ANOMALY: {date_anomaly_sum:.2f}')
                # Don't add to explain_sum - these need verification

            if wrong_date_sum > 0:
                tip_parts.append(f'WRONG_DATE: {wrong_date_sum:.2f}')
                # WRONG_DATE doesn't directly explain gap, but might indicate issues

            if potential_count > 0:
                tip_parts.append(f'POTENTIAL: {potential_count} поз.')
                # Note: POTENTIAL impact is uncertain, don't add to explain_sum

            remaining = abs(gap) - explain_sum
            if remaining < 0:
                remaining = 0

            gap_rows.append({
                'Месяц': '',
                'Тип': '>>> ПОДСКАЗКА',
                'Лист': '',
                'Строка': '',
                'Название': f'Проверьте: {", ".join(tip_parts)}',
                'Сумма H': '',
                'Влияние': f'БАЗА_NO_DATE объясняет {explain_sum:.2f} из {abs(gap):.2f}, остаток ~{remaining:.2f}',
                'Остаток': ''
            })
        # Empty row between months
        gap_rows.append({
            'Месяц': '', 'Тип': '', 'Лист': '', 'Строка': '',
            'Название': '', 'Сумма H': '', 'Влияние': '', 'Остаток': ''
        })

    df_gap = pd.DataFrame(gap_rows)

    # Russian descriptions for each sheet
    descriptions = {
        'Summary': [
            "СВОДКА ПО МЕСЯЦАМ",
            "Этот лист показывает итоги сверки за каждый месяц.",
            "",
            "Как читать:",
            "• sold_tovar - сумма продаж в файле 'tovar' (лист КупиПродай, строка 9)",
            "• sold_база - сумма продаж в файле 'База товара' (лист КупиПродай, строка 10)",
            "• expected_diff - ожидаемая разница (sold_tovar - sold_база)",
            "• computed_diff - вычисленная разница по позициям",
            "• gap - расхождение (если не 0, значит есть проблемы)",
            "• check_items - сумма проблемных позиций (сравните с gap)",
            "• matched - сколько позиций совпало",
            "• tolerance - совпало с допуском (опечатки в году, прочерк вместо пустого и т.д.)",
            "• wrong_date - совпало, но дата продажи отличается (в пределах 1 мес.)",
            "• date_anomaly - ⚠️ совпало, но даты СИЛЬНО отличаются (может быть ошибка!)",
            "• база_no_date - ★ ВАЖНО: совпало, но в Базе ПУСТАЯ дата (причина gap!)",
            "• potential - возможное совпадение (то же название и дата, но другая партия)",
            "• not_found - не найдено в Базе",
            "• база_extra - есть в Базе, но не найдено в tovar",
            "",
        ],
        'All Issues': [
            "ВСЕ ПРОБЛЕМЫ",
            "Этот лист содержит все найденные расхождения за все месяцы.",
            "",
            "Типы проблем (колонка type):",
            "• TOLERANCE - совпало с допуском (мелкие опечатки)",
            "• WRONG_DATE - дата продажи отличается (в пределах 1 месяца)",
            "• DATE_ANOMALY - ⚠️ даты отличаются на НЕСКОЛЬКО месяцев (подозрительно!)",
            "• БАЗА_NO_DATE - ★ совпало, но в Базе ПУСТАЯ дата (главная причина gap!)",
            "• POTENTIAL - возможное совпадение, но слишком много отличий в колонках A-H",
            "• NOT_FOUND - позиция из tovar не найдена в Базе",
            "• БАЗА_EXTRA - позиция в Базе с датой этого месяца, но не найдена в tovar",
            "",
        ],
        'TOLERANCE': [
            "ДОПУСКИ (мелкие опечатки)",
            "Эти позиции СОВПАЛИ, но с небольшими отличиями.",
            "",
            "★ ЭТИ ПОЗИЦИИ НЕ ВЛИЯЮТ НА РАСХОЖДЕНИЕ (gap) ★",
            "Они успешно сопоставлены, просто есть мелкие ошибки в данных.",
            "",
            "Виды допусков (колонка issue):",
            "• col_B_year - опечатка в году в колонке B (дата поставки)",
            "• col_V_year - опечатка в году в колонке V (дата продажи)",
            "• col3_dash_vs_empty - в одном файле прочерк '-', в другом пусто",
            "• col5_diff - небольшое отличие в колонке F (допускается 1 отличие)",
            "",
            "Действие: Проверить и исправить опечатки в исходных файлах.",
            "",
        ],
        'WRONG_DATE': [
            "НЕПРАВИЛЬНАЯ ДАТА ПРОДАЖИ",
            "Позиция найдена в Базе, но дата продажи (колонка V) отличается.",
            "",
            "★ ЭТИ ПОЗИЦИИ НЕ ВЛИЯЮТ НА РАСХОЖДЕНИЕ (gap) ★",
            "Они сопоставлены по колонкам A-H, но дата продажи разная.",
            "Однако они могут объяснять расхождение, если правильная пара",
            "находится в POTENTIAL (возможные совпадения).",
            "",
            "Это значит:",
            "• В tovar указана одна дата продажи",
            "• В Базе товара для этой же позиции указана другая дата",
            "",
            "Действие: Проверить какая дата правильная и исправить.",
            "Колонка issue показывает: tovar_дата vs база_дата",
            "",
        ],
        'БАЗА_NO_DATE': [
            "БАЗА БЕЗ ДАТЫ ПРОДАЖИ",
            "Позиция найдена в Базе по колонкам A-H, но в Базе ПУСТАЯ дата продажи.",
            "",
            "★★★ ЭТИ ПОЗИЦИИ СКОРЕЕ ВСЕГО ПРИЧИНА РАСХОЖДЕНИЯ (gap) ★★★",
            "",
            "Почему это проблема:",
            "• В tovar позиция помечена как проданная (есть дата)",
            "• В Базе эта позиция НЕ помечена как проданная (дата пустая)",
            "• Поэтому в КупиПродай: tovar считает её проданной, база - нет",
            "• Это создаёт разницу (gap) равную сумме H этих позиций",
            "",
            "Действие: Заполнить дату продажи в Базе товара!",
            "После исправления gap должен уменьшиться на сумму H этих позиций.",
            "",
        ],
        'DATE_ANOMALY': [
            "ПОДОЗРИТЕЛЬНАЯ РАЗНИЦА В ДАТАХ",
            "Позиция совпала по A-H, но даты продажи отличаются на НЕСКОЛЬКО МЕСЯЦЕВ!",
            "",
            "⚠️ ЭТИ ПОЗИЦИИ ТРЕБУЮТ ПРОВЕРКИ ⚠️",
            "",
            "Возможные проблемы:",
            "• Позиция случайно сопоставилась с ДРУГОЙ партией того же товара",
            "• Ошибка в дате продажи (в tovar или в Базе)",
            "• Товар продан позже/раньше, чем указано в одном из файлов",
            "",
            "Важно: Эти позиции могут быть НЕПРАВИЛЬНО сопоставлены!",
            "Если это разные партии - нужно найти правильную пару.",
            "",
            "Действие: Проверить вручную, правильно ли сопоставлены эти позиции.",
            "Колонка issue показывает разницу в месяцах.",
            "",
        ],
        'POTENTIAL': [
            "ВОЗМОЖНЫЕ СОВПАДЕНИЯ",
            "Найдена позиция с таким же названием и датой продажи,",
            "но слишком много отличий в колонках A-H (партия, документ и т.д.).",
            "",
            "★ ЭТИ ПОЗИЦИИ ПОМОГАЮТ ОБЪЯСНИТЬ РАСХОЖДЕНИЕ (gap) ★",
            "Если позиция из WRONG_DATE имеет POTENTIAL - возможно,",
            "она должна была сопоставиться с этой позицией вместо текущей.",
            "",
            "Это может означать:",
            "• Это разные партии одного товара",
            "• Или ошибка в данных (неправильный документ, дата поставки и т.д.)",
            "",
            "Колонка issue показывает какие колонки отличаются.",
            "Действие: Проверить вручную, должны ли эти позиции совпадать.",
            "",
        ],
        'NOT_FOUND': [
            "НЕ НАЙДЕНО В БАЗЕ",
            "Эти позиции есть в файле tovar, но не найдены в Базе товара.",
            "",
            "★ ЭТИ ПОЗИЦИИ ВЛИЯЮТ НА РАСХОЖДЕНИЕ (gap) ★",
            "Сумма (H) этих позиций добавляется к стороне tovar.",
            "Формула: gap = NOT_FOUND сумма - БАЗА_EXTRA сумма",
            "",
            "Возможные причины:",
            "• Позиция не внесена в Базу товара",
            "• Название или другие данные отличаются (опечатка)",
            "• Позиция в другом листе Базы",
            "",
            "Действие: Найти позицию в Базе вручную и исправить данные.",
            "",
        ],
        'БАЗА_EXTRA': [
            "ЛИШНЕЕ В БАЗЕ",
            "Эти позиции есть в Базе товара с датой продажи этого месяца,",
            "но не найдены в соответствующем месяце файла tovar.",
            "",
            "★ ЭТИ ПОЗИЦИИ ВЛИЯЮТ НА РАСХОЖДЕНИЕ (gap) ★",
            "Сумма (H) этих позиций добавляется к стороне База.",
            "Формула: gap = NOT_FOUND сумма - БАЗА_EXTRA сумма",
            "",
            "Возможные причины:",
            "• Позиция не внесена в tovar",
            "• Дата продажи в Базе неправильная (должен быть другой месяц)",
            "• Данные отличаются и позиция не сопоставилась",
            "",
            "Колонка issue показывает название и сумму позиции.",
            "Действие: Проверить, должна ли эта позиция быть в этом месяце.",
            "",
        ],
        'Gap Items': [
            "ПОЗИЦИИ СОЗДАЮЩИЕ РАСХОЖДЕНИЕ",
            "Простой список всех позиций, которые создают разницу.",
            "",
            "Как читать:",
            "• Каждый месяц начинается с ожидаемой разницы из КупиПродай",
            "• tovar (не найдено) - позиции из tovar без пары в Базе (+)",
            "• База (лишнее) - позиции из Базы без пары в tovar (-)",
            "",
            "Колонки:",
            "• Влияние - как позиция влияет на расчёт (+ или -)",
            "• Остаток - накопительный итог после каждой позиции",
            "",
            "В конце месяца:",
            "• ИТОГО - сумма всех позиций",
            "• СВЕРКА - сравнение ожидаемого и вычисленного, показан gap",
            "",
        ],
    }

    # Save XLS with descriptions
    with pd.ExcelWriter(OUTPUT_PATH / 'comparison_report.xlsx', engine='openpyxl') as writer:
        # Summary sheet
        df_desc = pd.DataFrame({'Описание': descriptions['Summary']})
        df_desc.to_excel(writer, sheet_name='Summary', index=False, header=False)
        df_summary.to_excel(writer, sheet_name='Summary', index=False, startrow=len(descriptions['Summary']) + 1)

        # Gap Items sheet - simple list
        df_desc = pd.DataFrame({'Описание': descriptions['Gap Items']})
        df_desc.to_excel(writer, sheet_name='Gap Items', index=False, header=False)
        df_gap.to_excel(writer, sheet_name='Gap Items', index=False, startrow=len(descriptions['Gap Items']) + 1)

        # All Issues sheet
        df_desc = pd.DataFrame({'Описание': descriptions['All Issues']})
        df_desc.to_excel(writer, sheet_name='All Issues', index=False, header=False)
        df_all.to_excel(writer, sheet_name='All Issues', index=False, startrow=len(descriptions['All Issues']) + 1)

        # Per-type sheets
        for issue_type in ['TOLERANCE', 'WRONG_DATE', 'DATE_ANOMALY', 'БАЗА_NO_DATE', 'POTENTIAL', 'NOT_FOUND', 'БАЗА_EXTRA']:
            df_type = df_all[df_all['type'] == issue_type]
            if len(df_type) > 0:
                desc = descriptions.get(issue_type, [])
                if desc:
                    df_desc = pd.DataFrame({'Описание': desc})
                    df_desc.to_excel(writer, sheet_name=issue_type[:31], index=False, header=False)
                    df_type.to_excel(writer, sheet_name=issue_type[:31], index=False, startrow=len(desc) + 1)
                else:
                    df_type.to_excel(writer, sheet_name=issue_type[:31], index=False)

    print(f"\n=== FILES SAVED ===")
    print(f"  {OUTPUT_PATH}/comparison_report.xlsx")
    print(f"  {OUTPUT_PATH}/comparison_details.csv")
    print(f"  {OUTPUT_PATH}/comparison_summary.csv")


def main():
    print("=" * 60)
    print("Excel Comparison - All Months 2025")
    print("=" * 60)

    # Load база once
    база_sheets = load_база_sheets()
    print(f"Loaded {len(база_sheets)} product sheets from База\n")

    # Check available sheets
    xl_tovar = pd.ExcelFile(FILE_TOVAR, engine="openpyxl")
    available_sheets = xl_tovar.sheet_names

    all_results = []

    for sheet_name, col_idx, month_num, month_name in MONTHS:
        if sheet_name not in available_sheets:
            continue

        result = process_month(month_name, sheet_name, col_idx, month_num, база_sheets)
        if result is None:
            continue

        all_results.append(result)

        # Console output
        gap = result['expected_diff'] - result['computed_diff']
        print(f"--- {month_name} ---")
        print(f"Expected: {result['expected_diff']:.2f} | Computed: {result['computed_diff']:.2f} | Gap: {gap:.2f}")
        print(f"Matched: {len(result['matched'])} | Tolerance: {len(result['tolerance_applied'])} | Wrong date: {len(result['wrong_date'])} | Date anomaly: {len(result['date_anomaly'])} | База no date: {len(result['база_no_date'])} | Potential: {len(result['potential_matches'])} | Not found: {len(result['not_found'])} | База extra: {len(result['база_unmatched'])}")

        # Analysis if gap exists
        if abs(gap) > 0.01:
            print(f"  Analysis: Gap {gap:.2f} may be explained by:")
            # БАЗА_NO_DATE is the primary cause
            for d in result['база_no_date']:
                print(f"    ★ tovar:{d['sale']['row']} H={d['sale']['price']} -> база has EMPTY date (likely cause!)")
            # Date anomalies are suspicious
            for d in result['date_anomaly']:
                print(f"    ⚠️ tovar:{d['sale']['row']} H={d['sale']['price']} -> {d['months_apart']} months apart!")
            for d in result['wrong_date']:
                print(f"    - tovar:{d['sale']['row']} H={d['sale']['price']} wrong date")
                for p in result['potential_matches']:
                    if p['sale']['row'] == d['sale']['row']:
                        print(f"      -> Potential база:{p['база']['sheet']}:{p['база']['row']}")
        print()

    if all_results:
        save_combined_results(all_results)

        # Final summary
        print("=" * 60)
        print("TOTAL SUMMARY")
        print("=" * 60)
        total_gap = sum(r['expected_diff'] - r['computed_diff'] for r in all_results)
        total_tolerance = sum(len(r['tolerance_applied']) for r in all_results)
        total_wrong = sum(len(r['wrong_date']) for r in all_results)
        total_date_anomaly = sum(len(r['date_anomaly']) for r in all_results)
        total_date_anomaly_sum = sum(
            sum(float(d['sale']['price']) if pd.notna(d['sale']['price']) else 0 for d in r['date_anomaly'])
            for r in all_results
        )
        total_база_no_date = sum(len(r['база_no_date']) for r in all_results)
        total_база_no_date_sum = sum(
            sum(float(d['sale']['price']) if pd.notna(d['sale']['price']) else 0 for d in r['база_no_date'])
            for r in all_results
        )
        total_potential = sum(len(r['potential_matches']) for r in all_results)
        print(f"Months processed: {len(all_results)}")
        print(f"Total gap: {total_gap:.2f}")
        print(f"Total tolerance applied: {total_tolerance}")
        print(f"Total wrong dates: {total_wrong}")
        print(f"⚠️ Total date anomalies: {total_date_anomaly} items, sum H={total_date_anomaly_sum:.2f}")
        print(f"★ Total база no date: {total_база_no_date} items, sum H={total_база_no_date_sum:.2f}")
        print(f"Total potential matches: {total_potential}")


if __name__ == "__main__":
    main()

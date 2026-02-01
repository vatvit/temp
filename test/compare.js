#!/usr/bin/env node
/**
 * Test script to compare JavaScript output with Python output
 * for the Excel comparison tool.
 *
 * Run with: docker run --rm -v "$(pwd)":/app -w /app node:20 node test/compare.js
 * Debug mode: docker run --rm -v "$(pwd)":/app -w /app node:20 sh -c "npm install xlsx 2>/dev/null && node test/compare.js --debug"
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Parse command line arguments
const DEBUG_MODE = process.argv.includes('--debug');

// Configuration
const BASE_PATH = '/app';
const FILE_TOVAR = path.join(BASE_PATH, 'tovar2025_PRN.xlsx');
const FILE_BAZA = path.join(BASE_PATH, 'База товара_PRN.xls');
const PYTHON_SUMMARY = path.join(BASE_PATH, 'output', 'comparison_summary.csv');
const PYTHON_DETAILS = path.join(BASE_PATH, 'output', 'comparison_details.csv');

// Month configuration
const MONTHS = [
    { sheetName: 'Январь продажи', colIdx: 1, monthNum: 1, monthName: 'Январь' },
    { sheetName: 'Февраль продажи', colIdx: 2, monthNum: 2, monthName: 'Февраль' },
    { sheetName: 'Март продажи', colIdx: 3, monthNum: 3, monthName: 'Март' },
    { sheetName: 'Апрель продажи', colIdx: 4, monthNum: 4, monthName: 'Апрель' },
    { sheetName: 'Май продажи', colIdx: 5, monthNum: 5, monthName: 'Май' },
    { sheetName: 'Июнь продажи', colIdx: 6, monthNum: 6, monthName: 'Июнь' },
    { sheetName: 'Июль продажи', colIdx: 7, monthNum: 7, monthName: 'Июль' },
    { sheetName: 'Август продажи', colIdx: 8, monthNum: 8, monthName: 'Август' },
    { sheetName: 'Сентябрь продажи', colIdx: 9, monthNum: 9, monthName: 'Сентябрь' },
    { sheetName: 'Октябрь продажи', colIdx: 10, monthNum: 10, monthName: 'Октябрь' },
    { sheetName: 'Ноябрь продажи', colIdx: 11, monthNum: 11, monthName: 'Ноябрь' },
    { sheetName: 'Декабрь продажи', colIdx: 12, monthNum: 12, monthName: 'Декабрь' },
];

// ============================================================================
// Utility functions (extracted from index.html)
// ============================================================================

/**
 * Normalize string for comparison: handles Cyrillic/Latin lookalikes,
 * whitespace normalization, and case-insensitive matching.
 *
 * Cyrillic letters that look identical to Latin:
 * А/A, В/B, С/C, Е/E, К/K, М/M, Н/H, О/O, Р/P, Т/T, Х/X, У/Y
 *
 * After toLowerCase(), uppercase Cyrillic becomes lowercase Cyrillic,
 * so we only need lowercase mappings for the post-lowercase phase.
 */
const CYRILLIC_TO_LATIN_LOWER = {
    // Lowercase Cyrillic -> lowercase Latin (applied after toLowerCase)
    '\u0430': 'a',  // а -> a
    '\u0432': 'b',  // в -> b (Cyrillic ve, visually similar to B/b in some fonts)
    '\u0441': 'c',  // с -> c
    '\u0435': 'e',  // е -> e
    '\u043A': 'k',  // к -> k
    '\u043C': 'm',  // м -> m
    '\u043D': 'h',  // н -> h (Cyrillic en looks like H)
    '\u043E': 'o',  // о -> o
    '\u0440': 'p',  // р -> p (Cyrillic er looks like p)
    '\u0442': 't',  // т -> t
    '\u0445': 'x',  // х -> x
    '\u0443': 'y',  // у -> y
};

function normalizeStringForComparison(str) {
    if (str === null || str === undefined) return null;
    if (typeof str !== 'string') return str;

    // Convert to lowercase first (this also lowercases Cyrillic)
    let result = str.toLowerCase();

    // Replace Cyrillic lookalikes with Latin equivalents
    result = result.split('').map(ch => CYRILLIC_TO_LATIN_LOWER[ch] || ch).join('');

    // Normalize whitespace: trim and collapse multiple spaces
    result = result.trim().replace(/\s+/g, ' ');

    return result;
}

function stringsMatchNormalized(str1, str2) {
    return normalizeStringForComparison(str1) === normalizeStringForComparison(str2);
}

function isNaN_or_null(val) {
    return val === null || val === undefined || (typeof val === 'number' && isNaN(val));
}

function isEmptyEquivalent(val) {
    if (val === null || val === undefined) return true;
    if (typeof val === 'number' && isNaN(val)) return true;
    if (typeof val === 'string' && ['', '-', '\u2014', '\u2013'].includes(val.trim())) return true;
    return false;
}

// Epsilon for floating-point comparisons (matches Python's round(v, 4) precision)
const FLOAT_EPSILON = 0.00005;

function normalizeValue(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') {
        if (isNaN(v)) return null;
        // Round to 4 decimal places (same as Python)
        return Math.round(v * 10000) / 10000;
    }
    // Convert Date objects to timestamp for consistent comparison
    if (isDate(v)) {
        return v.getTime();
    }
    return v;
}

function floatsEqual(a, b) {
    // Compare two normalized values with epsilon tolerance for floats
    if (a === b) return true;
    if (a === null || b === null) return a === b;
    if (typeof a === 'number' && typeof b === 'number') {
        return Math.abs(a - b) < FLOAT_EPSILON;
    }
    return false;
}

function normalizeTuple(arr) {
    return arr.map(normalizeValue);
}

function isDate(val) {
    return val instanceof Date && !isNaN(val.getTime());
}

function sameDate(d1, d2) {
    if (!isDate(d1) || !isDate(d2)) return false;
    return d1.getFullYear() === d2.getFullYear() &&
           d1.getMonth() === d2.getMonth() &&
           d1.getDate() === d2.getDate();
}

function formatDate(d) {
    if (!isDate(d)) return '';
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}.${month}.${year}`;
}

function formatVal(v) {
    if (isNaN_or_null(v)) return 'empty';
    if (isDate(v)) return formatDate(v);
    const s = String(v);
    return s.length > 12 ? s.substring(0, 12) : s;
}

function safeFloat(v) {
    if (isNaN_or_null(v)) return 0;
    const n = parseFloat(v);
    return isNaN(n) ? 0 : n;
}

function getSheetDataRaw(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return null;
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
}

// ============================================================================
// Core matching logic (extracted from index.html)
// ============================================================================

function matchWithTolerance(saleKey, rowKey) {
    // Check exact match first (using epsilon for floats)
    let exactMatch = true;
    for (let i = 0; i < 8; i++) {
        const s = normalizeValue(saleKey[i]);
        const r = normalizeValue(rowKey[i]);
        if (!floatsEqual(s, r) && !(s === null && r === null)) {
            exactMatch = false;
            break;
        }
    }
    if (exactMatch) return { match: true, tolerances: [] };

    const tolerances = [];
    let hardMismatches = 0;
    const mismatchCols = [];

    for (let i = 0; i < 8; i++) {
        const s = saleKey[i];
        const r = rowKey[i];
        const sn = normalizeValue(s);
        const rn = normalizeValue(r);

        // Use epsilon comparison for floats
        if (floatsEqual(sn, rn)) continue;
        if (sn === null && rn === null) continue;

        // Column B (index 1) - year typo in date
        if (i === 1 && isDate(s) && isDate(r)) {
            if (s.getMonth() === r.getMonth() && s.getDate() === r.getDate() && s.getFullYear() !== r.getFullYear()) {
                tolerances.push(`col_B_year:${s.getFullYear()}vs${r.getFullYear()}`);
                continue;
            }
        }

        // Dash/empty equivalence
        if (isEmptyEquivalent(s) && isEmptyEquivalent(r)) {
            tolerances.push(`col${i}_dash_vs_empty`);
            continue;
        }

        hardMismatches++;
        mismatchCols.push(i);
        tolerances.push(`col${i}_diff:${formatVal(s)}vs${formatVal(r)}`);
    }

    // Special case: F (col 5) and H (col 7) differ = single price difference
    if (hardMismatches === 2 && mismatchCols.includes(5) && mismatchCols.includes(7) && mismatchCols.length === 2) {
        const newTol = tolerances.filter(t => !t.startsWith('col5_diff') && !t.startsWith('col7_diff'));
        newTol.push(`price_diff:F=${formatVal(saleKey[5])}vs${formatVal(rowKey[5])},H=${formatVal(saleKey[7])}vs${formatVal(rowKey[7])}`);
        return { match: true, tolerances: newTol };
    }

    if (hardMismatches > 1) return { match: false, tolerances: [] };

    return { match: true, tolerances };
}

function checkDateMatch(date1, date2) {
    if (sameDate(date1, date2)) return { match: true, tolerances: [], gap: 0 };
    if (!isDate(date1) || !isDate(date2)) return { match: false, tolerances: [], gap: null };

    const monthGap = Math.abs((date1.getFullYear() - date2.getFullYear()) * 12 + (date1.getMonth() - date2.getMonth()));

    // Exact day match but different year (typo)
    if (date1.getMonth() === date2.getMonth() && date1.getDate() === date2.getDate() && date1.getFullYear() !== date2.getFullYear()) {
        return { match: true, tolerances: [`col_V_year:${date1.getFullYear()}vs${date2.getFullYear()}`], gap: monthGap };
    }

    // Same month tolerance
    if (date1.getFullYear() === date2.getFullYear() && date1.getMonth() === date2.getMonth()) {
        return { match: true, tolerances: [`same_month:${date1.getDate()}vs${date2.getDate()}`], gap: 0 };
    }

    return { match: false, tolerances: [], gap: monthGap };
}

function scoreDateMatch(tovarDate, bazaDate) {
    if (!isDate(bazaDate)) return 1;
    if (!isDate(tovarDate)) return 2;
    if (sameDate(tovarDate, bazaDate)) return 4;
    if (tovarDate.getFullYear() === bazaDate.getFullYear() && tovarDate.getMonth() === bazaDate.getMonth()) return 3;
    return 2;
}

function getMismatches(saleKey, rowKey) {
    const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
    const mismatches = [];
    for (let i = 0; i < 8; i++) {
        const s = saleKey[i];
        const r = rowKey[i];
        const sn = normalizeValue(s);
        const rn = normalizeValue(r);
        // Use epsilon comparison for floats
        if (floatsEqual(sn, rn)) continue;
        if (sn === null && rn === null) continue;
        if (isEmptyEquivalent(s) && isEmptyEquivalent(r)) continue;
        mismatches.push(`${cols[i]}:${formatVal(s)}vs${formatVal(r)}`);
    }
    return mismatches;
}

// ============================================================================
// Data loading functions
// ============================================================================

function loadBazaSheets(workbook) {
    const skip = ['КупиПродай', 'Курс'];
    const sheets = {};
    for (const name of workbook.SheetNames) {
        if (skip.includes(name)) continue;
        const data = getSheetDataRaw(workbook, name);
        if (data && data.length > 2) {
            sheets[name] = data.slice(2); // Skip first 2 rows
        }
    }
    return sheets;
}

function loadMonthSales(workbook, sheetName) {
    const data = getSheetDataRaw(workbook, sheetName);
    if (!data) return [];

    const sales = [];
    let currentDay = null;

    for (let idx = 0; idx < data.length; idx++) {
        const row = data[idx];
        const valA = row[0];
        const valV = row[21];

        if (valA === null || valA === undefined || valA === 'Наименование товара') continue;

        if (isDate(valA)) {
            currentDay = valA;
            continue;
        }

        sales.push({
            row: idx + 1,
            name: valA,
            colsAH: row.slice(0, 8),
            price: row[7],
            sellDate: valV,
            dayMarker: currentDay
        });
    }

    return sales;
}

function getMonthDifference(workbook, colIdx) {
    const data = getSheetDataRaw(workbook, 'КупиПродай');
    if (!data || data.length < 10) return { soldTovar: null, soldBaza: null, diff: null };

    try {
        const soldTovar = safeFloat(data[8][colIdx]);
        const soldBaza = safeFloat(data[9][colIdx]);
        if (isNaN(soldTovar) || isNaN(soldBaza)) return { soldTovar: null, soldBaza: null, diff: null };
        return { soldTovar, soldBaza, diff: soldTovar - soldBaza };
    } catch (e) {
        return { soldTovar: null, soldBaza: null, diff: null };
    }
}

// ============================================================================
// Search functions
// ============================================================================

function findInBaza(sale, bazaSheets) {
    const name = sale.name;
    const normalizedName = normalizeStringForComparison(name);
    const saleKey = normalizeTuple(sale.colsAH);
    const tovarDate = sale.sellDate;
    const matches = [];

    for (const [sheetName, data] of Object.entries(bazaSheets)) {
        for (let idx = 0; idx < data.length; idx++) {
            const row = data[idx];
            if (normalizeStringForComparison(row[0]) !== normalizedName) continue;

            const rowKey = normalizeTuple(row.slice(0, 8));
            const result = matchWithTolerance(saleKey, rowKey);

            if (result.match) {
                const bazaDate = row[21];
                const dateScore = scoreDateMatch(tovarDate, bazaDate);
                matches.push({
                    sheet: sheetName,
                    row: idx + 3,
                    colsAH: rowKey,
                    price: row[7],
                    sellDate: bazaDate,
                    tolerances: result.tolerances,
                    dateScore
                });
            }
        }
    }

    matches.sort((a, b) => b.dateScore - a.dateScore);
    return matches;
}

function findPotentialMatches(sale, bazaSheets) {
    const name = sale.name;
    const normalizedName = normalizeStringForComparison(name);
    const sellDate = sale.sellDate;
    const saleKey = normalizeTuple(sale.colsAH);
    const potentials = [];

    for (const [sheetName, data] of Object.entries(bazaSheets)) {
        for (let idx = 0; idx < data.length; idx++) {
            const row = data[idx];
            if (normalizeStringForComparison(row[0]) !== normalizedName) continue;

            const rowSellDate = row[21];
            const dateResult = checkDateMatch(sellDate, rowSellDate);
            if (!dateResult.match) continue;

            const rowKey = normalizeTuple(row.slice(0, 8));
            const matchResult = matchWithTolerance(saleKey, rowKey);
            if (matchResult.match) continue;

            const mismatches = getMismatches(saleKey, rowKey);
            potentials.push({
                sheet: sheetName,
                row: idx + 3,
                colsAH: rowKey,
                price: row[7],
                sellDate: rowSellDate,
                mismatches
            });
        }
    }

    return potentials;
}

// ============================================================================
// Main processing function
// ============================================================================

function processMonth(monthConfig, tovarWorkbook, bazaSheets) {
    const { sheetName, colIdx, monthNum, monthName } = monthConfig;

    const { soldTovar, soldBaza, diff: expectedDiff } = getMonthDifference(tovarWorkbook, colIdx);
    if (expectedDiff === null) return null;

    const sales = loadMonthSales(tovarWorkbook, sheetName);
    if (!sales.length) return null;

    const matched = [];
    const wrongDate = [];
    const dateAnomaly = [];
    const notFound = [];
    const toleranceApplied = [];
    const potentialMatches = [];
    const bazaNoDate = [];
    const usedBaza = new Set();

    for (const sale of sales) {
        const matches = findInBaza(sale, bazaSheets);

        if (!matches.length) {
            const potential = findPotentialMatches(sale, bazaSheets);
            if (potential.length) {
                notFound.push(sale);
                for (const p of potential) {
                    potentialMatches.push({ sale, baza: p, mismatches: p.mismatches });
                }
            } else {
                notFound.push(sale);
            }
            continue;
        }

        let foundMatch = false;

        for (const m of matches) {
            const key = `${m.sheet}:${m.row}`;
            if (usedBaza.has(key)) continue;

            const dateResult = checkDateMatch(sale.sellDate, m.sellDate);
            if (dateResult.match) {
                usedBaza.add(key);
                const allTol = [...m.tolerances, ...dateResult.tolerances];
                matched.push({ sale, baza: m });
                if (allTol.length) {
                    toleranceApplied.push({ sale, baza: m, tolerances: allTol });
                }
                foundMatch = true;
                break;
            }
        }

        if (!foundMatch) {
            for (const m of matches) {
                const key = `${m.sheet}:${m.row}`;
                if (usedBaza.has(key)) continue;

                usedBaza.add(key);
                const bazaDateEmpty = !isDate(m.sellDate);

                if (bazaDateEmpty) {
                    bazaNoDate.push({
                        sale, baza: m,
                        tovarDate: sale.sellDate,
                        bazaDate: m.sellDate
                    });
                } else {
                    const dateResult = checkDateMatch(sale.sellDate, m.sellDate);
                    if (dateResult.gap !== null && dateResult.gap > 1) {
                        dateAnomaly.push({
                            sale, baza: m,
                            tovarDate: sale.sellDate,
                            bazaDate: m.sellDate,
                            monthsApart: dateResult.gap
                        });
                    } else {
                        wrongDate.push({
                            sale, baza: m,
                            tovarDate: sale.sellDate,
                            bazaDate: m.sellDate
                        });
                    }
                }

                if (m.tolerances.length) {
                    toleranceApplied.push({ sale, baza: m, tolerances: m.tolerances });
                }

                const potential = findPotentialMatches(sale, bazaSheets);
                for (const p of potential) {
                    const pKey = `${p.sheet}:${p.row}`;
                    if (!usedBaza.has(pKey)) {
                        potentialMatches.push({ sale, baza: p, mismatches: p.mismatches });
                    }
                }
                foundMatch = true;
                break;
            }

            if (!foundMatch) {
                wrongDate.push({
                    sale, baza: matches[0],
                    tovarDate: sale.sellDate,
                    bazaDate: 'ALL_USED'
                });
            }
        }
    }

    // Find baza items with this month's dates NOT matched
    const bazaUnmatched = [];
    for (const [sheetName, data] of Object.entries(bazaSheets)) {
        for (let idx = 0; idx < data.length; idx++) {
            const row = data[idx];
            const sellDate = row[21];
            if (!isDate(sellDate)) continue;
            if (sellDate.getFullYear() === 2025 && sellDate.getMonth() + 1 === monthNum) {
                const key = `${sheetName}:${idx + 3}`;
                if (!usedBaza.has(key)) {
                    bazaUnmatched.push({
                        sheet: sheetName,
                        row: idx + 3,
                        name: row[0],
                        price: row[7],
                        date: sellDate
                    });
                }
            }
        }
    }

    const notFoundSum = notFound.reduce((sum, s) => sum + safeFloat(s.price), 0);
    const unmatchedSum = bazaUnmatched.reduce((sum, b) => sum + safeFloat(b.price), 0);
    const computedDiff = notFoundSum - unmatchedSum;

    return {
        month: monthName,
        monthNum,
        soldTovar,
        soldBaza,
        expectedDiff,
        computedDiff,
        matched,
        wrongDate,
        dateAnomaly,
        bazaNoDate,
        notFound,
        toleranceApplied,
        potentialMatches,
        bazaUnmatched,
        usedBaza
    };
}

// ============================================================================
// Parse Python CSV output
// ============================================================================

function parsePythonSummary(csvPath) {
    const content = fs.readFileSync(csvPath, 'utf-8');
    const lines = content.trim().split('\n');
    const headers = lines[0].split(',');

    const results = [];
    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(',');
        const row = {};
        for (let j = 0; j < headers.length; j++) {
            row[headers[j]] = values[j];
        }
        results.push(row);
    }
    return results;
}

/**
 * Parse Python details CSV with proper CSV handling for quoted fields
 */
function parsePythonDetails(csvPath) {
    if (!fs.existsSync(csvPath)) {
        console.warn(`WARNING: Python details file not found: ${csvPath}`);
        return [];
    }

    const content = fs.readFileSync(csvPath, 'utf-8');
    const lines = content.trim().split('\n');

    // Parse header
    const headers = lines[0].split(',');

    const results = [];
    for (let i = 1; i < lines.length; i++) {
        const row = parseCSVLine(lines[i], headers);
        if (row) {
            results.push(row);
        }
    }
    return results;
}

/**
 * Parse a single CSV line, handling quoted fields with commas
 */
function parseCSVLine(line, headers) {
    const values = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            values.push(current);
            current = '';
        } else {
            current += char;
        }
    }
    values.push(current); // Last field

    const row = {};
    for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = values[j] || '';
    }
    return row;
}

// ============================================================================
// Generate JS issues list (similar to Python details)
// ============================================================================

/**
 * Convert JS processing results to a flat issues list comparable to Python's details
 */
function generateJSIssuesList(jsResults) {
    const issues = [];

    for (const result of jsResults) {
        const month = result.month;

        // TOLERANCE issues
        for (const item of result.toleranceApplied) {
            issues.push({
                month,
                type: 'TOLERANCE',
                tovar_row: item.sale.row,
                tovar_name: item.sale.name,
                tovar_price_H: item.sale.price,
                tovar_sell_date: formatDate(item.sale.sellDate),
                baza_sheet: item.baza.sheet,
                baza_row: item.baza.row,
                baza_sell_date: formatDate(item.baza.sellDate),
                issue: item.tolerances.join(', ')
            });
        }

        // WRONG_DATE issues
        for (const item of result.wrongDate) {
            issues.push({
                month,
                type: 'WRONG_DATE',
                tovar_row: item.sale.row,
                tovar_name: item.sale.name,
                tovar_price_H: item.sale.price,
                tovar_sell_date: formatDate(item.tovarDate),
                baza_sheet: item.baza ? item.baza.sheet : '',
                baza_row: item.baza ? item.baza.row : '',
                baza_sell_date: item.bazaDate === 'ALL_USED' ? 'ALL_USED' : formatDate(item.baza ? item.baza.sellDate : null),
                issue: `tovar:${formatDate(item.tovarDate)} vs baza:${item.bazaDate === 'ALL_USED' ? 'ALL_USED' : formatDate(item.baza ? item.baza.sellDate : null)}`
            });
        }

        // DATE_ANOMALY issues
        for (const item of result.dateAnomaly) {
            issues.push({
                month,
                type: 'DATE_ANOMALY',
                tovar_row: item.sale.row,
                tovar_name: item.sale.name,
                tovar_price_H: item.sale.price,
                tovar_sell_date: formatDate(item.tovarDate),
                baza_sheet: item.baza.sheet,
                baza_row: item.baza.row,
                baza_sell_date: formatDate(item.bazaDate),
                issue: `${item.monthsApart} months apart: tovar:${formatDate(item.tovarDate)} vs baza:${formatDate(item.bazaDate)}`
            });
        }

        // NOT_FOUND issues
        for (const item of result.notFound) {
            issues.push({
                month,
                type: 'NOT_FOUND',
                tovar_row: item.row,
                tovar_name: item.name,
                tovar_price_H: item.price,
                tovar_sell_date: formatDate(item.sellDate),
                baza_sheet: '',
                baza_row: '',
                baza_sell_date: '',
                issue: 'Not found in baza'
            });
        }

        // BAZA_EXTRA issues
        for (const item of result.bazaUnmatched) {
            issues.push({
                month,
                type: 'BAZA_EXTRA',
                tovar_row: '',
                tovar_name: '',
                tovar_price_H: '',
                tovar_sell_date: '',
                baza_sheet: item.sheet,
                baza_row: item.row,
                baza_sell_date: formatDate(item.date),
                issue: `${item.name} H=${item.price}`
            });
        }

        // POTENTIAL issues
        for (const item of result.potentialMatches) {
            issues.push({
                month,
                type: 'POTENTIAL',
                tovar_row: item.sale.row,
                tovar_name: item.sale.name,
                tovar_price_H: item.sale.price,
                tovar_sell_date: formatDate(item.sale.sellDate),
                baza_sheet: item.baza.sheet,
                baza_row: item.baza.row,
                baza_sell_date: formatDate(item.baza.sellDate),
                issue: item.mismatches.join(', ')
            });
        }
    }

    return issues;
}

/**
 * Create a unique key for an issue to enable comparison
 */
function issueKey(issue) {
    // Normalize type name (Python uses БАЗА_EXTRA, JS uses BAZA_EXTRA)
    const type = issue.type.replace('БАЗА', 'BAZA');

    // For BAZA_EXTRA, use baza coordinates
    if (type === 'BAZA_EXTRA') {
        const sheet = issue.baza_sheet || issue['база_sheet'] || '';
        const row = issue.baza_row || issue['база_row'] || '';
        return `${issue.month}|${type}|${sheet}|${row}`;
    }
    // For others, use tovar coordinates
    return `${issue.month}|${type}|${issue.tovar_row}|${issue.tovar_name}`;
}

/**
 * Create a more relaxed key for fuzzy matching (just month/type/row)
 */
function issueKeyRelaxed(issue) {
    const type = issue.type.replace('БАЗА', 'BAZA');
    if (type === 'BAZA_EXTRA') {
        const sheet = issue.baza_sheet || issue['база_sheet'] || '';
        const row = issue.baza_row || issue['база_row'] || '';
        return `${issue.month}|${type}|${sheet}|${row}`;
    }
    return `${issue.month}|${type}|${issue.tovar_row}`;
}

/**
 * Normalize Python issue to match JS field names
 */
function normalizePythonIssue(pyIssue) {
    return {
        month: pyIssue.month,
        type: pyIssue.type.replace('БАЗА', 'BAZA'),
        tovar_row: pyIssue.tovar_row,
        tovar_name: pyIssue.tovar_name,
        tovar_price_H: pyIssue.tovar_price_H,
        tovar_sell_date: pyIssue.tovar_sell_date,
        baza_sheet: pyIssue['база_sheet'] || '',
        baza_row: pyIssue['база_row'] || '',
        baza_sell_date: pyIssue['база_sell_date'] || '',
        issue: pyIssue.issue
    };
}

/**
 * Compare Python and JS issues lists
 */
function compareIssuesLists(pyIssues, jsIssues) {
    const pyByKey = new Map();
    const jsByKey = new Map();

    // Normalize Python issues and group by key
    for (const issue of pyIssues) {
        const normalized = normalizePythonIssue(issue);
        const key = issueKey(normalized);
        if (!pyByKey.has(key)) {
            pyByKey.set(key, []);
        }
        pyByKey.get(key).push(normalized);
    }

    for (const issue of jsIssues) {
        const key = issueKey(issue);
        if (!jsByKey.has(key)) {
            jsByKey.set(key, []);
        }
        jsByKey.get(key).push(issue);
    }

    const onlyInPython = [];
    const onlyInJS = [];
    const matched = [];
    const valueDifferences = [];

    // Find items only in Python
    for (const [key, pyItems] of pyByKey) {
        const jsItems = jsByKey.get(key) || [];
        if (jsItems.length === 0) {
            onlyInPython.push(...pyItems);
        } else {
            // Compare values
            for (let i = 0; i < pyItems.length; i++) {
                if (i < jsItems.length) {
                    matched.push({ py: pyItems[i], js: jsItems[i] });
                    const diffs = compareIssueValues(pyItems[i], jsItems[i]);
                    if (diffs.length > 0) {
                        valueDifferences.push({ py: pyItems[i], js: jsItems[i], diffs });
                    }
                } else {
                    onlyInPython.push(pyItems[i]);
                }
            }
        }
    }

    // Find items only in JS
    for (const [key, jsItems] of jsByKey) {
        const pyItems = pyByKey.get(key) || [];
        if (pyItems.length === 0) {
            onlyInJS.push(...jsItems);
        } else if (jsItems.length > pyItems.length) {
            // Extra items in JS beyond what Python has
            for (let i = pyItems.length; i < jsItems.length; i++) {
                onlyInJS.push(jsItems[i]);
            }
        }
    }

    return { onlyInPython, onlyInJS, matched, valueDifferences };
}

/**
 * Compare values between a Python and JS issue
 */
function compareIssueValues(pyIssue, jsIssue) {
    const diffs = [];
    const fields = ['tovar_row', 'tovar_name', 'tovar_price_H', 'tovar_sell_date',
                    'baza_sheet', 'baza_row', 'baza_sell_date'];

    for (const field of fields) {
        const pyVal = String(pyIssue[field] || '');
        const jsVal = String(jsIssue[field] || '');

        // Normalize values for comparison
        const pyNorm = normalizeCompareValue(pyVal);
        const jsNorm = normalizeCompareValue(jsVal);

        if (pyNorm !== jsNorm) {
            diffs.push({ field, python: pyVal, javascript: jsVal });
        }
    }

    return diffs;
}

/**
 * Normalize a value for comparison
 */
function normalizeCompareValue(val) {
    if (!val || val === '' || val === 'undefined' || val === 'null') return '';
    // Remove trailing zeros from numbers
    const numMatch = val.match(/^(\d+\.\d*?)0+$/);
    if (numMatch) return numMatch[1].replace(/\.$/, '');
    return val;
}

/**
 * Print detailed debug report
 */
function printDebugReport(comparison, pyIssues, jsIssues) {
    console.log('\n' + '='.repeat(100));
    console.log('DETAILED DEBUG REPORT - Python vs JavaScript Issues Comparison');
    console.log('='.repeat(100));

    // Summary by type
    console.log('\n--- ISSUES COUNT BY TYPE ---');
    const pyByType = {};
    const jsByType = {};

    for (const issue of pyIssues) {
        const type = issue.type.replace('БАЗА', 'BAZA');
        pyByType[type] = (pyByType[type] || 0) + 1;
    }
    for (const issue of jsIssues) {
        jsByType[issue.type] = (jsByType[issue.type] || 0) + 1;
    }

    const allTypes = new Set([...Object.keys(pyByType), ...Object.keys(jsByType)]);
    console.log(padRight('Type', 20) + padRight('Python', 12) + padRight('JavaScript', 12) + 'Diff');
    console.log('-'.repeat(60));
    for (const type of allTypes) {
        const pyCount = pyByType[type] || 0;
        const jsCount = jsByType[type] || 0;
        const diff = jsCount - pyCount;
        console.log(
            padRight(type, 20) +
            padRight(String(pyCount), 12) +
            padRight(String(jsCount), 12) +
            (diff === 0 ? 'OK' : `${diff > 0 ? '+' : ''}${diff}`)
        );
    }

    // Items only in Python
    if (comparison.onlyInPython.length > 0) {
        console.log('\n' + '='.repeat(100));
        console.log(`ITEMS ONLY IN PYTHON (${comparison.onlyInPython.length} items)`);
        console.log('These items are in Python output but NOT found in JavaScript output');
        console.log('='.repeat(100));

        // Group by month and type
        const grouped = groupByMonthAndType(comparison.onlyInPython);
        for (const [monthType, items] of grouped) {
            console.log(`\n--- ${monthType} ---`);
            for (const item of items) {
                printIssueDetail(item, 'Python');
            }
        }
    }

    // Items only in JS
    if (comparison.onlyInJS.length > 0) {
        console.log('\n' + '='.repeat(100));
        console.log(`ITEMS ONLY IN JAVASCRIPT (${comparison.onlyInJS.length} items)`);
        console.log('These items are in JavaScript output but NOT found in Python output');
        console.log('='.repeat(100));

        const grouped = groupByMonthAndType(comparison.onlyInJS);
        for (const [monthType, items] of grouped) {
            console.log(`\n--- ${monthType} ---`);
            for (const item of items) {
                printIssueDetail(item, 'JavaScript');
            }
        }
    }

    // Value differences
    if (comparison.valueDifferences.length > 0) {
        console.log('\n' + '='.repeat(100));
        console.log(`VALUE DIFFERENCES (${comparison.valueDifferences.length} items)`);
        console.log('These items exist in both but have different values');
        console.log('='.repeat(100));

        for (const diff of comparison.valueDifferences) {
            console.log(`\n${diff.py.month} | ${diff.py.type} | Row ${diff.py.tovar_row}`);
            console.log(`  Name: ${diff.py.tovar_name}`);
            for (const d of diff.diffs) {
                console.log(`  ${d.field}:`);
                console.log(`    Python:     "${d.python}"`);
                console.log(`    JavaScript: "${d.javascript}"`);
            }
        }
    }

    // Summary
    console.log('\n' + '='.repeat(100));
    console.log('DEBUG SUMMARY');
    console.log('='.repeat(100));
    console.log(`Total Python issues:     ${pyIssues.length}`);
    console.log(`Total JavaScript issues: ${jsIssues.length}`);
    console.log(`Matched issues:          ${comparison.matched.length}`);
    console.log(`Only in Python:          ${comparison.onlyInPython.length}`);
    console.log(`Only in JavaScript:      ${comparison.onlyInJS.length}`);
    console.log(`With value differences:  ${comparison.valueDifferences.length}`);
}

/**
 * Group issues by month and type
 */
function groupByMonthAndType(issues) {
    const groups = new Map();
    for (const issue of issues) {
        const key = `${issue.month} / ${issue.type}`;
        if (!groups.has(key)) {
            groups.set(key, []);
        }
        groups.get(key).push(issue);
    }
    return groups;
}

/**
 * Print a single issue in detail
 */
function printIssueDetail(issue, source) {
    if (issue.type === 'BAZA_EXTRA') {
        console.log(`  [${source}] baza: ${issue.baza_sheet} row ${issue.baza_row}, date: ${issue.baza_sell_date}`);
        console.log(`    Issue: ${issue.issue}`);
    } else {
        console.log(`  [${source}] tovar row ${issue.tovar_row}: ${issue.tovar_name}`);
        console.log(`    Price H: ${issue.tovar_price_H}, Sell date: ${issue.tovar_sell_date}`);
        if (issue.baza_sheet) {
            console.log(`    baza: ${issue.baza_sheet} row ${issue.baza_row}, date: ${issue.baza_sell_date}`);
        }
        console.log(`    Issue: ${issue.issue}`);
    }
}

// ============================================================================
// Comparison and reporting
// ============================================================================

function compareResults(jsResults, pyResults) {
    const comparison = [];
    const fieldsToCompare = ['gap', 'matched', 'wrong_date', 'date_anomaly', 'база_no_date', 'not_found', 'база_extra'];

    for (const jsResult of jsResults) {
        const pyResult = pyResults.find(p => p.month === jsResult.month);
        if (!pyResult) {
            console.error(`WARNING: Month ${jsResult.month} not found in Python results`);
            continue;
        }

        const jsGap = jsResult.expectedDiff - jsResult.computedDiff;

        const monthComparison = {
            month: jsResult.month,
            differences: [],
            values: {
                js: {
                    gap: jsGap,
                    matched: jsResult.matched.length,
                    wrong_date: jsResult.wrongDate.length,
                    date_anomaly: jsResult.dateAnomaly.length,
                    'база_no_date': jsResult.bazaNoDate.length,
                    not_found: jsResult.notFound.length,
                    'база_extra': jsResult.bazaUnmatched.length
                },
                py: {
                    gap: parseFloat(pyResult.gap),
                    matched: parseInt(pyResult.matched),
                    wrong_date: parseInt(pyResult.wrong_date),
                    date_anomaly: parseInt(pyResult.date_anomaly),
                    'база_no_date': parseInt(pyResult['база_no_date']),
                    not_found: parseInt(pyResult.not_found),
                    'база_extra': parseInt(pyResult['база_extra'])
                }
            }
        };

        for (const field of fieldsToCompare) {
            const jsVal = monthComparison.values.js[field];
            const pyVal = monthComparison.values.py[field];

            // Use tolerance for floating point comparisons
            const tolerance = field === 'gap' ? 0.01 : 0;
            const diff = Math.abs(jsVal - pyVal);

            if (diff > tolerance) {
                monthComparison.differences.push({
                    field,
                    js: jsVal,
                    py: pyVal,
                    diff: diff
                });
            }
        }

        comparison.push(monthComparison);
    }

    return comparison;
}

function printComparisonReport(comparison) {
    console.log('\n' + '='.repeat(100));
    console.log('PYTHON vs JAVASCRIPT COMPARISON REPORT');
    console.log('='.repeat(100));

    let totalDifferences = 0;
    let monthsWithDifferences = 0;

    // Print header
    console.log('\n' + '-'.repeat(100));
    console.log(
        padRight('Month', 12) +
        padRight('Field', 15) +
        padRight('Python', 15) +
        padRight('JavaScript', 15) +
        padRight('Diff', 15) +
        'Status'
    );
    console.log('-'.repeat(100));

    for (const monthComp of comparison) {
        const fields = ['gap', 'matched', 'wrong_date', 'date_anomaly', 'база_no_date', 'not_found', 'база_extra'];
        let monthHasDiff = false;

        for (const field of fields) {
            const pyVal = monthComp.values.py[field];
            const jsVal = monthComp.values.js[field];
            const diff = monthComp.differences.find(d => d.field === field);

            const pyStr = field === 'gap' ? pyVal.toFixed(2) : String(pyVal);
            const jsStr = field === 'gap' ? jsVal.toFixed(2) : String(jsVal);

            let status = 'OK';
            let diffStr = '-';

            if (diff) {
                status = 'DIFF';
                diffStr = field === 'gap' ? diff.diff.toFixed(2) : String(diff.diff);
                totalDifferences++;
                monthHasDiff = true;
            }

            // Only show month name on first field row
            const monthStr = field === 'gap' ? monthComp.month : '';

            console.log(
                padRight(monthStr, 12) +
                padRight(field, 15) +
                padRight(pyStr, 15) +
                padRight(jsStr, 15) +
                padRight(diffStr, 15) +
                (status === 'OK' ? status : `*** ${status} ***`)
            );
        }

        if (monthHasDiff) {
            monthsWithDifferences++;
        }

        console.log('-'.repeat(100));
    }

    // Summary
    console.log('\n' + '='.repeat(100));
    console.log('SUMMARY');
    console.log('='.repeat(100));
    console.log(`Total months analyzed: ${comparison.length}`);
    console.log(`Months with differences: ${monthsWithDifferences}`);
    console.log(`Total field differences: ${totalDifferences}`);

    if (totalDifferences === 0) {
        console.log('\n*** ALL VALUES MATCH - JavaScript implementation matches Python ***\n');
        return true;
    } else {
        console.log('\n*** DIFFERENCES FOUND - JavaScript and Python implementations differ ***\n');

        // Print detailed differences
        console.log('DETAILED DIFFERENCES:');
        for (const monthComp of comparison) {
            if (monthComp.differences.length > 0) {
                console.log(`\n${monthComp.month}:`);
                for (const diff of monthComp.differences) {
                    console.log(`  ${diff.field}: Python=${diff.field === 'gap' ? diff.py.toFixed(2) : diff.py}, JS=${diff.field === 'gap' ? diff.js.toFixed(2) : diff.js} (diff: ${diff.field === 'gap' ? diff.diff.toFixed(2) : diff.diff})`);
                }
            }
        }
        return false;
    }
}

function padRight(str, len) {
    str = String(str);
    while (str.length < len) str += ' ';
    return str;
}

// ============================================================================
// Main entry point
// ============================================================================

function main() {
    console.log('Excel Comparison Tool - Python vs JavaScript Test');
    console.log('='.repeat(50));

    // Check files exist
    if (!fs.existsSync(FILE_TOVAR)) {
        console.error(`ERROR: tovar file not found: ${FILE_TOVAR}`);
        process.exit(1);
    }
    if (!fs.existsSync(FILE_BAZA)) {
        console.error(`ERROR: baza file not found: ${FILE_BAZA}`);
        process.exit(1);
    }
    if (!fs.existsSync(PYTHON_SUMMARY)) {
        console.error(`ERROR: Python summary not found: ${PYTHON_SUMMARY}`);
        console.error('Please run the Python script first: docker compose run --rm app python src/main.py');
        process.exit(1);
    }

    console.log('\nLoading Excel files...');

    // Load workbooks
    const tovarWorkbook = XLSX.readFile(FILE_TOVAR, { cellDates: true });
    const bazaWorkbook = XLSX.readFile(FILE_BAZA, { cellDates: true });

    console.log(`Tovar sheets: ${tovarWorkbook.SheetNames.length}`);
    console.log(`Baza sheets: ${bazaWorkbook.SheetNames.length}`);

    // Load baza sheets
    const bazaSheets = loadBazaSheets(bazaWorkbook);
    console.log(`Loaded ${Object.keys(bazaSheets).length} product sheets from Baza`);

    // Process all months
    console.log('\nProcessing months...');
    const jsResults = [];
    const availableSheets = tovarWorkbook.SheetNames;

    for (const monthConfig of MONTHS) {
        if (!availableSheets.includes(monthConfig.sheetName)) {
            console.log(`  Skipping ${monthConfig.monthName} (sheet not found)`);
            continue;
        }

        process.stdout.write(`  Processing ${monthConfig.monthName}...`);
        const result = processMonth(monthConfig, tovarWorkbook, bazaSheets);

        if (result) {
            jsResults.push(result);
            const gap = result.expectedDiff - result.computedDiff;
            console.log(` OK (gap: ${gap.toFixed(2)}, matched: ${result.matched.length})`);
        } else {
            console.log(' No data');
        }
    }

    // Load Python results
    console.log('\nLoading Python results...');
    const pyResults = parsePythonSummary(PYTHON_SUMMARY);
    console.log(`Loaded ${pyResults.length} months from Python summary`);

    // Compare and report
    const comparison = compareResults(jsResults, pyResults);
    const success = printComparisonReport(comparison);

    // Debug mode: detailed item-by-item comparison
    if (DEBUG_MODE) {
        console.log('\n' + '='.repeat(100));
        console.log('DEBUG MODE ENABLED - Loading detailed comparison');
        console.log('='.repeat(100));

        // Load Python details
        const pyDetails = parsePythonDetails(PYTHON_DETAILS);
        console.log(`Loaded ${pyDetails.length} detailed issues from Python`);

        // Generate JS issues list
        const jsIssues = generateJSIssuesList(jsResults);
        console.log(`Generated ${jsIssues.length} detailed issues from JavaScript`);

        // Compare issues lists
        const issuesComparison = compareIssuesLists(pyDetails, jsIssues);

        // Print detailed debug report
        printDebugReport(issuesComparison, pyDetails, jsIssues);
    }

    process.exit(success ? 0 : 1);
}

// Run
main();

#!/usr/bin/env node
/**
 * Test script to compare JavaScript output with Python output
 * for the Excel comparison tool.
 *
 * Run with: docker run --rm -v "$(pwd)":/app -w /app node:20 node test/compare.js
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Configuration
const BASE_PATH = '/app';
const FILE_TOVAR = path.join(BASE_PATH, 'tovar2025_PRN.xlsx');
const FILE_BAZA = path.join(BASE_PATH, 'База товара_PRN.xls');
const PYTHON_SUMMARY = path.join(BASE_PATH, 'output', 'comparison_summary.csv');

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

function isNaN_or_null(val) {
    return val === null || val === undefined || (typeof val === 'number' && isNaN(val));
}

function isEmptyEquivalent(val) {
    if (val === null || val === undefined) return true;
    if (typeof val === 'number' && isNaN(val)) return true;
    if (typeof val === 'string' && ['', '-', '\u2014', '\u2013'].includes(val.trim())) return true;
    return false;
}

function normalizeValue(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') {
        if (isNaN(v)) return null;
        return Math.round(v * 10000) / 10000;
    }
    return v;
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
    // Check exact match first
    let exactMatch = true;
    for (let i = 0; i < 8; i++) {
        const s = normalizeValue(saleKey[i]);
        const r = normalizeValue(rowKey[i]);
        if (s !== r && !(s === null && r === null)) {
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

        if (sn === rn) continue;
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
        if (sn === rn) continue;
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
    const saleKey = normalizeTuple(sale.colsAH);
    const tovarDate = sale.sellDate;
    const matches = [];

    for (const [sheetName, data] of Object.entries(bazaSheets)) {
        for (let idx = 0; idx < data.length; idx++) {
            const row = data[idx];
            if (row[0] !== name) continue;

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
    const sellDate = sale.sellDate;
    const saleKey = normalizeTuple(sale.colsAH);
    const potentials = [];

    for (const [sheetName, data] of Object.entries(bazaSheets)) {
        for (let idx = 0; idx < data.length; idx++) {
            const row = data[idx];
            if (row[0] !== name) continue;

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

    process.exit(success ? 0 : 1);
}

// Run
main();

/**
 * formulaTranslator.js
 *
 * Translates human-readable header references in formulas to A1 notation
 * that HyperFormula can evaluate.
 *
 * SYNTAX:
 *   ColumnHeader_RowHeader  → cell reference  e.g. Revenue_Q1 → A1
 *   ColumnHeader_RowNumber  → cell reference  e.g. Revenue_1  → A1
 *   ColumnLetter_RowHeader  → cell reference  e.g. A_Q1       → A1
 *
 * RULES:
 *   - Separator is always underscore (_)
 *   - Works with any HF formula: =SUM, =AVERAGE, =IF, =VLOOKUP etc.
 *   - Ranges work:  Revenue_Q1:Revenue_Q4 → A1:A4
 *   - Mixed refs:   Revenue_Q1 + B2 → A1 + B2 (A1 notation still works)
 *   - Case sensitive — header names must match exactly as configured
 *   - If no match found, token is left unchanged (safe fallback)
 *
 * EXAMPLES:
 *   columns:   [{ header: "Revenue" }, { header: "Cost" }, { header: "Profit" }]
 *   rowLabels: ["Q1", "Q2", "Q3", "Q4"]
 *
 *   =SUM(Revenue_Q1, Revenue_Q2)        → =SUM(A1, A2)
 *   =AVERAGE(Revenue_Q1:Revenue_Q4)     → =AVERAGE(A1:A4)
 *   =Cost_Q3 - Revenue_Q3              → =B3 - A3
 *   =IF(Revenue_Q1>1000,"Good","Bad")  → =IF(A1>1000,"Good","Bad")
 *   =SUM(Revenue_1:Revenue_4)          → =SUM(A1:A4)
 *   =A_Q1 + B_Q2                       → =A1 + B2
 */

// ── Column index → spreadsheet letter(s) ─────────────────────────────────────
// 0 → A, 1 → B, 25 → Z, 26 → AA, 27 → AB ...
export function columnIndexToLetter(index) {
    let letter = "";
    let n = index;
    while (n >= 0) {
        letter = String.fromCharCode(65 + (n % 26)) + letter;
        n = Math.floor(n / 26) - 1;
    }
    return letter;
}

// ── Build reference map ───────────────────────────────────────────────────────
// Returns a Map of "token" → "A1ref"
// Called once per render when columns or rowLabels change.
//
// Generates all valid token combinations:
//   ColumnHeader_RowHeader  e.g. Revenue_Q1
//   ColumnHeader_RowNumber  e.g. Revenue_1
//   ColumnLetter_RowHeader  e.g. A_Q1
//
export function buildHeaderRefMap(columns, rowLabels) {
    const map = new Map();

    const cols = Array.isArray(columns)   ? columns   : [];
    const rows = Array.isArray(rowLabels) ? rowLabels : [];

    const colCount = cols.length;
    const rowCount = rows.length;

    cols.forEach((col, colIndex) => {
        const colHeader = col?.header ? String(col.header).trim() : "";
        const colLetter = columnIndexToLetter(colIndex);

        // ── ColumnHeader_RowHeader  e.g. Revenue_Q1 → A1
        if (colHeader && rowCount > 0) {
            rows.forEach((rowLabel, rowIndex) => {
                const rowLabelStr = rowLabel ? String(rowLabel).trim() : "";
                if (!rowLabelStr) return;

                const token = `${colHeader}_${rowLabelStr}`;
                const ref   = `${colLetter}${rowIndex + 1}`;
                map.set(token, ref);
            });
        }

        // ── ColumnHeader_RowNumber  e.g. Revenue_1 → A1
        // Always available — user can mix header cols with row numbers
        if (colHeader) {
            for (let rowIndex = 0; rowIndex < Math.max(rowCount, 100); rowIndex++) {
                const token = `${colHeader}_${rowIndex + 1}`;
                const ref   = `${colLetter}${rowIndex + 1}`;
                // Only add if not already covered by RowHeader variant
                if (!map.has(token)) {
                    map.set(token, ref);
                }
            }
        }
    });

    // ── ColumnLetter_RowHeader  e.g. A_Q1 → A1
    // Available when row headers configured but columns use default letters
    if (rowCount > 0) {
        rows.forEach((rowLabel, rowIndex) => {
            const rowLabelStr = rowLabel ? String(rowLabel).trim() : "";
            if (!rowLabelStr) return;

            // Support up to 26*26 columns (ZZ)
            for (let colIndex = 0; colIndex < 702; colIndex++) {
                const colLetter = columnIndexToLetter(colIndex);
                const token     = `${colLetter}_${rowLabelStr}`;
                const ref       = `${colLetter}${rowIndex + 1}`;
                if (!map.has(token)) {
                    map.set(token, ref);
                }
            }
        });
    }

    return map;
}

// ── Translate a formula string ────────────────────────────────────────────────
// Replaces all header reference tokens with A1 notation.
// Leaves everything else (operators, function names, strings, numbers) intact.
//
// Strategy:
//   1. Only process strings starting with "="
//   2. Tokenise: split on formula delimiters (, ; : ( ) + - * / ! space)
//      but keep delimiters in result so we can reconstruct
//   3. For each token, check if it exists in the map
//   4. If yes, replace with A1 ref
//   5. Reconstruct the formula
//
export function translateFormula(formula, refMap) {
    if (!formula || typeof formula !== "string") return formula;
    if (!formula.startsWith("=")) return formula;
    if (!refMap || refMap.size === 0) return formula;

    // Regex: split on formula structural characters but keep them
    // Captures: ( ) , ; : + - * / space ! = < > " '
    // We keep quoted strings intact (don't translate inside "...")
    const result = formula.replace(
        /("(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|[A-Za-z_][A-Za-z0-9_]*)/g,
        (match) => {
            // Skip quoted strings
            if (match.startsWith('"') || match.startsWith("'")) return match;
            // Check if this token is in the map
            const translated = refMap.get(match);
            return translated !== undefined ? translated : match;
        }
    );

    return result;
}

// ── Convenience: translate if formula, passthrough otherwise ─────────────────
export function maybeTranslate(value, refMap) {
    if (typeof value === "string" && value.startsWith("=")) {
        return translateFormula(value, refMap);
    }
    return value;
}
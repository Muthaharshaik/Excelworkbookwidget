/**
 * dataService.js
 *
 * NEW ARCHITECTURE: One sheet per widget instance.
 * sheetJson contains data for a single Spreadsheet entity.
 *
 * JSON shape stored in Spreadsheet.sheetJson:
 * {
 *   "data":        [[row data]],
 *   "columns":     [{ key, header, type, width, source, format, readOnly }],
 *   "rowLabels":   ["Q1", "Q2", "", "Revenue", ...],
 *   "cellMeta":    { "r0_c0": { bold, italic, fontColor, bgColor, align } },
 *   "colWidths":   [120, 80, 150],
 *   "rowHeights":  [23, 23, ...],
 *   "mergedCells": [{ row, col, rowspan, colspan }]
 * }
 *
 * allSheetsJson shape stored in Workbook.allSheetsJson:
 * [
 *   { "sheetId": "1", "sheetName": "Revenue",  "data": [[1,2],[3,4]] },
 *   { "sheetId": "2", "sheetName": "Expenses", "data": [[5,6],[7,8]] }
 * ]
 * Only sheetId, sheetName, data are needed — HF doesn't need formatting.
 */

import { MIN_COLS } from "../utils/constants";

// ─────────────────────────────────────────────────────────────────────────────
//  PARSE  (Mendix sheetJson string → widget state object)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * parseSheetJson
 * Parses the JSON string from Spreadsheet.sheetJson into widget state.
 *
 * @param {string} jsonString   - value of Spreadsheet.sheetJson attribute
 * @param {number} rowCount     - configured row count (from widget prop)
 * @returns {object}            - { data, columns, rowLabels, cellMeta, colWidths, rowHeights, mergedCells }
 */
export function parseSheetJson(jsonString, rowCount = 50) {
    const empty = buildEmptySheetData(rowCount);

    if (!jsonString || typeof jsonString !== "string" || jsonString.trim() === "") {
        return empty;
    }

    let raw;
    try {
        raw = JSON.parse(jsonString);
    } catch (e) {
        console.error("[ExcelWidget] Failed to parse sheetJson:", e.message);
        return empty;
    }

    if (typeof raw !== "object" || Array.isArray(raw)) {
        console.error("[ExcelWidget] sheetJson must be a JSON object, not array or primitive.");
        return empty;
    }

    return {
        data:        normaliseData(raw.data, rowCount),
        columns:     Array.isArray(raw.columns)     ? raw.columns                          : [],
        rowLabels:   Array.isArray(raw.rowLabels)   ? raw.rowLabels.map(l => String(l ?? "")) : [],
        cellMeta:    (raw.cellMeta && typeof raw.cellMeta === "object") ? raw.cellMeta     : {},
        colWidths:   Array.isArray(raw.colWidths)   ? raw.colWidths                        : [],
        rowHeights:  Array.isArray(raw.rowHeights)  ? raw.rowHeights                       : [],
        mergedCells: Array.isArray(raw.mergedCells) ? raw.mergedCells                      : [],
    };
}

// ─────────────────────────────────────────────────────────────────────────────
//  PARSE ALL SHEETS  (Workbook.allSheetsJson → array of sheet summaries)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * parseAllSheetsJson
 *
 * Parses the combined all-sheets JSON from Workbook.allSheetsJson.
 * Used by useHyperformula to register ALL sheets into the HF engine
 * so cross-sheet references like =Revenue!A1 resolve correctly.
 *
 * Only extracts sheetId, sheetName, data — HF doesn't need formatting.
 *
 * @param {string} jsonString  - value of Workbook.allSheetsJson attribute
 * @returns {Array<{ sheetId: string, sheetName: string, data: any[][] }>}
 *          Returns empty array if invalid — widget falls back to
 *          single-sheet mode gracefully.
 */
export function parseAllSheetsJson(jsonString) {
    // Not provided — single-sheet mode fallback
    if (!jsonString || typeof jsonString !== "string" || jsonString.trim() === "") {
        return [];
    }

    let raw;
    try {
        raw = JSON.parse(jsonString);
    } catch (e) {
        console.error("[ExcelWidget] Failed to parse allSheetsJson:", e.message);
        return [];
    }

    // Must be an array
    if (!Array.isArray(raw)) {
        console.error("[ExcelWidget] allSheetsJson must be a JSON array.");
        return [];
    }

    // Validate and normalise each sheet entry
    return raw
        .filter(entry => {
            if (!entry || typeof entry !== "object") return false;
            if (!entry.sheetName || typeof entry.sheetName !== "string") return false;
            return true;
        })
        .map(entry => ({
            sheetId:   String(entry.sheetId   ?? ""),
            sheetName: String(entry.sheetName ?? ""),
            // Normalise data — must be array of arrays
            data: Array.isArray(entry.data)
                ? entry.data.map(row => Array.isArray(row) ? row : [])
                : [[]],
        }));
}

// ─────────────────────────────────────────────────────────────────────────────
//  SERIALIZE  (widget state → JSON string for Mendix setValue)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * serializeSheet
 * Converts widget state back to JSON string for Spreadsheet.sheetJson.
 *
 * @param {object} sheetData
 * @returns {string}
 */
export function serializeSheet(sheetData) {
    try {
        return JSON.stringify({
            data:        trimData(sheetData.data || []),
            columns:     sheetData.columns     || [],
            rowLabels:   trimRowLabels(sheetData.rowLabels || []),
            cellMeta:    sheetData.cellMeta    || {},
            colWidths:   sheetData.colWidths   || [],
            rowHeights:  sheetData.rowHeights  || [],
            mergedCells: sheetData.mergedCells || [],
        });
    } catch (e) {
        console.error("[ExcelWidget] Failed to serialize sheet:", e.message);
        return "{}";
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  INTERNAL HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function buildEmptySheetData(rowCount) {
    return {
        data:        Array.from({ length: rowCount }, () => Array(MIN_COLS).fill(null)),
        columns:     [],
        rowLabels:   [],
        cellMeta:    {},
        colWidths:   [],
        rowHeights:  [],
        mergedCells: [],
    };
}

function normaliseData(raw, rowCount) {
    if (!Array.isArray(raw) || raw.length === 0) {
        return Array.from({ length: rowCount }, () => Array(MIN_COLS).fill(null));
    }

    const maxCols = Math.max(MIN_COLS, ...raw.map(r => Array.isArray(r) ? r.length : 0));

    const paddedRows = raw.map(row => {
        if (!Array.isArray(row)) return Array(maxCols).fill(null);
        if (row.length >= maxCols) return row;
        return [...row, ...Array(maxCols - row.length).fill(null)];
    });

    while (paddedRows.length < rowCount) {
        paddedRows.push(Array(maxCols).fill(null));
    }

    return paddedRows;
}

function trimData(data) {
    if (!Array.isArray(data) || data.length === 0) return [[]];

    let rows = data.map(row => [...(Array.isArray(row) ? row : [])]);

    while (rows.length > 1 && rows[rows.length - 1].every(isEmpty)) {
        rows.pop();
    }

    let maxUsedCol = 0;
    rows.forEach(row => {
        for (let c = row.length - 1; c >= 0; c--) {
            if (!isEmpty(row[c])) {
                if (c > maxUsedCol) maxUsedCol = c;
                break;
            }
        }
    });

    return rows.map(row => row.slice(0, maxUsedCol + 1));
}

function trimRowLabels(labels) {
    if (!Array.isArray(labels) || labels.length === 0) return [];
    const arr = [...labels];
    while (arr.length > 0 && arr[arr.length - 1] === "") {
        arr.pop();
    }
    return arr;
}

function isEmpty(v) {
    return v === null || v === undefined || v === "";
}
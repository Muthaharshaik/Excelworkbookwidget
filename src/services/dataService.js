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
 *   "cellMeta":    { "r0_c0": { bold, italic, fontColor, bgColor, align } },
 *   "colWidths":   [120, 80, 150],
 *   "rowHeights":  [23, 23, ...],
 *   "mergedCells": [{ row, col, rowspan, colspan }]
 * }
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
 * @returns {object}            - { data, columns, cellMeta, colWidths, rowHeights, mergedCells }
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
        columns:     Array.isArray(raw.columns)     ? raw.columns     : [],
        cellMeta:    (raw.cellMeta && typeof raw.cellMeta === "object") ? raw.cellMeta : {},
        colWidths:   Array.isArray(raw.colWidths)   ? raw.colWidths   : [],
        rowHeights:  Array.isArray(raw.rowHeights)  ? raw.rowHeights  : [],
        mergedCells: Array.isArray(raw.mergedCells) ? raw.mergedCells : [],
    };
}

// ─────────────────────────────────────────────────────────────────────────────
//  SERIALIZE  (widget state → JSON string for Mendix setValue)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * serializeSheet
 * Converts widget state back to JSON string for Spreadsheet.sheetJson.
 * Trims empty trailing rows/cols to keep JSON small.
 *
 * @param {object} sheetData  - { data, columns, cellMeta, colWidths, rowHeights, mergedCells }
 * @returns {string}
 */
export function serializeSheet(sheetData) {
    try {
        return JSON.stringify({
            data:        trimData(sheetData.data || []),
            columns:     sheetData.columns     || [],
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

    // Pad to rowCount if needed
    while (paddedRows.length < rowCount) {
        paddedRows.push(Array(maxCols).fill(null));
    }

    return paddedRows;
}

function trimData(data) {
    if (!Array.isArray(data) || data.length === 0) return [[]];

    let rows = data.map(row => [...(Array.isArray(row) ? row : [])]);

    // Remove empty trailing rows
    while (rows.length > 1 && rows[rows.length - 1].every(isEmpty)) {
        rows.pop();
    }

    // Find last used column
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

function isEmpty(v) {
    return v === null || v === undefined || v === "";
}
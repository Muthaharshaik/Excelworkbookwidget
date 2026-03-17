/**
 * dataService.js — FIXED
 *
 * CHANGE: parseAllSheetsJson is now more robust.
 *
 * Your allSheetsJson sample contains entries where "data" is a string
 * (a partial JSON fragment) instead of a 2D array:
 *   { "sheetId": "Sheey6", "data": "\"sheetId\":\"null\"..." }
 *
 * When this string is passed to HyperFormula's setSheetContent(), HF
 * throws internally. The sheet gets registered with no content, but
 * subsequent cross-sheet formulas referencing it return INVALID_REFERENCE
 * or bleed values from the previously-active sheet.
 *
 * FIX: normaliseSheetData() validates and coerces all data field values.
 * Strings that look like partial JSON are parsed if possible; otherwise
 * the sheet gets an empty [[]] data. This prevents HF from receiving
 * invalid input while still registering the sheet (so cross-references
 * to it resolve to empty rather than erroring).
 */

import { MIN_COLS } from "../utils/constants";

// ─────────────────────────────────────────────────────────────────────────────
//  PARSE  (Mendix sheetJson string → widget state object)
// ─────────────────────────────────────────────────────────────────────────────

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
        _sheetId:    typeof raw.sheetId === "string" ? raw.sheetId : null,
        data:        normaliseData(raw.data, rowCount),
        columns:     Array.isArray(raw.columns)     ? raw.columns                             : [],
        rowLabels:   Array.isArray(raw.rowLabels)   ? raw.rowLabels.map(l => String(l ?? "")) : [],
        cellMeta:    (raw.cellMeta && typeof raw.cellMeta === "object") ? raw.cellMeta        : {},
        colWidths:   Array.isArray(raw.colWidths)   ? raw.colWidths                           : [],
        rowHeights:  Array.isArray(raw.rowHeights)  ? raw.rowHeights                          : [],
        mergedCells: Array.isArray(raw.mergedCells) ? raw.mergedCells                         : [],
    };
}

// ─────────────────────────────────────────────────────────────────────────────
//  PARSE ALL SHEETS  (Workbook.allSheetsJson → array of sheet summaries)
//
//  FIXED: normaliseSheetData handles the case where "data" is a string
//  (malformed JSON fragment) rather than a 2D array. This was causing HF
//  to receive invalid setSheetContent calls and silently bleed formula values.
// ─────────────────────────────────────────────────────────────────────────────

export function parseAllSheetsJson(jsonString) {
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

    if (!Array.isArray(raw)) {
        console.error("[ExcelWidget] allSheetsJson must be a JSON array.");
        return [];
    }

    return raw
        .filter(entry => {
            if (!entry || typeof entry !== "object") return false;
            if (!entry.sheetName || typeof entry.sheetName !== "string") return false;
            return true;
        })
        .map(entry => ({
            sheetId:   String(entry.sheetId   ?? ""),
            sheetName: String(entry.sheetName ?? ""),
            data:      normaliseSheetData(entry.data),
        }));
}

/**
 * normaliseSheetData
 *
 * Safely converts the "data" field from allSheetsJson entries into a
 * valid 2D array for HyperFormula's setSheetContent.
 *
 * Handles:
 *   - Proper 2D arrays → returned as-is (with empty row fallback)
 *   - null / undefined / empty string → [[]]
 *   - String values (malformed JSON fragments from broken Mendix Java actions)
 *     → attempt JSON.parse; if result is a 2D array use it, else use [[]]
 *   - Objects → [[]]  (should not happen but defensive)
 */
function normaliseSheetData(data) {
    if (Array.isArray(data)) {
        if (data.length === 0) return [[]];
        return data.map(row => Array.isArray(row) ? row : []);
    }

    if (typeof data === "string" && data.trim().length > 0) {
        // Some Mendix Java actions may accidentally serialize sheetJson into
        // this field instead of the actual 2D data array.
        // Attempt recovery by parsing the string as JSON.
        try {
            const parsed = JSON.parse(data);
            if (Array.isArray(parsed)) {
                if (parsed.length === 0) return [[]];
                return parsed.map(row => Array.isArray(row) ? row : []);
            }
        } catch {
            // Not valid JSON — just return empty
        }
        console.warn("[ExcelWidget] normaliseSheetData: data field is a string but not a valid 2D array JSON. Defaulting to empty sheet.");
    }

    return [[]];
}

// ─────────────────────────────────────────────────────────────────────────────
//  SERIALIZE  (widget state → JSON string for Mendix setValue)
// ─────────────────────────────────────────────────────────────────────────────

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
        _sheetId:    null,
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
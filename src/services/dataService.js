/**
 * dataService.js
 *
 * KEY CHANGE: serializeSheet now includes a `metadata` field.
 * This gives Mendix everything needed for formula builder dropdowns:
 *
 * metadata: {
 *   columnHeaders: ["Name", "Age", "C", "D"...]   // custom if configured, else A/B/C
 *   rowLabels:     ["Sreenadh", "Rutesh", "3"...]  // custom if configured, else 1/2/3
 *   columnCount:   26,
 *   rowCount:      3
 * }
 *
 * Mendix developer reads Spreadsheet.sheetJson → parses metadata
 * → populates From Column / To Column / Row dropdowns in formula builder.
 *
 * Widget reads metadata back but ignores it (doesn't affect rendering).
 */

import { MIN_COLS } from "../utils/constants";

// ─────────────────────────────────────────────────────────────────────────────
//  PARSE
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
        console.error("[ExcelWidget] sheetJson must be a JSON object.");
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
        // ── NEW: locked cells from formula destination ──
        lockedCells: Array.isArray(raw.lockedCells) ? raw.lockedCells                         : [],
        // metadata is read-only — widget doesn't use it for rendering
    };
}

// ─────────────────────────────────────────────────────────────────────────────
//  PARSE ALL SHEETS
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
        .filter(entry => entry && typeof entry === "object" && entry.sheetName)
        .map(entry => ({
            sheetId:   String(entry.sheetId   ?? ""),
            sheetName: String(entry.sheetName ?? ""),
            data:      normaliseSheetData(entry.data),
        }));
}

function normaliseSheetData(data) {
    if (Array.isArray(data)) {
        if (data.length === 0) return [[]];
        return data.map(row => Array.isArray(row) ? row : []);
    }
    if (typeof data === "string" && data.trim().length > 0) {
        try {
            const parsed = JSON.parse(data);
            if (Array.isArray(parsed)) {
                return parsed.length === 0 ? [[]] : parsed.map(row => Array.isArray(row) ? row : []);
            }
        } catch {}
    }
    return [[]];
}

// ─────────────────────────────────────────────────────────────────────────────
//  SERIALIZE — includes metadata for Mendix formula builder
// ─────────────────────────────────────────────────────────────────────────────

export function serializeSheet(sheetData) {
    try {
        const data        = trimData(sheetData.data || []);
        const columns     = sheetData.columns   || [];
        const rowLabels   = trimRowLabels(sheetData.rowLabels || []);

        // Build metadata for Mendix formula builder dropdowns
        // columnHeaders: custom names if configured, else A/B/C for all columns
        // rowLabels: custom labels if configured, else 1/2/3 for all rows
        const dataColCount = data.length > 0 ? Math.max(...data.map(r => Array.isArray(r) ? r.length : 0), 1) : MIN_COLS;
        const dataRowCount = data.length || 0;

        const fullColCount = columns.length > 0 ? columns.length : Math.max(dataColCount, MIN_COLS);
        const fullRowCount = rowLabels.length > 0 ? rowLabels.length : Math.max(dataRowCount, 50);

        const columnHeaders = Array.from({ length: fullColCount }, (_, i) =>
            columns[i]?.header || colIndexToLetter(i)
        );

        const rowLabelsMeta = Array.from({ length: fullRowCount }, (_, i) =>
            rowLabels[i] || String(i + 1)
        );

        return JSON.stringify({
            data,
            columns,
            rowLabels,
            cellMeta:    sheetData.cellMeta    || {},
            colWidths:   sheetData.colWidths   || [],
            rowHeights:  sheetData.rowHeights  || [],
            mergedCells: sheetData.mergedCells || [],
            // ── NEW: preserve lockedCells so they survive widget save cycles ──
            lockedCells: sheetData.lockedCells || [],
            metadata: {
                columnHeaders,
                rowLabels: rowLabelsMeta,
                columnCount: fullColCount,
                rowCount:    fullRowCount,
            },
        });
    } catch (e) {
        console.error("[ExcelWidget] Failed to serialize sheet:", e.message);
        return "{}";
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function colIndexToLetter(index) {
    let result = "";
    let n = index;
    while (n >= 0) {
        result = String.fromCharCode((n % 26) + 65) + result;
        n = Math.floor(n / 26) - 1;
    }
    return result;
}

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
        lockedCells: [],
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
    while (arr.length > 0 && arr[arr.length - 1] === "") arr.pop();
    return arr;
}

function isEmpty(v) {
    return v === null || v === undefined || v === "";
}
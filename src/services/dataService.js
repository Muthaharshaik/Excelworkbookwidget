/**
 * dataService.js
 *
 * Single source of truth for all data transformation between
 * Mendix (JSON strings) and the widget's internal sheet objects.
 *
 * TWO RESPONSIBILITIES:
 *   1. PARSE  — take the raw sheetsJson string from Mendix and
 *               produce a clean, validated array of sheet objects
 *               the widget can trust.
 *
 *   2. SERIALIZE — take a sheet object after the user has edited it
 *                  and produce the JSON string Mendix needs to commit.
 *
 * NOTHING in this file touches Mendix actions or React state.
 * It is pure data transformation — easy to unit test in isolation.
 */

import { DEFAULT_SHEET, EMPTY_SHEET_DATA, MIN_ROWS, MIN_COLS } from "../utils/constants";

// ─────────────────────────────────────────────────────────────────────────────
//  PARSE  (Mendix → Widget)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * parseSheets
 *
 * Takes the raw string from the Mendix `sheetsJson` attribute and returns
 * a clean array of sheet objects the widget can trust.
 *
 * Handles:
 *  - JSON parse errors (returns empty array, logs warning)
 *  - Missing fields (fills in defaults from DEFAULT_SHEET)
 *  - Empty / null data arrays (fills with EMPTY_SHEET_DATA())
 *  - Ensures minimum grid size (MIN_ROWS × MIN_COLS)
 *  - Sorts sheets by orderIndex
 *
 * @param   {string} jsonString  - raw value of Mendix sheetsJson attribute
 * @returns {SheetObject[]}      - validated, sorted sheet array
 *
 * SheetObject shape:
 * {
 *   sheetId:     string,
 *   sheetName:   string,
 *   orderIndex:  number,
 *   isEditable:  boolean,
 *   data:        (string|number|null)[][],  ← 2D array, safe to pass to HotTable
 *   cellMeta:    object,                    ← { "row,col": { bold, italic, ... } }
 *   colWidths:   number[],
 *   rowHeights:  number[],
 *   mergedCells: MergedCellConfig[],
 * }
 */
export function parseSheets(jsonString) {
    // ── Guard: empty / missing prop ───────────────────────────────
    if (!jsonString || typeof jsonString !== "string" || jsonString.trim() === "") {
        console.warn("[ExcelWidget] sheetsJson is empty or not a string.");
        return [];
    }

    // ── Parse JSON ────────────────────────────────────────────────
    let rawSheets;
    try {
        rawSheets = JSON.parse(jsonString);
    } catch (e) {
        console.error("[ExcelWidget] Failed to parse sheetsJson:", e.message);
        console.error("[ExcelWidget] Raw value:", jsonString.slice(0, 200));
        return [];
    }

    // ── Validate it is an array ───────────────────────────────────
    if (!Array.isArray(rawSheets)) {
        console.error("[ExcelWidget] sheetsJson must be a JSON array of sheet objects.");
        return [];
    }

    // ── Normalise each sheet ──────────────────────────────────────
    const normalised = rawSheets.map((raw, index) => normaliseSheet(raw, index));

    // ── Sort by orderIndex ────────────────────────────────────────
    normalised.sort((a, b) => a.orderIndex - b.orderIndex);

    return normalised;
}

/**
 * normaliseSheet
 *
 * Takes a single raw sheet object from the parsed JSON and fills in
 * any missing fields with safe defaults. Never mutates the input.
 *
 * @param   {object} raw    - one item from the parsed JSON array
 * @param   {number} index  - fallback orderIndex if missing
 * @returns {SheetObject}
 */
function normaliseSheet(raw, index) {
    if (!raw || typeof raw !== "object") {
        console.warn(`[ExcelWidget] Sheet at index ${index} is not an object, using empty sheet.`);
        return buildEmptySheet(`sheet-${index}`, `Sheet${index + 1}`, index);
    }

    return {
        // Identity
        sheetId:    raw.sheetId    ?? `sheet-${index}`,
        sheetName:  raw.sheetName  ?? `Sheet${index + 1}`,
        orderIndex: typeof raw.orderIndex === "number" ? raw.orderIndex : index,

        // Permission — default SAFE (read-only) if not explicitly granted
        isEditable: raw.isEditable === true,

        // Grid data — ensure minimum size
        data: normaliseData(raw.data),

        // Metadata
        cellMeta:    raw.cellMeta    && typeof raw.cellMeta === "object"  ? raw.cellMeta    : {},
        colWidths:   Array.isArray(raw.colWidths)                         ? raw.colWidths   : [],
        rowHeights:  Array.isArray(raw.rowHeights)                        ? raw.rowHeights  : [],
        mergedCells: Array.isArray(raw.mergedCells)                       ? raw.mergedCells : [],
    };
}

/**
 * normaliseData
 *
 * Ensures the data array:
 *  - Is a 2D array
 *  - Has at least MIN_ROWS rows
 *  - Each row has at least MIN_COLS columns
 *
 * Handsontable can crash with jagged arrays, so we pad here.
 *
 * @param   {any}       raw  - raw data from JSON
 * @returns {(any)[][]}      - safe 2D array
 */
function normaliseData(raw) {
    // No data → return empty grid
    if (!Array.isArray(raw) || raw.length === 0) {
        return EMPTY_SHEET_DATA();
    }

    // Determine the target column count:
    // max of MIN_COLS and widest row in existing data
    const maxCols = Math.max(MIN_COLS, ...raw.map(r => (Array.isArray(r) ? r.length : 0)));

    // Ensure each row is an array padded to maxCols
    const paddedRows = raw.map(row => {
        if (!Array.isArray(row)) return Array(maxCols).fill(null);
        if (row.length >= maxCols) return row;
        return [...row, ...Array(maxCols - row.length).fill(null)];
    });

    // Ensure minimum row count
    while (paddedRows.length < MIN_ROWS) {
        paddedRows.push(Array(maxCols).fill(null));
    }

    return paddedRows;
}

/**
 * buildEmptySheet
 * Convenience helper for a fully empty sheet with safe defaults.
 */
function buildEmptySheet(sheetId, sheetName, orderIndex) {
    return {
        ...DEFAULT_SHEET,
        sheetId,
        sheetName,
        orderIndex,
        data: EMPTY_SHEET_DATA(),
    };
}

// ─────────────────────────────────────────────────────────────────────────────
//  SERIALIZE  (Widget → Mendix)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * serializeSheets
 *
 * Takes the full sheets array (after user edits) and serializes it
 * back to a JSON string ready to write into Mendix.sheetsJson.
 *
 * Strips rows/cols that are entirely null from the bottom/right
 * to keep the stored JSON lean. Does NOT strip user data.
 *
 * @param   {SheetObject[]} sheets  - current widget sheets state
 * @returns {string}                - JSON string for Mendix attribute
 */
export function serializeSheets(sheets) {
    if (!Array.isArray(sheets)) return "[]";

    const output = sheets.map(sheet => serializeSheet(sheet));

    try {
        return JSON.stringify(output);
    } catch (e) {
        console.error("[ExcelWidget] Failed to serialize sheets:", e.message);
        return "[]";
    }
}

/**
 * serializeSheet
 *
 * Serializes a single sheet object.
 * Trims trailing empty rows and columns to reduce JSON size.
 *
 * @param   {SheetObject} sheet
 * @returns {object}             - plain object (not yet stringified)
 */
export function serializeSheet(sheet) {
    return {
        sheetId:     sheet.sheetId,
        sheetName:   sheet.sheetName,
        orderIndex:  sheet.orderIndex,
        isEditable:  sheet.isEditable,
        data:        trimData(sheet.data),
        cellMeta:    sheet.cellMeta    ?? {},
        colWidths:   sheet.colWidths   ?? [],
        rowHeights:  sheet.rowHeights  ?? [],
        mergedCells: sheet.mergedCells ?? [],
    };
}

/**
 * trimData
 *
 * Removes trailing all-null rows and trailing all-null columns
 * from the bottom-right of the grid.
 * Keeps at least 1 row so Mendix always gets a valid array.
 *
 * @param   {(any)[][]} data
 * @returns {(any)[][]}
 */
function trimData(data) {
    if (!Array.isArray(data) || data.length === 0) return [[]];

    // Deep clone so we don't mutate widget state
    let rows = data.map(row => [...(Array.isArray(row) ? row : [])]);

    // Remove trailing all-null rows
    while (rows.length > 1 && rows[rows.length - 1].every(v => v === null || v === undefined || v === "")) {
        rows.pop();
    }

    // Find max used column index across all rows
    let maxUsedCol = 0;
    rows.forEach(row => {
        for (let c = row.length - 1; c >= 0; c--) {
            if (row[c] !== null && row[c] !== undefined && row[c] !== "") {
                if (c > maxUsedCol) maxUsedCol = c;
                break;
            }
        }
    });

    // Trim each row to maxUsedCol + 1
    rows = rows.map(row => row.slice(0, maxUsedCol + 1));

    return rows;
}

// ─────────────────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * updateSheetData
 *
 * Returns a NEW sheets array with one sheet's data replaced.
 * Pure function — never mutates input.
 *
 * @param   {SheetObject[]} sheets        - current sheets array
 * @param   {string}        sheetId       - which sheet to update
 * @param   {(any)[][]}     newData       - new 2D data array from HotTable
 * @returns {SheetObject[]}               - new array with updated sheet
 */
export function updateSheetData(sheets, sheetId, newData) {
    return sheets.map(sheet =>
        sheet.sheetId === sheetId
            ? { ...sheet, data: newData }
            : sheet
    );
}

/**
 * updateSheetMeta
 *
 * Returns a NEW sheets array with one sheet's cellMeta replaced.
 *
 * @param   {SheetObject[]} sheets        - current sheets array
 * @param   {string}        sheetId       - which sheet to update
 * @param   {object}        newMeta       - new cellMeta object
 * @returns {SheetObject[]}
 */
export function updateSheetMeta(sheets, sheetId, newMeta) {
    return sheets.map(sheet =>
        sheet.sheetId === sheetId
            ? { ...sheet, cellMeta: newMeta }
            : sheet
    );
}

/**
 * updateSheetDimensions
 *
 * Returns a NEW sheets array with one sheet's colWidths / rowHeights updated.
 *
 * @param   {SheetObject[]} sheets
 * @param   {string}        sheetId
 * @param   {{ colWidths?, rowHeights? }} dimensions
 * @returns {SheetObject[]}
 */
export function updateSheetDimensions(sheets, sheetId, dimensions) {
    return sheets.map(sheet =>
        sheet.sheetId === sheetId
            ? {
                ...sheet,
                colWidths:  dimensions.colWidths  ?? sheet.colWidths,
                rowHeights: dimensions.rowHeights ?? sheet.rowHeights,
            }
            : sheet
    );
}

// ─────────────────────────────────────────────────────────────────────────────
//  SHEET STRUCTURE MUTATIONS  (add / delete / rename)
//  All return NEW arrays — never mutate input.
//  These are called by WorkbookContainer handlers → auto-save persists them.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * addSheet
 * Adds a new empty sheet to the sheets array.
 * Sheet name auto-increments: Sheet1, Sheet2, Sheet3 ...
 *
 * @param   {SheetObject[]} sheets
 * @returns {SheetObject[]}
 */
export function addSheet(sheets) {
    const existingNumbers = sheets
        .map(s => {
            const match = s.sheetName.match(/^Sheet(\d+)$/);
            return match ? parseInt(match[1], 10) : 0;
        })
        .filter(n => n > 0);

    const nextNumber = existingNumbers.length
        ? Math.max(...existingNumbers) + 1
        : sheets.length + 1;

    const newSheet = {
        sheetId:     `sheet-${Date.now()}`,
        sheetName:   `Sheet${nextNumber}`,
        orderIndex:  sheets.length,
        isEditable:  true,
        data:        EMPTY_SHEET_DATA(),
        cellMeta:    {},
        colWidths:   [],
        rowHeights:  [],
        mergedCells: [],
    };

    return [...sheets, newSheet];
}

/**
 * deleteSheet
 * Removes a sheet by sheetId.
 * Recalculates orderIndex for remaining sheets.
 *
 * @param   {SheetObject[]} sheets
 * @param   {string}        sheetId
 * @returns {SheetObject[]}
 */
export function deleteSheet(sheets, sheetId) {
    return sheets
        .filter(s => s.sheetId !== sheetId)
        .map((s, i) => ({ ...s, orderIndex: i }));
}

/**
 * renameSheet
 * Updates the sheetName for a given sheetId.
 *
 * @param   {SheetObject[]} sheets
 * @param   {string}        sheetId
 * @param   {string}        newName
 * @returns {SheetObject[]}
 */
export function renameSheet(sheets, sheetId, newName) {
    return sheets.map(s =>
        s.sheetId === sheetId
            ? { ...s, sheetName: newName.trim() }
            : s
    );
}
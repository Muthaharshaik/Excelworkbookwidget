/**
 * dataService.js
 *
 * All data transformation between Mendix (JSON strings) and
 * the widget's internal sheet objects.
 */

import { DEFAULT_SHEET, EMPTY_SHEET_DATA, MIN_ROWS, MIN_COLS } from "../utils/constants";

// ─────────────────────────────────────────────────────────────────────────────
//  PARSE  (Mendix → Widget)
// ─────────────────────────────────────────────────────────────────────────────

export function parseSheets(jsonString) {
    if (!jsonString || typeof jsonString !== "string" || jsonString.trim() === "") {
        console.warn("[ExcelWidget] sheetsJson is empty or not a string.");
        return [];
    }

    let rawSheets;
    try {
        rawSheets = JSON.parse(jsonString);
    } catch (e) {
        console.error("[ExcelWidget] Failed to parse sheetsJson:", e.message);
        return [];
    }

    if (!Array.isArray(rawSheets)) {
        console.error("[ExcelWidget] sheetsJson must be a JSON array.");
        return [];
    }

    const normalised = rawSheets.map((raw, index) => normaliseSheet(raw, index));
    normalised.sort((a, b) => a.orderIndex - b.orderIndex);
    return normalised;
}

function normaliseSheet(raw, index) {
    if (!raw || typeof raw !== "object") {
        return buildEmptySheet(`sheet-${index}`, `Sheet${index + 1}`, index);
    }

    return {
        sheetId:     raw.sheetId     ?? `sheet-${index}`,
        sheetName:   raw.sheetName   ?? `Sheet${index + 1}`,
        orderIndex:  typeof raw.orderIndex === "number" ? raw.orderIndex : index,
        isEditable:  raw.isEditable  === true,
        columns:     Array.isArray(raw.columns)     ? raw.columns     : [],
        data:        normaliseData(raw.data),
        cellMeta:    (raw.cellMeta && typeof raw.cellMeta === "object") ? raw.cellMeta : {},
        colWidths:   Array.isArray(raw.colWidths)   ? raw.colWidths   : [],
        rowHeights:  Array.isArray(raw.rowHeights)  ? raw.rowHeights  : [],
        mergedCells: Array.isArray(raw.mergedCells) ? raw.mergedCells : [],
    };
}

function normaliseData(raw) {
    if (!Array.isArray(raw) || raw.length === 0) return EMPTY_SHEET_DATA();

    const maxCols = Math.max(MIN_COLS, ...raw.map(r => (Array.isArray(r) ? r.length : 0)));

    const paddedRows = raw.map(row => {
        if (!Array.isArray(row)) return Array(maxCols).fill(null);
        if (row.length >= maxCols) return row;
        return [...row, ...Array(maxCols - row.length).fill(null)];
    });

    while (paddedRows.length < MIN_ROWS) {
        paddedRows.push(Array(maxCols).fill(null));
    }

    return paddedRows;
}

function buildEmptySheet(sheetId, sheetName, orderIndex) {
    return { ...DEFAULT_SHEET, sheetId, sheetName, orderIndex, data: EMPTY_SHEET_DATA(), columns: [] };
}

// ─────────────────────────────────────────────────────────────────────────────
//  SERIALIZE  (Widget → Mendix)
// ─────────────────────────────────────────────────────────────────────────────

export function serializeSheets(sheets) {
    if (!Array.isArray(sheets)) return "[]";
    try {
        return JSON.stringify(sheets.map(sheet => serializeSheet(sheet)));
    } catch (e) {
        console.error("[ExcelWidget] Failed to serialize sheets:", e.message);
        return "[]";
    }
}

export function serializeSheet(sheet) {
    return {
        sheetId:     sheet.sheetId,
        sheetName:   sheet.sheetName,
        orderIndex:  sheet.orderIndex,
        isEditable:  sheet.isEditable,
        columns:     sheet.columns     ?? [],
        data:        trimData(sheet.data),
        cellMeta:    sheet.cellMeta    ?? {},
        colWidths:   sheet.colWidths   ?? [],
        rowHeights:  sheet.rowHeights  ?? [],
        mergedCells: sheet.mergedCells ?? [],
    };
}

function trimData(data) {
    if (!Array.isArray(data) || data.length === 0) return [[]];

    let rows = data.map(row => [...(Array.isArray(row) ? row : [])]);

    while (rows.length > 1 && rows[rows.length - 1].every(v => v === null || v === undefined || v === "")) {
        rows.pop();
    }

    let maxUsedCol = 0;
    rows.forEach(row => {
        for (let c = row.length - 1; c >= 0; c--) {
            if (row[c] !== null && row[c] !== undefined && row[c] !== "") {
                if (c > maxUsedCol) maxUsedCol = c;
                break;
            }
        }
    });

    return rows.map(row => row.slice(0, maxUsedCol + 1));
}

// ─────────────────────────────────────────────────────────────────────────────
//  CELL / META / DIMENSION UPDATES
// ─────────────────────────────────────────────────────────────────────────────

export function updateSheetData(sheets, sheetId, newData) {
    return sheets.map(s => s.sheetId === sheetId ? { ...s, data: newData } : s);
}

export function updateSheetMeta(sheets, sheetId, newMeta) {
    return sheets.map(s => s.sheetId === sheetId ? { ...s, cellMeta: newMeta } : s);
}

export function updateSheetDimensions(sheets, sheetId, dimensions) {
    return sheets.map(s =>
        s.sheetId === sheetId
            ? { ...s,
                colWidths:  dimensions.colWidths  ?? s.colWidths,
                rowHeights: dimensions.rowHeights ?? s.rowHeights,
              }
            : s
    );
}

// ─────────────────────────────────────────────────────────────────────────────
//  SHEET STRUCTURE  (add / delete / rename)
// ─────────────────────────────────────────────────────────────────────────────

export function addSheet(sheets) {
    const existingNumbers = sheets
        .map(s => { const m = s.sheetName.match(/^Sheet(\d+)$/); return m ? parseInt(m[1], 10) : 0; })
        .filter(n => n > 0);

    const nextNumber = existingNumbers.length ? Math.max(...existingNumbers) + 1 : sheets.length + 1;

    return [...sheets, {
        sheetId:     `sheet-${Date.now()}`,
        sheetName:   `Sheet${nextNumber}`,
        orderIndex:  sheets.length,
        isEditable:  true,
        columns:     [],
        data:        EMPTY_SHEET_DATA(),
        cellMeta:    {},
        colWidths:   [],
        rowHeights:  [],
        mergedCells: [],
    }];
}

export function deleteSheet(sheets, sheetId) {
    return sheets
        .filter(s => s.sheetId !== sheetId)
        .map((s, i) => ({ ...s, orderIndex: i }));
}

export function renameSheet(sheets, sheetId, newName) {
    return sheets.map(s => s.sheetId === sheetId ? { ...s, sheetName: newName.trim() } : s);
}

// ─────────────────────────────────────────────────────────────────────────────
//  COLUMN DEFINITION MUTATIONS
// ─────────────────────────────────────────────────────────────────────────────

export function addColumn(sheets, sheetId) {
    return sheets.map(sheet => {
        if (sheet.sheetId !== sheetId) return sheet;

        const colIndex = sheet.columns.length;
        const newCol   = {
            key:      `col-${Date.now()}`,
            header:   `Column ${colIndex + 1}`,
            type:     "text",
            width:    120,
            source:   [],
            format:   "",
            readOnly: false,
        };

        const newData = (sheet.data || []).map(row => [...row, null]);
        return { ...sheet, columns: [...sheet.columns, newCol], data: newData };
    });
}

export function updateColumn(sheets, sheetId, colKey, changes) {
    return sheets.map(sheet => {
        if (sheet.sheetId !== sheetId) return sheet;
        return {
            ...sheet,
            columns: sheet.columns.map(col => col.key === colKey ? { ...col, ...changes } : col),
        };
    });
}

export function deleteColumn(sheets, sheetId, colKey) {
    return sheets.map(sheet => {
        if (sheet.sheetId !== sheetId) return sheet;

        const colIndex = sheet.columns.findIndex(c => c.key === colKey);
        if (colIndex === -1) return sheet;

        const newColumns = sheet.columns.filter(c => c.key !== colKey);
        const newData    = (sheet.data || []).map(row => {
            const r = [...row];
            r.splice(colIndex, 1);
            return r;
        });

        return { ...sheet, columns: newColumns, data: newData };
    });
}

export function reorderColumn(sheets, sheetId, fromIndex, toIndex) {
    return sheets.map(sheet => {
        if (sheet.sheetId !== sheetId) return sheet;

        const cols = [...sheet.columns];
        const [moved] = cols.splice(fromIndex, 1);
        cols.splice(toIndex, 0, moved);

        const newData = (sheet.data || []).map(row => {
            const r = [...row];
            const [movedCell] = r.splice(fromIndex, 1);
            r.splice(toIndex, 0, movedCell);
            return r;
        });

        return { ...sheet, columns: cols, data: newData };
    });
}
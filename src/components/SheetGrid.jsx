/**
 * SheetGrid.jsx
 * Typed columns + custom cell formatting + custom row labels.
 *
 * KEY FIX: Custom renderer is only applied to text/numeric/date columns.
 * Checkbox and dropdown use HotTable's native renderers — applying our
 * custom text renderer on top of them was overriding their UI entirely.
 *
 * ROW HEADER WIDTH FIX:
 * When custom row labels are set, we measure the longest label with a
 * canvas and pass rowHeaderWidth to HotTable so the frozen row-header
 * pane is exactly wide enough. Without this HotTable defaults to ~50px,
 * labels get cropped and the row/column borders misalign.
 *
 * TAB NAVIGATION FIX:
 * HotTable's Tab shortcut has preventDefault:false, so the browser's native
 * Tab (moving DOM focus to the next element) fires alongside HT and wins.
 * Fix: native capture-phase keydown listener on the container div calls
 * preventDefault() to keep focus in the grid, then manually moves the
 * cell selection using hot.selectCell(). Shift+Tab moves backwards.
 *
 * VALIDATION FIX:
 * allowInvalid:true → HotTable shows htInvalid (red cell) when a typed
 * column receives the wrong data type. We track invalid cells in a ref
 * (NOT state — avoids React error #185) and patch getData() in afterChange
 * so invalid values are never saved to React state or Mendix.
 */

import { createElement, useRef, useCallback, useEffect, memo, useMemo } from "react";
import { HotTable }    from "@handsontable/react";
import Handsontable    from "handsontable";
import "handsontable/dist/handsontable.full.min.css";

import {
    HOT_LICENSE_KEY,
    CONTEXT_MENU_ITEMS,
    DEFAULT_COL_WIDTH,
    DEFAULT_ROW_HEIGHT,
    DEFAULT_DATE_FORMAT,
    DEFAULT_NUMERIC_FORMAT,
} from "../utils/constants";
import { cellKey, deepClone } from "../utils/helpers";

// ── Row header width helper ────────────────────────────────────────────────────
// Uses an offscreen canvas to measure text accurately.
// Falls back to a character-width estimate if canvas is unavailable.

let _measureCanvas = null;

function measureTextPx(text, font) {
    try {
        if (!_measureCanvas) _measureCanvas = document.createElement("canvas");
        const ctx = _measureCanvas.getContext("2d");
        ctx.font  = font;
        return ctx.measureText(String(text)).width;
    } catch {
        return String(text).length * 7.5; // ~7.5px per char at 13px
    }
}

/**
 * Calculates the pixel width the row header column needs so all labels fit.
 *
 *  - When using default numeric row numbers → returns 50 (HotTable default)
 *  - When using custom string labels → measures the longest label and adds
 *    horizontal padding so text is never clipped
 */
function calcRowHeaderWidth(labels, hasCustomLabels) {
    if (!hasCustomLabels || labels.length === 0) return 50;

    const FONT      = "500 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
    const PAD       = 20; // 10px left + 10px right
    const MIN       = 60;
    const MAX       = 240;

    const maxTextPx = labels.reduce((best, label) => {
        const w = measureTextPx(label ?? "", FONT);
        return w > best ? w : best;
    }, 0);

    return Math.min(MAX, Math.max(MIN, Math.ceil(maxTextPx) + PAD));
}

// ─────────────────────────────────────────────────────────────────────────────

export const SheetGrid = memo(function SheetGrid({
    sheet,
    isEditable,
    isAdmin,
    height,
    rowHeaders,
    colHeaders,
    hotRef,
    onCellChange,
    onMetaChange,
    onDimensionChange,
    onAuditLog,
    auditJson,
}) {
    if (!sheet) return null;

    const internalRef   = useRef(null);
    const gridRef       = hotRef ?? internalRef;
    const rendererName  = `ewwRenderer_${sheet.sheetId}`;
    const cellMetaRef   = useRef(sheet.cellMeta);
    cellMetaRef.current = sheet.cellMeta;

    // Tracks cells that currently have invalid values.
    // Key = "row_col", value = true.
    // Using a ref (NOT state) to avoid React setState inside HT hooks → error #185.
    const invalidCellsRef = useRef(new Map());

    // ── Register custom renderer ───────────────────────────────────────────
    // Only used for text / numeric / date columns.
    // Checkbox and dropdown columns must NOT use this — they need
    // HotTable's native renderers to display correctly.
    useMemo(() => {
        Handsontable.renderers.registerRenderer(
            rendererName,
            function (hotInstance, TD, row, col, prop, value, cellProperties) {
                Handsontable.renderers.TextRenderer.apply(this, arguments);
                const meta = cellMetaRef.current?.[cellKey(row, col)];
                if (!meta) return;
                if (meta.bold)      TD.style.fontWeight      = "bold";
                if (meta.italic)    TD.style.fontStyle       = "italic";
                if (meta.underline) TD.style.textDecoration  = "underline";
                if (meta.fontColor) TD.style.color           = meta.fontColor;
                if (meta.bgColor)   TD.style.backgroundColor = meta.bgColor;
                if (meta.align)     TD.style.textAlign       = meta.align;
            }
        );
    }, [rendererName]);

    // ── Build column definitions ───────────────────────────────────────────
    const { hotColumns, hotColHeaders } = useMemo(() => {
        const cols = sheet.columns || [];
        if (cols.length === 0) {
            return { hotColumns: undefined, hotColHeaders: colHeaders };
        }

        const hotCols = cols.map(col => {
            const baseReadOnly = col.readOnly || !isEditable;

            switch (col.type) {

                case "numeric":
                    return {
                        type:          "numeric",
                        renderer:      rendererName,
                        width:         col.width || DEFAULT_COL_WIDTH,
                        readOnly:      baseReadOnly,
                        numericFormat: { pattern: col.format || DEFAULT_NUMERIC_FORMAT },
                        // allowInvalid:true → HT marks cell red (htInvalid class)
                        // and keeps focus. We block the invalid value from saving
                        // in afterChange via invalidCellsRef.
                        allowInvalid:  true,
                    };

                case "date":
                    return {
                        type:          "date",
                        renderer:      rendererName,
                        width:         col.width || DEFAULT_COL_WIDTH,
                        readOnly:      baseReadOnly,
                        dateFormat:    col.format || DEFAULT_DATE_FORMAT,
                        correctFormat: true,
                        allowInvalid:  true,
                    };

                case "checkbox":
                    // DO NOT apply custom renderer — HotTable needs its own
                    // checkbox renderer to draw the actual checkbox input.
                    return {
                        type:     "checkbox",
                        width:    col.width || DEFAULT_COL_WIDTH,
                        readOnly: baseReadOnly,
                    };

                case "dropdown":
                    // DO NOT apply custom renderer — HotTable needs its own
                    // autocomplete/dropdown renderer for the select UI.
                    return {
                        type:         "dropdown",
                        width:        col.width || DEFAULT_COL_WIDTH,
                        readOnly:     baseReadOnly,
                        source:       Array.isArray(col.source) && col.source.length > 0
                                          ? col.source
                                          : ["Option 1", "Option 2", "Option 3"],
                        strict:       true,
                        allowInvalid: true,
                    };

                default:
                    return {
                        type:     "text",
                        renderer: rendererName,
                        width:    col.width || DEFAULT_COL_WIDTH,
                        readOnly: baseReadOnly,
                    };
            }
        });

        const headerLabels = cols.map(col =>
            String(col.header || "")
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
        );

        return { hotColumns: hotCols, hotColHeaders: headerLabels };
    }, [sheet.columns, isEditable, colHeaders, rendererName]);

    // ── Row labels ─────────────────────────────────────────────────────────
    const rowLabels    = sheet.rowLabels || [];
    const hasRowLabels = rowLabels.length > 0;

    // Auto-fit: recalculate width whenever labels change
    const rowHeaderWidth = useMemo(
        () => calcRowHeaderWidth(rowLabels, hasRowLabels),
        [rowLabels, hasRowLabels]
    );

    const gridData = useMemo(() => {
        const fullData = deepClone(sheet.data);
        if (!hasRowLabels) return fullData;
        return fullData.slice(0, rowLabels.length);
    }, [sheet.data, rowLabels.length, hasRowLabels]);

    // Row header renderer: function for custom labels, boolean for default numbers
    const hotRowHeaders = useMemo(() => {
        if (!hasRowLabels) return rowHeaders;
        return (rowIndex) => rowLabels[rowIndex] ?? "";
    }, [rowLabels, hasRowLabels, rowHeaders]);

    // ── cells callback ─────────────────────────────────────────────────────
    const cells = useCallback(() => {
        if (sheet.columns?.length) return {};
        return { renderer: rendererName };
    }, [rendererName, sheet.columns]);

    // ── HotTable event handlers ────────────────────────────────────────────
    const afterChange = useCallback((changes, source) => {
        if (source === "loadData" || !changes || !onCellChange) return;
        const hot = gridRef.current?.hotInstance;
        if (!hot) return;

        // ── Validation gate ────────────────────────────────────────────
        // allowInvalid:true means HT keeps invalid values in its internal data.
        // hot.getData() returns the FULL grid — including cells that were marked
        // invalid in a PREVIOUS edit (not just the current changes batch).
        //
        // Example of the bug this fixes:
        //   1. User types "abc" in a Number cell → marked invalid, not saved ✓
        //   2. User edits a different valid cell → afterChange fires for that cell
        //   3. dataToSave = hot.getData() → includes "abc" still sitting in HT data
        //   4. Without this fix, "abc" gets saved to Mendix ✗
        //
        // Fix: always patch ALL currently invalid cells when building dataToSave,
        // regardless of whether they appear in the current changes batch.
        const invalids = invalidCellsRef.current;

        // If ALL changes in this batch are invalid, nothing valid to save at all
        const hasValidChange = changes.some(([r, c]) => !invalids.has(`${r}_${c}`));
        if (!hasValidChange && invalids.size > 0 && changes.every(([r, c]) => invalids.has(`${r}_${c}`))) return;

        // Build a safe copy of the grid data with ALL invalid cells set to null
        // (we use null because we don't have the original value for cells that
        // were invalid before this change batch — null is safer than wrong data)
        let dataToSave;
        if (invalids.size > 0) {
            dataToSave = hot.getData().map(row => [...row]);
            // Patch cells invalid in the current batch — we have their old values
            changes.forEach(([r, c, oldValue]) => {
                if (invalids.has(`${r}_${c}`)) {
                    dataToSave[r][c] = oldValue ?? null;
                }
            });
            // Patch any other currently invalid cells we DON'T have old values for
            // Set them to null — better than saving a wrong-type value
            invalids.forEach((_, key) => {
                const [r, c] = key.split("_").map(Number);
                const alreadyPatched = changes.some(([cr, cc]) => cr === r && cc === c);
                if (!alreadyPatched) {
                    dataToSave[r][c] = null;
                }
            });
        } else {
            dataToSave = hot.getData();
        }

        // ── Save cell data ─────────────────────────────────────────────
        onCellChange(sheet.sheetId, dataToSave);

        // ── Audit log ──────────────────────────────────────────────────
        if (!onAuditLog || !auditJson) return;

        const cols         = sheet.columns  || [];
        const rowLabelList = sheet.rowLabels || [];

        const auditChanges = changes
            .filter(([r, c, oldVal, newVal]) => {
                if (invalids.has(`${r}_${c}`)) return false; // skip invalid
                const o = oldVal === null || oldVal === undefined ? "" : String(oldVal);
                const n = newVal === null || newVal === undefined ? "" : String(newVal);
                return o !== n;
            })
            .map(([row, col, oldVal, newVal]) => ({
                row,
                col,
                colHeader: cols[col]?.header || String.fromCharCode(65 + (col % 26)),
                rowLabel:  rowLabelList[row]  || String(row + 1),
                oldValue:  oldVal === null || oldVal === undefined ? "" : String(oldVal),
                newValue:  newVal === null || newVal === undefined ? "" : String(newVal),
            }));

        if (auditChanges.length === 0) return;

        const auditPayload = JSON.stringify({
            sheetId:   sheet.sheetId,
            sheetName: sheet.sheetName,
            changes:   auditChanges,
        });

        try {
            if (auditJson.status === "available" && typeof auditJson.setValue === "function") {
                auditJson.setValue(auditPayload);
                if (onAuditLog.canExecute) {
                    onAuditLog.execute();
                }
            }
        } catch (err) {
            console.error("[ExcelWidget] Audit log failed:", err.message);
        }
    }, [sheet.sheetId, sheet.sheetName, sheet.columns, sheet.rowLabels,
        onCellChange, onAuditLog, auditJson, gridRef]);

    const afterColumnResize = useCallback(() => {
        if (!onDimensionChange) return;
        const hot = gridRef.current?.hotInstance;
        if (!hot) return;
        const colWidths = Array.from({ length: hot.countCols() },
            (_, i) => hot.getColWidth(i) ?? DEFAULT_COL_WIDTH);
        onDimensionChange(sheet.sheetId, { colWidths });
    }, [sheet.sheetId, onDimensionChange, gridRef]);

    const afterRowResize = useCallback(() => {
        if (!onDimensionChange) return;
        const hot = gridRef.current?.hotInstance;
        if (!hot) return;
        const rowHeights = Array.from({ length: hot.countRows() },
            (_, i) => hot.getRowHeight(i) ?? DEFAULT_ROW_HEIGHT);
        onDimensionChange(sheet.sheetId, { rowHeights });
    }, [sheet.sheetId, onDimensionChange, gridRef]);

    const afterMergeCells = useCallback(() => {
        const hot = gridRef.current?.hotInstance;
        if (!hot || !onMetaChange) return;
        const plugin      = hot.getPlugin("mergeCells");
        const mergedCells = plugin?.mergedCellsCollection?.mergedCells ?? [];
        onMetaChange(sheet.sheetId, { ...sheet.cellMeta, _mergedCells: mergedCells });
    }, [sheet.sheetId, sheet.cellMeta, onMetaChange, gridRef]);

    // afterValidate fires synchronously inside HT's own validation cycle.
    // MUST only mutate a ref here — never call setState.
    // setState here → React re-render → HT re-validates → afterValidate again → error #185.
    const afterValidate = useCallback((isValid, value, row, prop) => {
        const col = typeof prop === "number" ? prop : parseInt(prop, 10);
        const cols = sheet.columns || [];
        if (!cols[col]) return; // no typed column — nothing to validate
        const key = `${row}_${col}`;
        if (isValid) {
            invalidCellsRef.current.delete(key);
        } else {
            invalidCellsRef.current.set(key, true);
        }
    }, [sheet.columns]);

    // Tab navigation fix — see header comment for full explanation.
    const containerRef = useRef(null);

    useEffect(() => {
        const el = containerRef.current;
        if (!el) return;
        const onKeyDown = (event) => {
            if (event.key !== "Tab") return;
            if (!el.contains(document.activeElement)) return;
            event.preventDefault();
            event.stopImmediatePropagation();
            const hot = gridRef.current?.hotInstance;
            if (!hot || !hot.selection) return;
            if (event.shiftKey) {
                hot.selection.transformStart(0, -1);
            } else {
                hot.selection.transformStart(0, 1);
            }
        };
        document.addEventListener("keydown", onKeyDown, { capture: true });
        return () => document.removeEventListener("keydown", onKeyDown, { capture: true });
    }, [gridRef]);

    const colWidths  = sheet.colWidths?.length  ? sheet.colWidths  : DEFAULT_COL_WIDTH;
    const rowHeights = sheet.rowHeights?.length ? sheet.rowHeights : DEFAULT_ROW_HEIGHT;

    return (
        <div ref={containerRef} style={{ width: "100%", height: "100%" }}>
        <HotTable
            ref={gridRef}
            data={gridData}
            licenseKey={HOT_LICENSE_KEY}
            colHeaders={hotColHeaders}
            columns={hotColumns}
            rowHeaders={hotRowHeaders}
            rowHeaderWidth={rowHeaderWidth}
            width="100%"
            height={height}
            colWidths={colWidths}
            rowHeights={rowHeights}
            readOnly={!isEditable}
            manualColumnResize={true}
            manualRowResize={true}
            contextMenu={isEditable ? CONTEXT_MENU_ITEMS : ["copy"]}
            multiColumnSorting={true}
            filters={true}
            dropdownMenu={true}
            autoWrapRow={true}
            autoWrapCol={true}
            fillHandle={isEditable}
            copyPaste={true}
            outsideClickDeselects={false}
            undo={isEditable}
            stretchH="last"
            mergeCells={sheet.mergedCells?.length ? sheet.mergedCells : true}
            cells={hotColumns ? undefined : cells}
            afterChange={afterChange}
            afterColumnResize={afterColumnResize}
            afterRowResize={afterRowResize}
            afterMergeCells={afterMergeCells}
            afterUnmergeCells={afterMergeCells}
            afterValidate={afterValidate}
        />
        </div>
    );
});
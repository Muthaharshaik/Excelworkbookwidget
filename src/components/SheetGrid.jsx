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
 */

import { createElement, useRef, useCallback, memo, useMemo } from "react";
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
}) {
    if (!sheet) return null;

    const internalRef   = useRef(null);
    const gridRef       = hotRef ?? internalRef;
    const rendererName  = `ewwRenderer_${sheet.sheetId}`;
    const cellMetaRef   = useRef(sheet.cellMeta);
    cellMetaRef.current = sheet.cellMeta;

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
                        type:         "numeric",
                        renderer:     rendererName,
                        width:        col.width || DEFAULT_COL_WIDTH,
                        readOnly:     baseReadOnly,
                        numericFormat: { pattern: col.format || DEFAULT_NUMERIC_FORMAT },
                        allowInvalid: false,
                    };

                case "date":
                    return {
                        type:          "date",
                        renderer:      rendererName,
                        width:         col.width || DEFAULT_COL_WIDTH,
                        readOnly:      baseReadOnly,
                        dateFormat:    col.format || DEFAULT_DATE_FORMAT,
                        correctFormat: true,
                        allowInvalid:  false,
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
                        type:     "dropdown",
                        width:    col.width || DEFAULT_COL_WIDTH,
                        readOnly: baseReadOnly,
                        source:   Array.isArray(col.source) && col.source.length > 0
                                    ? col.source
                                    : ["Option 1", "Option 2", "Option 3"],
                        strict:       true,
                        allowInvalid: false,
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
        onCellChange(sheet.sheetId, hot.getData());
    }, [sheet.sheetId, onCellChange, gridRef]);

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

    const colWidths  = sheet.colWidths?.length  ? sheet.colWidths  : DEFAULT_COL_WIDTH;
    const rowHeights = sheet.rowHeights?.length ? sheet.rowHeights : DEFAULT_ROW_HEIGHT;

    return (
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
        />
    );
});
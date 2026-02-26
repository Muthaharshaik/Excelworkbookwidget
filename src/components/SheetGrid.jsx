/**
 * SheetGrid.jsx - Step A
 * Typed columns + custom cell formatting. No HyperFormula. No CDN.
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

    // Register renderer synchronously before HotTable first paints
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

    // Build typed columns from sheet.columns[]
    // If columns[] is empty → use HotTable default A,B,C headers
    // If columns[] is defined → custom header names + data types
    const { hotColumns, hotColHeaders } = useMemo(() => {
        const cols = sheet.columns || [];
        if (cols.length === 0) {
            return { hotColumns: undefined, hotColHeaders: colHeaders };
        }

        const typeIcons = { text: "T", numeric: "#", date: "📅", time: "⏰", checkbox: "☑", dropdown: "▾" };

        const hotCols = cols.map(col => {
            const base = {
                renderer: rendererName,
                width:    col.width || DEFAULT_COL_WIDTH,
                readOnly: col.readOnly || !isEditable,
            };
            switch (col.type) {
                case "numeric":  return { ...base, type: "numeric",  numericFormat: { pattern: col.format || DEFAULT_NUMERIC_FORMAT } };
                case "date":     return { ...base, type: "date",     dateFormat: col.format || DEFAULT_DATE_FORMAT, correctFormat: true };
                case "checkbox": return { ...base, type: "checkbox" };
                case "dropdown": return { ...base, type: "dropdown", source: col.source || [] };
                case "time":     return { ...base, type: "time",     timeFormat: col.format || "hh:mm:ss", correctFormat: true };
                default:         return { ...base, type: "text" };
            }
        });

        const headerLabels = cols.map(col => {
            const icon = typeIcons[col.type] || "T";
            const safe = String(col.header || "")
                .replace(/&/g, "&amp;").replace(/</g, "&lt;")
                .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
            return isAdmin
                ? `<span title="Type: ${col.type}">${icon} ${safe}</span>`
                : safe;
        });

        return { hotColumns: hotCols, hotColHeaders: headerLabels };
    }, [sheet.columns, isEditable, isAdmin, colHeaders, rendererName]);

    const cells = useCallback(() => {
        if (sheet.columns?.length) return {};
        return { renderer: rendererName };
    }, [rendererName, sheet.columns]);

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
            data={deepClone(sheet.data)}
            licenseKey={HOT_LICENSE_KEY}
            colHeaders={hotColHeaders}
            columns={hotColumns}
            rowHeaders={rowHeaders}
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
/**
 * SheetGrid.jsx
 *
 * KEY FIX: Custom renderer registered SYNCHRONOUSLY via useMemo (not useEffect).
 * HotTable calls the renderer during first paint — before effects run.
 * cellMeta stored in a ref so renderer always reads latest value without re-registering.
 */

import { createElement, useRef, useCallback, memo, useMemo } from "react";
import { HotTable }  from "@handsontable/react";
import Handsontable  from "handsontable";
import "handsontable/dist/handsontable.full.min.css";

import {
    HOT_LICENSE_KEY,
    CONTEXT_MENU_ITEMS,
    DEFAULT_COL_WIDTH,
    DEFAULT_ROW_HEIGHT,
} from "../utils/constants";
import { cellKey, deepClone } from "../utils/helpers";

export const SheetGrid = memo(function SheetGrid({
    sheet,
    isEditable,
    height,
    rowHeaders,
    colHeaders,
    hotRef,
    onCellChange,
    onMetaChange,
    onDimensionChange,
}) {
    if (!sheet) return null;

    const internalRef = useRef(null);
    const gridRef     = hotRef ?? internalRef;

    // Unique renderer name per sheet
    const rendererName = `ewwRenderer_${sheet.sheetId}`;

    // Keep cellMeta in a ref — renderer reads this directly (no stale closure)
    const cellMetaRef = useRef(sheet.cellMeta);
    cellMetaRef.current = sheet.cellMeta;

    // Register renderer SYNCHRONOUSLY during render (useMemo, not useEffect)
    // HotTable needs the renderer to exist on its very first paint
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

    const cells = useCallback(() => ({ renderer: rendererName }), [rendererName]);

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
        const colWidths = Array.from({ length: hot.countCols() }, (_, i) => hot.getColWidth(i) ?? DEFAULT_COL_WIDTH);
        onDimensionChange(sheet.sheetId, { colWidths });
    }, [sheet.sheetId, onDimensionChange, gridRef]);

    const afterRowResize = useCallback(() => {
        if (!onDimensionChange) return;
        const hot = gridRef.current?.hotInstance;
        if (!hot) return;
        const rowHeights = Array.from({ length: hot.countRows() }, (_, i) => hot.getRowHeight(i) ?? DEFAULT_ROW_HEIGHT);
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
            rowHeaders={rowHeaders}
            colHeaders={colHeaders}
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
            cells={cells}
            afterChange={afterChange}
            afterColumnResize={afterColumnResize}
            afterRowResize={afterRowResize}
            afterMergeCells={afterMergeCells}
            afterUnmergeCells={afterMergeCells}
        />
    );
});
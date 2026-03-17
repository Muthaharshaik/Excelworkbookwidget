/**
 * SheetGrid.jsx — DEFINITIVE FINAL
 *
 * ROOT CAUSE (confirmed from handsontable.full.js source):
 *
 *   HotTable's Formulas plugin registers BOTH afterLoadData AND afterUpdateData
 *   to the same handler _onAfterLoadData, which does:
 *
 *     if (!hotWasInitializedWithEmptyData) {
 *       setSheetContent(HOT data)   // safe path
 *     } else {
 *       switchSheet(sheetName)      // BLEED: reads HF formula data into grid
 *     }
 *
 *   hotWasInitializedWithEmptyData = isUndefined(hot.getSettings().data)
 *   = true when no 'data' prop was passed to HotTable at init time.
 *
 * THE FIX:
 *
 *   1. Pass real sheet data as 'data' prop to <HotTable> at mount.
 *      → hotWasInitializedWithEmptyData = false permanently.
 *      → afterLoadData always takes setSheetContent path (safe).
 *      → switchSheet is never called. No bleed.
 *
 *   2. For subsequent data updates (Mendix sends new sheetJson):
 *      Use hot.batch(() => hot.setDataAtCell(changes, 'loadData'))
 *      → fires afterSetDataAtCell → HF updated correctly ✓
 *      → fires afterChange(source='loadData') → blocked by USER_EDIT_SOURCES ✓
 *      → does NOT fire afterLoadData/afterUpdateData → no switchSheet → no bleed ✓
 *      → does NOT fire afterChange with user-edit source → no spurious save ✓
 *
 *   NEVER call hot.loadData() or hot.updateData() — both fire _onAfterLoadData
 *   which would call switchSheet if hotWasInitializedWithEmptyData were true,
 *   and even with the data prop fix, they reset internal HotTable state in ways
 *   that can re-trigger the bleed path on edge cases.
 */

import { createElement, useRef, useCallback, useEffect, memo, useMemo } from "react";
import { HotTable }    from "@handsontable/react";
import Handsontable    from "handsontable";
import "handsontable/dist/handsontable.full.min.css";

import {
    HOT_LICENSE_KEY, CONTEXT_MENU_ITEMS,
    DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT,
    DEFAULT_DATE_FORMAT, DEFAULT_NUMERIC_FORMAT, MIN_COLS,
} from "../utils/constants";
import { cellKey, deepClone }                from "../utils/helpers";
import { buildHeaderRefMap, maybeTranslate } from "../utils/formulaTranslator";

const USER_EDIT_SOURCES = new Set([
    "edit", "CopyPaste.paste", "Autofill.fill",
    "UndoRedo.undo", "UndoRedo.redo", "populateFromArray",
]);

let _measureCanvas = null;
function measureTextPx(text, font) {
    try {
        if (!_measureCanvas) _measureCanvas = document.createElement("canvas");
        const ctx = _measureCanvas.getContext("2d"); ctx.font = font;
        return ctx.measureText(String(text)).width;
    } catch { return String(text).length * 7.5; }
}
function calcRowHeaderWidth(labels, hasCustomLabels) {
    if (!hasCustomLabels || labels.length === 0) return 50;
    const FONT = "500 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
    const maxPx = labels.reduce((best, l) => Math.max(best, measureTextPx(l ?? "", FONT)), 0);
    return Math.min(240, Math.max(60, Math.ceil(maxPx) + 20));
}
function buildEmptyGrid(rowCount) {
    return Array.from({ length: rowCount || 50 }, () => Array(MIN_COLS).fill(null));
}
function computeGridData(sheetData, rowLabels, hasRowLabels, headerRefMap) {
    const fullData = deepClone(sheetData);
    const data = hasRowLabels ? fullData.slice(0, rowLabels.length) : fullData;
    if (headerRefMap.size > 0) {
        data.forEach((row, ri) => {
            if (!Array.isArray(row)) return;
            row.forEach((val, ci) => {
                if (typeof val === "string" && val.startsWith("=")) {
                    const t = maybeTranslate(val, headerRefMap);
                    if (t !== val) data[ri][ci] = t;
                }
            });
        });
    }
    return data;
}
function buildOriginalFormulaMap(sheetData, headerRefMap, targetMap) {
    targetMap.clear();
    if (!headerRefMap.size) return;
    (sheetData || []).forEach((row, ri) => {
        if (!Array.isArray(row)) return;
        row.forEach((val, ci) => {
            if (typeof val === "string" && val.startsWith("=")) {
                const t = maybeTranslate(val, headerRefMap);
                if (t !== val) targetMap.set(`${ri}_${ci}`, val);
            }
        });
    });
}

export const SheetGrid = memo(function SheetGrid({
    sheet, isEditable, isAdmin, height, rowHeaders, colHeaders,
    hotRef, hfRef, onCellChange, onMetaChange, onDimensionChange, onAuditLog, auditJson,
}) {
    if (!sheet) return null;

    const internalRef      = useRef(null);
    const gridRef          = hotRef ?? internalRef;
    const sheetIdAtMount   = useRef(sheet.sheetId);
    const sheetNameAtMount = useRef(sheet.sheetName);

    const rendererName  = `ewwRenderer_${sheet.sheetId}`;
    const cellMetaRef   = useRef(sheet.cellMeta);
    cellMetaRef.current = sheet.cellMeta;

    const invalidCellsRef     = useRef(new Map());
    const originalFormulasRef = useRef(new Map());
    const isLoadingRef        = useRef(true);

    const headerRefMap = useMemo(
        () => buildHeaderRefMap(sheet.columns, sheet.rowLabels),
        [sheet.columns, sheet.rowLabels]
    );
    const rowLabels    = sheet.rowLabels || [];
    const hasRowLabels = rowLabels.length > 0;

    // ── Initial data: computed once at mount, passed as 'data' prop ───────
    // This is the bleed fix. Passing real data sets hotWasInitializedWithEmptyData=false.
    // HotTable's afterLoadData then calls setSheetContent(HOT data) — safe path.
    // switchSheet() is never called. No formula data bleeds between sheets.
    const gridDataRef = useRef(null);
    if (gridDataRef.current === null) {
        // Runs synchronously during first render (ref init pattern)
        const data = computeGridData(sheet.data, rowLabels, hasRowLabels, headerRefMap);
        buildOriginalFormulaMap(sheet.data, headerRefMap, originalFormulasRef.current);
        gridDataRef.current = data;
    }

    // ── formulasConfig: built once at mount with correct sheetName ────────
    const formulasConfig = useMemo(() => {
        const hf = hfRef?.current;
        if (!hf) return false;
        return { engine: hf, sheetName: sheetNameAtMount.current, evaluateNullToZero: true };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []); // empty deps — one config per HotTable instance

    // ── Handle data updates from Mendix (new sheetJson arrives) ──────────
    // Uses hot.batch + setDataAtCell with source='loadData'.
    // This correctly updates HF (via afterSetDataAtCell) without:
    //   - firing afterLoadData (no switchSheet, no bleed)
    //   - firing afterChange with user-edit source (no spurious save)
    useEffect(() => {
        isLoadingRef.current = true;

        const newData = computeGridData(sheet.data, rowLabels, hasRowLabels, headerRefMap);
        buildOriginalFormulaMap(sheet.data, headerRefMap, originalFormulasRef.current);
        gridDataRef.current = newData;

        const hot = gridRef.current?.hotInstance;
        if (hot) {
            // Build flat change list for batch setDataAtCell
            const changes = [];
            newData.forEach((row, ri) => {
                if (!Array.isArray(row)) return;
                row.forEach((val, ci) => changes.push([ri, ci, val]));
            });

            if (changes.length > 0) {
                // batch() suspends render until all changes applied — performant
                hot.batch(() => hot.setDataAtCell(changes, "loadData"));
            }
        }

        // 3 rAFs: covers HF's full recalculation cycle
        let f = requestAnimationFrame(() =>
            requestAnimationFrame(() =>
                requestAnimationFrame(() => {
                    if (sheetIdAtMount.current === sheet.sheetId) {
                        isLoadingRef.current = false;
                    }
                })
            )
        );
        return () => cancelAnimationFrame(f);
    }, [sheet.sheetId, sheet.data]);

    // ── Custom renderer ───────────────────────────────────────────────────
    useMemo(() => {
        Handsontable.renderers.registerRenderer(rendererName, function (hot, TD, row, col) {
            Handsontable.renderers.TextRenderer.apply(this, arguments);
            const meta = cellMetaRef.current?.[cellKey(row, col)];
            if (!meta) return;
            if (meta.bold)      TD.style.fontWeight      = "bold";
            if (meta.italic)    TD.style.fontStyle       = "italic";
            if (meta.underline) TD.style.textDecoration  = "underline";
            if (meta.fontColor) TD.style.color           = meta.fontColor;
            if (meta.bgColor)   TD.style.backgroundColor = meta.bgColor;
            if (meta.align)     TD.style.textAlign       = meta.align;
        });
    }, [rendererName]);

    // ── Column definitions ────────────────────────────────────────────────
    const { hotColumns, hotColHeaders } = useMemo(() => {
        const cols = sheet.columns || [];
        if (!cols.length) return { hotColumns: undefined, hotColHeaders: colHeaders };
        const hotCols = cols.map(col => {
            const ro = col.readOnly || !isEditable;
            switch (col.type) {
                case "numeric":  return { type: "numeric",  renderer: rendererName, width: col.width || DEFAULT_COL_WIDTH, readOnly: ro, numericFormat: { pattern: col.format || DEFAULT_NUMERIC_FORMAT }, allowInvalid: true };
                case "date":     return { type: "date",     renderer: rendererName, width: col.width || DEFAULT_COL_WIDTH, readOnly: ro, dateFormat: col.format || DEFAULT_DATE_FORMAT, correctFormat: true, allowInvalid: true };
                case "checkbox": return { type: "checkbox", width: col.width || DEFAULT_COL_WIDTH, readOnly: ro };
                case "dropdown": return { type: "dropdown", width: col.width || DEFAULT_COL_WIDTH, readOnly: ro, source: col.source?.length ? col.source : ["Option 1", "Option 2", "Option 3"], strict: true, allowInvalid: true };
                default:         return { type: "text",     renderer: rendererName, width: col.width || DEFAULT_COL_WIDTH, readOnly: ro };
            }
        });
        const headers = cols.map(c =>
            String(c.header || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
        );
        return { hotColumns: hotCols, hotColHeaders: headers };
    }, [sheet.columns, isEditable, colHeaders, rendererName]);

    const rowHeaderWidth = useMemo(() => calcRowHeaderWidth(rowLabels, hasRowLabels), [rowLabels, hasRowLabels]);
    const hotRowHeaders  = useMemo(() => hasRowLabels ? (i) => rowLabels[i] ?? "" : rowHeaders, [rowLabels, hasRowLabels, rowHeaders]);
    const cells          = useCallback(() => sheet.columns?.length ? {} : { renderer: rendererName }, [rendererName, sheet.columns]);

    // ── afterChange: save user edits ──────────────────────────────────────
    const afterChange = useCallback((changes, source) => {
        // "loadData" = our setDataAtCell calls above — intentionally blocked
        // "updateData" = HF recalculation — blocked
        // Only real user edits pass through
        if (!USER_EDIT_SOURCES.has(source)) return;
        if (!changes || !onCellChange) return;
        if (isLoadingRef.current) return;
        if (sheetIdAtMount.current !== sheet.sheetId) return;

        const hot = gridRef.current?.hotInstance;
        if (!hot) return;

        const invalids = invalidCellsRef.current;
        const hasValid = changes.some(([r, c]) => !invalids.has(`${r}_${c}`));
        if (!hasValid && invalids.size && changes.every(([r, c]) => invalids.has(`${r}_${c}`))) return;

        let dataToSave;
        if (invalids.size) {
            dataToSave = hot.getSourceData().map(r => [...r]);
            changes.forEach(([r, c, ov]) => { if (invalids.has(`${r}_${c}`)) dataToSave[r][c] = ov ?? null; });
            invalids.forEach((_, key) => {
                const [r, c] = key.split("_").map(Number);
                if (!changes.some(([cr, cc]) => cr === r && cc === c)) dataToSave[r][c] = null;
            });
        } else {
            dataToSave = hot.getSourceData();
        }

        // Restore original (un-translated) formulas before persisting
        originalFormulasRef.current.forEach((orig, key) => {
            const [r, c] = key.split("_").map(Number);
            if (dataToSave[r]?.[c] !== undefined) dataToSave[r][c] = orig;
        });

        onCellChange(sheet.sheetId, dataToSave);

        if (!onAuditLog || !auditJson) return;
        const auditChanges = changes
            .filter(([r, c, ov, nv]) => !invalids.has(`${r}_${c}`) && String(ov ?? "") !== String(nv ?? ""))
            .map(([row, col, ov, nv]) => ({
                row, col,
                colHeader: (sheet.columns || [])[col]?.header || String.fromCharCode(65 + col % 26),
                rowLabel:  (sheet.rowLabels || [])[row] || String(row + 1),
                oldValue: String(ov ?? ""), newValue: String(nv ?? ""),
            }));
        if (!auditChanges.length) return;
        try {
            if (auditJson.status === "available" && typeof auditJson.setValue === "function") {
                auditJson.setValue(JSON.stringify({ sheetId: sheet.sheetId, sheetName: sheet.sheetName, changes: auditChanges }));
                if (onAuditLog.canExecute) onAuditLog.execute();
            }
        } catch (e) { console.info("[SG] Audit log failed:", e.message); }
    }, [sheet.sheetId, sheet.sheetName, sheet.columns, sheet.rowLabels, onCellChange, onAuditLog, auditJson, gridRef]);

    // ── beforeChange: translate header refs to A1 ─────────────────────────
    const beforeChange = useCallback((changes, source) => {
        if (source === "loadData" || !changes || !headerRefMap.size) return;
        changes.forEach((change, idx) => {
            if (!change) return;
            const [row, col, , newVal] = change;
            if (typeof newVal === "string" && newVal.startsWith("=")) {
                const t = maybeTranslate(newVal, headerRefMap);
                if (t !== newVal) {
                    originalFormulasRef.current.set(`${row}_${col}`, newVal);
                    changes[idx][3] = t;
                }
            }
        });
    }, [headerRefMap]);

    // ── afterBeginEditing: show original formula in editor ────────────────
    const afterBeginEditing = useCallback((row, col) => {
        const orig = originalFormulasRef.current.get(`${row}_${col}`);
        if (!orig) return;
        try {
            const editor = gridRef.current?.hotInstance?.getActiveEditor();
            if (editor?.isOpened()) editor.setValue(orig);
        } catch {}
    }, [gridRef]);

    // ── Resize / merge / validate ─────────────────────────────────────────
    const afterColumnResize = useCallback(() => {
        const hot = gridRef.current?.hotInstance; if (!hot || !onDimensionChange) return;
        onDimensionChange(sheet.sheetId, { colWidths: Array.from({ length: hot.countCols() }, (_, i) => hot.getColWidth(i) ?? DEFAULT_COL_WIDTH) });
    }, [sheet.sheetId, onDimensionChange, gridRef]);

    const afterRowResize = useCallback(() => {
        const hot = gridRef.current?.hotInstance; if (!hot || !onDimensionChange) return;
        onDimensionChange(sheet.sheetId, { rowHeights: Array.from({ length: hot.countRows() }, (_, i) => hot.getRowHeight(i) ?? DEFAULT_ROW_HEIGHT) });
    }, [sheet.sheetId, onDimensionChange, gridRef]);

    const afterMergeCells = useCallback(() => {
        const hot = gridRef.current?.hotInstance; if (!hot || !onMetaChange) return;
        const mergedCells = hot.getPlugin("mergeCells")?.mergedCellsCollection?.mergedCells ?? [];
        onMetaChange(sheet.sheetId, { ...sheet.cellMeta, _mergedCells: mergedCells });
    }, [sheet.sheetId, sheet.cellMeta, onMetaChange, gridRef]);

    const afterValidate = useCallback((isValid, value, row, prop) => {
        const col = typeof prop === "number" ? prop : parseInt(prop, 10);
        if (!(sheet.columns || [])[col]) return;
        const key = `${row}_${col}`;
        if (isValid) invalidCellsRef.current.delete(key); else invalidCellsRef.current.set(key, true);
    }, [sheet.columns]);

    // ── Tab key fix ───────────────────────────────────────────────────────
    const containerRef = useRef(null);
    useEffect(() => {
        const el = containerRef.current; if (!el) return;
        const handler = (e) => {
            if (e.key !== "Tab" || !el.contains(document.activeElement)) return;
            e.preventDefault(); e.stopImmediatePropagation();
            const hot = gridRef.current?.hotInstance;
            if (hot?.selection) hot.selection.transformStart(0, e.shiftKey ? -1 : 1);
        };
        document.addEventListener("keydown", handler, { capture: true });
        return () => document.removeEventListener("keydown", handler, { capture: true });
    }, [gridRef]);

    const colWidths  = sheet.colWidths?.length  ? sheet.colWidths  : DEFAULT_COL_WIDTH;
    const rowHeights = sheet.rowHeights?.length ? sheet.rowHeights : DEFAULT_ROW_HEIGHT;

    return (
        <div ref={containerRef} style={{ width: "100%", height: "100%" }}>
            <HotTable
                ref={gridRef}
                data={gridDataRef.current}
                licenseKey={HOT_LICENSE_KEY}
                formulas={formulasConfig}
                colHeaders={hotColHeaders}
                columns={hotColumns}
                rowHeaders={hotRowHeaders}
                rowHeaderWidth={rowHeaderWidth}
                width="100%" height={height}
                colWidths={colWidths} rowHeights={rowHeights}
                readOnly={!isEditable}
                manualColumnResize manualRowResize
                contextMenu={isEditable ? CONTEXT_MENU_ITEMS : ["copy"]}
                multiColumnSorting filters dropdownMenu
                autoWrapRow autoWrapCol
                fillHandle={isEditable} copyPaste
                outsideClickDeselects={false} undo={isEditable}
                stretchH="last"
                mergeCells={sheet.mergedCells?.length ? sheet.mergedCells : true}
                cells={hotColumns ? undefined : cells}
                beforeChange={beforeChange}
                afterBeginEditing={afterBeginEditing}
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
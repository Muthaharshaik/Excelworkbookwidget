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
 * pane is exactly wide enough.
 *
 * TAB NAVIGATION FIX:
 * Native capture-phase keydown listener prevents browser Tab from stealing
 * focus away from the grid.
 *
 * HYPERFORMULA:
 * hfRef is passed in from WorkbookContainer (one instance per widget mount).
 * Passed to HotTable via the formulas prop. SheetGrid is only rendered after
 * hfReady=true (gated in WorkbookContainer) so HotTable always mounts with
 * a live HF engine — formulas activate correctly on first render.
 *
 * STRUCTURED REFERENCES:
 * Users can reference cells using custom column/row header names instead
 * of default A1 notation. Separator is underscore (_).
 *
 * Examples (columns: Revenue/Cost/Profit, rows: Q1/Q2/Q3/Q4):
 *   =SUM(Revenue_Q1, Revenue_Q2)       works
 *   =AVERAGE(Revenue_Q1:Revenue_Q4)    works
 *   =Cost_Q3 - Revenue_Q3             works
 *   =IF(Revenue_Q1>1000,"Good","Bad") works
 *   =Revenue_1 + Cost_2               works (row number fallback)
 *   =A_Q1 + B_Q2                      works (column letter + row header)
 *   =A1 + B2                          works (standard A1 always works)
 *
 * EDIT MODE FIX:
 * originalFormulasRef stores the original formula with header names.
 * Populated on mount by scanning sheet.data (handles page refresh).
 * Also populated in beforeChange when user types new formulas.
 * afterBeginEditing restores original in editor on double click.
 * afterChange restores originals before saving to Mendix.
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
import { cellKey, deepClone }                  from "../utils/helpers";
import { buildHeaderRefMap, maybeTranslate }   from "../utils/formulaTranslator";

// ── Row header width helper ────────────────────────────────────────────────────

let _measureCanvas = null;

function measureTextPx(text, font) {
    try {
        if (!_measureCanvas) _measureCanvas = document.createElement("canvas");
        const ctx = _measureCanvas.getContext("2d");
        ctx.font  = font;
        return ctx.measureText(String(text)).width;
    } catch {
        return String(text).length * 7.5;
    }
}

function calcRowHeaderWidth(labels, hasCustomLabels) {
    if (!hasCustomLabels || labels.length === 0) return 50;

    const FONT      = "500 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
    const PAD       = 20;
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
    hfRef,
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

    const invalidCellsRef = useRef(new Map());

    // ── Store original formulas (with header names) ────────────────────────
    // Key: "row_col", Value: original formula e.g. "=SUM(Cost_Q1, Cost_Q3)"
    //
    // Populated in TWO ways:
    // 1. On mount — scans sheet.data for existing header formulas (page refresh)
    // 2. In beforeChange — when user types new header formulas
    //
    // Used by:
    // - afterBeginEditing → show original formula in editor on double click
    // - afterChange       → save original formula to Mendix (not A1 version)
    const originalFormulasRef = useRef(new Map());

    // ── Build header reference map ─────────────────────────────────────────
    const headerRefMap = useMemo(
        () => buildHeaderRefMap(sheet.columns, sheet.rowLabels),
        [sheet.columns, sheet.rowLabels]
    );

    // ── Scan sheet.data on mount/sheet-switch to populate originalFormulasRef
    // This handles page refresh — sheet.data is loaded from Mendix with
    // header-based formulas (e.g. =AVERAGE(Cost_Q1:Cost_Q3)).
    // We scan and register them so afterBeginEditing and afterChange work.
    // beforeChange has loadData guard so we MUST populate here for refresh.
    useEffect(() => {
        originalFormulasRef.current.clear();
        if (headerRefMap.size === 0) return;

        const data = sheet.data || [];
        data.forEach((row, rowIndex) => {
            if (!Array.isArray(row)) return;
            row.forEach((cellValue, colIndex) => {
                if (typeof cellValue === "string" && cellValue.startsWith("=")) {
                    const translated = maybeTranslate(cellValue, headerRefMap);
                    if (translated !== cellValue) {
                        // This cell has a header-based formula → store original
                        originalFormulasRef.current.set(
                            `${rowIndex}_${colIndex}`,
                            cellValue
                        );
                    }
                }
            });
        });
    }, [sheet.sheetId, sheet.data, headerRefMap]);

    // ── HyperFormula config ────────────────────────────────────────────────
    const formulasConfig = hfRef?.current
        ? {
            engine:             hfRef.current,
            evaluateNullToZero: true,
          }
        : false;

    // ── Register custom renderer ───────────────────────────────────────────
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
                    return {
                        type:     "checkbox",
                        width:    col.width || DEFAULT_COL_WIDTH,
                        readOnly: baseReadOnly,
                    };
                case "dropdown":
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

    const rowHeaderWidth = useMemo(
        () => calcRowHeaderWidth(rowLabels, hasRowLabels),
        [rowLabels, hasRowLabels]
    );

const gridData = useMemo(() => {
    const fullData = deepClone(sheet.data);
    const data = hasRowLabels ? fullData.slice(0, rowLabels.length) : fullData;

    // Translate header-based formulas to A1 notation for HF
    // sheet.data stores =AVERAGE(Cost_Q1:Cost_Q3) → translate to =AVERAGE(B1:B3)
    // so HF can evaluate correctly on load
    if (headerRefMap.size > 0) {
        data.forEach((row, rowIndex) => {
            if (!Array.isArray(row)) return;
            row.forEach((cellValue, colIndex) => {
                if (typeof cellValue === "string" && cellValue.startsWith("=")) {
                    const translated = maybeTranslate(cellValue, headerRefMap);
                    if (translated !== cellValue) {
                        data[rowIndex][colIndex] = translated;
                    }
                }
            });
        });
    }

    return data;
}, [sheet.data, rowLabels.length, hasRowLabels, headerRefMap]);

    const hotRowHeaders = useMemo(() => {
        if (!hasRowLabels) return rowHeaders;
        return (rowIndex) => rowLabels[rowIndex] ?? "";
    }, [rowLabels, hasRowLabels, rowHeaders]);

    // ── cells callback ─────────────────────────────────────────────────────
    const cells = useCallback(() => {
        if (sheet.columns?.length) return {};
        return { renderer: rendererName };
    }, [rendererName, sheet.columns]);

    // ── beforeChange: translate header refs → A1 before HF evaluates ──────
    // loadData guard is KEPT — translation on loadData would cause #ERROR.
    // Instead, originalFormulasRef is populated via useEffect above for refresh.
    // This hook only handles NEW formulas typed by the user.
    const beforeChange = useCallback((changes, source) => {
        if (source === "loadData" || !changes || headerRefMap.size === 0) return;

        changes.forEach((change, index) => {
            if (!change) return;
            const [row, col, , newValue] = change;
            if (typeof newValue === "string" && newValue.startsWith("=")) {
                const translated = maybeTranslate(newValue, headerRefMap);
                if (translated !== newValue) {
                    // Store original BEFORE replacing with A1 notation
                    originalFormulasRef.current.set(`${row}_${col}`, newValue);
                    changes[index][3] = translated;
                }
            }
        });
    }, [headerRefMap]);

    // ── afterBeginEditing: restore original formula in editor ──────────────
    // HotTable opens editor with HF-stored A1 formula (=AVERAGE(B1:B3)).
    // We replace it with the original header formula (=AVERAGE(Cost_Q1:Cost_Q3)).
    const afterBeginEditing = useCallback((row, col) => {
        const key      = `${row}_${col}`;
        const original = originalFormulasRef.current.get(key);
        if (!original) return;

        try {
            const hot    = gridRef.current?.hotInstance;
            if (!hot) return;
            const editor = hot.getActiveEditor();
            if (editor && editor.isOpened()) {
                editor.setValue(original);
            }
        } catch (e) {
            // safe to ignore
        }
    }, [gridRef]);

    // ── afterChange ────────────────────────────────────────────────────────
    // Restores original header formulas before saving to Mendix.
    // getSourceData() returns A1 translated formulas — we replace them
    // with originals so Mendix stores =AVERAGE(Cost_Q1:Cost_Q3) not =AVERAGE(B1:B3).
    const afterChange = useCallback((changes, source) => {
        if (source === "loadData" || !changes || !onCellChange) return;

        const hot = gridRef.current?.hotInstance;
        if (!hot) return;

        const invalids = invalidCellsRef.current;

        const hasValidChange = changes.some(([r, c]) => !invalids.has(`${r}_${c}`));
        if (!hasValidChange && invalids.size > 0 && changes.every(([r, c]) => invalids.has(`${r}_${c}`))) return;

        let dataToSave;
        if (invalids.size > 0) {
            dataToSave = hot.getSourceData().map(row => [...row]);
            changes.forEach(([r, c, oldValue]) => {
                if (invalids.has(`${r}_${c}`)) {
                    dataToSave[r][c] = oldValue ?? null;
                }
            });
            invalids.forEach((_, key) => {
                const [r, c] = key.split("_").map(Number);
                const alreadyPatched = changes.some(([cr, cc]) => cr === r && cc === c);
                if (!alreadyPatched) {
                    dataToSave[r][c] = null;
                }
            });
        } else {
            dataToSave = hot.getSourceData();
        }

        // Restore original header formulas before saving
        // Replaces A1 notation back to header names for Mendix storage
        if (originalFormulasRef.current.size > 0) {
            originalFormulasRef.current.forEach((originalFormula, key) => {
                const [r, c] = key.split("_").map(Number);
                if (dataToSave[r] !== undefined && dataToSave[r][c] !== undefined) {
                    dataToSave[r][c] = originalFormula;
                }
            });
        }

        onCellChange(sheet.sheetId, dataToSave);

        if (!onAuditLog || !auditJson) return;

        const cols         = sheet.columns  || [];
        const rowLabelList = sheet.rowLabels || [];

        const auditChanges = changes
            .filter(([r, c, oldVal, newVal]) => {
                if (invalids.has(`${r}_${c}`)) return false;
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

    const afterValidate = useCallback((isValid, value, row, prop) => {
        const col = typeof prop === "number" ? prop : parseInt(prop, 10);
        const cols = sheet.columns || [];
        if (!cols[col]) return;
        const key = `${row}_${col}`;
        if (isValid) {
            invalidCellsRef.current.delete(key);
        } else {
            invalidCellsRef.current.set(key, true);
        }
    }, [sheet.columns]);

    // ── Tab navigation fix ─────────────────────────────────────────────────
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
                formulas={formulasConfig}
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
/**
 * WorkbookContainer.jsx — FINAL FIX
 *
 * ROOT CAUSE SUMMARY (proven by reading HotTable source):
 *
 * HotTable Formulas plugin sets:
 *   hotWasInitializedWithEmptyData = isUndefined(hot.getSettings().data)
 *
 * When hotWasInitializedWithEmptyData=true, every afterLoadData calls:
 *   switchSheet() → getSheetSerialized() → loads HF formula data into grid = BLEED
 *
 * When hotWasInitializedWithEmptyData=false (data prop is any array), afterLoadData calls:
 *   setSheetContent(HOT_data) → writes HOT's data into HF = CORRECT
 *
 * THE FIX:
 * Don't render SheetGrid until sheetJsonValue has arrived from Mendix.
 * This means initialGridData in SheetGrid will always be the REAL data,
 * not 50 rows of nulls. Passing this as data={initialGridData} sets
 * hotWasInitializedWithEmptyData=false → switchSheet never runs → no bleed.
 *
 * We track sheetJsonReady: true once we've received at least one valid
 * sheetJson for the current sheetId. SheetGrid only mounts after this.
 */

import { createElement, useRef, useCallback, useState, useEffect } from "react";

import { SheetGrid }           from "./SheetGrid";
import { Toolbar }             from "./Toolbar";
import { ColumnSettingsPanel } from "./ColumnSettingsPanel";
import { RowSettingsPanel }    from "./RowSettingsPanel";
import { ReadOnlyBadge }       from "./ReadOnlyBadge";

import { parseSheetJson, serializeSheet, parseAllSheetsJson } from "../services/dataService";
import { triggerSheetChange }             from "../services/mendixBridge";
import { CSS, AUTOSAVE_DEBOUNCE_MS }      from "../utils/constants";
import { useHyperformula }                from "../hooks/useHyperformula";

export function WorkbookContainer(props) {
    const {
        sheetId, sheetName, sheetJson,
        currentUserId, accessUserId, permissionType, isAdmin,
        onSheetChange, onAuditLog, auditJson,
        allSheetsJson,
        gridHeight = 600, rowCount = 50,
        showToolbar = true, showSheetName = true,
        rowHeaders = true, colHeaders = true,
    } = props;

    const sheetIdValue       = resolveAttr(sheetId)        ?? "";
    const sheetNameValue     = resolveAttr(sheetName)      ?? "Sheet";
    const sheetJsonValue     = resolveAttr(sheetJson);
    const isAdminValue       = resolveAttr(isAdmin)        ?? false;
    const currentUserValue   = resolveAttr(currentUserId)  ?? "";
    const accessUserValue    = resolveAttr(accessUserId)   ?? "";
    const permissionValue    = resolveAttr(permissionType) ?? "View";
    const allSheetsJsonValue = resolveAttr(allSheetsJson)  ?? "";

    const isUserMatch    = currentUserValue && accessUserValue
        && currentUserValue.trim() === accessUserValue.trim();
    const canEditCells   = isAdminValue || (isUserMatch && permissionValue === "Edit");
    const canEditColumns = isAdminValue;

    const [sheetData, setSheetData]             = useState(null); // null = not loaded yet
    const [savingStatus, setSavingStatus]       = useState("idle");
    const [showColumnPanel, setShowColumnPanel] = useState(false);
    const [showRowPanel, setShowRowPanel]       = useState(false);

    // Track whether we have real data for the current sheet
    // SheetGrid only mounts once this is true
    const [sheetJsonReady, setSheetJsonReady]   = useState(false);

    const hotRef        = useRef(null);
    const debounceTimer = useRef(null);
    const savedTimer    = useRef(null);
    const isFirstLoad   = useRef(true);

    const [allSheets, setAllSheets] = useState(() => parseAllSheetsJson(allSheetsJsonValue));

    useEffect(() => {
        setAllSheets(parseAllSheetsJson(allSheetsJsonValue));
    }, [allSheetsJsonValue]);

    const { hfRef, hfReady } = useHyperformula(allSheets, sheetNameValue, hotRef);

    // ── Reset on sheet switch ─────────────────────────────────────────────
    const prevSheetIdRef = useRef(sheetIdValue);
    useEffect(() => {
        if (prevSheetIdRef.current === sheetIdValue) return;
        prevSheetIdRef.current = sheetIdValue;
        // Reset: hide SheetGrid until real data arrives for new sheet
        setSheetData(null);
        setSheetJsonReady(false);
        isFirstLoad.current = true;
    }, [sheetIdValue]);

    // ── Load sheetJson — only show grid once real data is here ────────────
    useEffect(() => {
        if (!sheetJsonValue) return;
        const parsed = parseSheetJson(sheetJsonValue, rowCount);
        if (parsed._sheetId !== null && parsed._sheetId !== sheetIdValue) return;
        setSheetData(parsed);
        setSheetJsonReady(true);
        isFirstLoad.current = true;
    }, [sheetJsonValue]);

    // ── Auto-save ─────────────────────────────────────────────────────────
    useEffect(() => {
        if (!sheetData) return;
        if (isFirstLoad.current) { isFirstLoad.current = false; return; }
        setSavingStatus("saving");
        clearTimeout(debounceTimer.current);
        debounceTimer.current = setTimeout(() => performSave(), AUTOSAVE_DEBOUNCE_MS);
        return () => clearTimeout(debounceTimer.current);
    }, [sheetData]);

    useEffect(() => () => {
        clearTimeout(debounceTimer.current);
        clearTimeout(savedTimer.current);
    }, []);

    const performSave = useCallback(() => {
        if (!sheetData) return;
        try {
            const newJson = serializeSheet(sheetData);
            const success = triggerSheetChange(sheetJson, newJson, onSheetChange);
            if (!success) { setSavingStatus("idle"); return; }
            setSavingStatus("saved");
            clearTimeout(savedTimer.current);
            savedTimer.current = setTimeout(() => setSavingStatus("idle"), 2500);
        } catch (err) {
            setSavingStatus("idle");
        }
    }, [sheetData, sheetJson, onSheetChange]);

    // ── Change handlers ───────────────────────────────────────────────────
    const handleCellChange = useCallback((incomingId, newData) => {
        if (incomingId !== sheetIdValue) return;
        setSheetData(prev => prev ? { ...prev, data: newData } : prev);
    }, [sheetIdValue]);

    const handleMetaChange      = useCallback((m) => setSheetData(p => p ? { ...p, cellMeta: m } : p), []);
    const handleDimensionChange = useCallback((d) => setSheetData(p => p ? { ...p, ...d } : p), []);

    const handleAddColumn = useCallback(() => {
        setSheetData(prev => {
            if (!prev) return prev;
            const cols   = prev.columns || [];
            const newCol = { key: `col-${Date.now()}`, header: `Column ${cols.length + 1}`, type: "text", width: 120, source: [], format: "", readOnly: false };
            return { ...prev, columns: [...cols, newCol], data: (prev.data || []).map(r => [...r, null]) };
        });
    }, []);
    const handleUpdateColumn = useCallback((colKey, ch) => {
        setSheetData(prev => prev ? { ...prev, columns: (prev.columns || []).map(c => c.key === colKey ? { ...c, ...ch } : c) } : prev);
    }, []);
    const handleDeleteColumn = useCallback((colKey) => {
        setSheetData(prev => {
            if (!prev) return prev;
            const idx = (prev.columns || []).findIndex(c => c.key === colKey);
            if (idx === -1) return prev;
            return { ...prev, columns: prev.columns.filter(c => c.key !== colKey), data: (prev.data || []).map(row => { const r = [...row]; r.splice(idx, 1); return r; }) };
        });
    }, []);
    const handleReorderColumn = useCallback((from, to) => {
        setSheetData(prev => {
            if (!prev) return prev;
            const cols = [...(prev.columns || [])]; const [m] = cols.splice(from, 1); cols.splice(to, 0, m);
            const newData = (prev.data || []).map(row => { const r = [...row]; const [mc] = r.splice(from, 1); r.splice(to, 0, mc); return r; });
            return { ...prev, columns: cols, data: newData };
        });
    }, []);
    const handleAddRow = useCallback(() => setSheetData(prev => prev ? { ...prev, rowLabels: [...(prev.rowLabels || []), ""] } : prev), []);
    const handleUpdateRow = useCallback((idx, label) => {
        setSheetData(prev => {
            if (!prev) return prev;
            const labels = [...(prev.rowLabels || [])];
            while (labels.length <= idx) labels.push(""); labels[idx] = label;
            return { ...prev, rowLabels: labels };
        });
    }, []);
    const handleDeleteRow  = useCallback((idx) => setSheetData(prev => prev ? { ...prev, rowLabels: (prev.rowLabels || []).filter((_, i) => i !== idx) } : prev), []);
    const handleReorderRow = useCallback((from, to) => {
        setSheetData(prev => {
            if (!prev) return prev;
            const labels = [...(prev.rowLabels || [])]; const [m] = labels.splice(from, 1); labels.splice(to, 0, m);
            return { ...prev, rowLabels: labels };
        });
    }, []);

    // Build sheet object only when we have data
    const sheet = sheetData ? {
        sheetId: sheetIdValue, sheetName: sheetNameValue, isEditable: canEditCells,
        data: sheetData.data || [], columns: sheetData.columns || [],
        rowLabels: sheetData.rowLabels || [], cellMeta: sheetData.cellMeta || {},
        colWidths: sheetData.colWidths || [], rowHeights: sheetData.rowHeights || [],
        mergedCells: sheetData.mergedCells || [],
    } : {
        sheetId: sheetIdValue, sheetName: sheetNameValue, isEditable: canEditCells,
        data: [], columns: [], rowLabels: [], cellMeta: {},
        colWidths: [], rowHeights: [], mergedCells: [],
    };

    const hasCustomColumns = sheet.columns.length > 0;
    const hasCustomRows    = sheet.rowLabels.length > 0;

    // SheetGrid only renders when BOTH hfReady AND sheetJsonReady
    const gridReady = hfReady && sheetJsonReady;

    return (
        <div className={CSS.WORKBOOK_ROOT}>
            {showSheetName && (
                <div className={CSS.HEADER}>
                    <div className="eww-header__left">
                        <span className="eww-header__sheet-icon">📄</span>
                        <span className="eww-header__title">{sheetNameValue}</span>
                        {canEditColumns && (
                            <div className="eww-header__config-group">
                                <button className={["eww-col-config-btn", hasCustomColumns ? "eww-col-config-btn--active" : ""].filter(Boolean).join(" ")} onClick={() => setShowColumnPanel(true)} title={hasCustomColumns ? `${sheet.columns.length} columns configured` : "Configure column headers"}>
                                    <span className="eww-col-config-btn__icon">⊞</span>
                                    <span className="eww-col-config-btn__label">{hasCustomColumns ? `${sheet.columns.length} Column${sheet.columns.length !== 1 ? "s" : ""}` : "Columns"}</span>
                                    {hasCustomColumns && <span className="eww-col-config-btn__badge">{sheet.columns.length}</span>}
                                </button>
                                <button className={["eww-col-config-btn", "eww-row-config-btn", hasCustomRows ? "eww-col-config-btn--active eww-row-config-btn--active" : ""].filter(Boolean).join(" ")} onClick={() => setShowRowPanel(true)} title={hasCustomRows ? `${sheet.rowLabels.length} row labels configured` : "Configure row labels"}>
                                    <span className="eww-col-config-btn__icon">☰</span>
                                    <span className="eww-col-config-btn__label">{hasCustomRows ? `${sheet.rowLabels.length} Row${sheet.rowLabels.length !== 1 ? "s" : ""}` : "Rows"}</span>
                                    {hasCustomRows && <span className="eww-col-config-btn__badge eww-row-config-btn__badge">{sheet.rowLabels.length}</span>}
                                </button>
                            </div>
                        )}
                    </div>
                    <div className="eww-header__meta">
                        <SavingIndicator status={savingStatus} />
                        {!canEditCells && <ReadOnlyBadge />}
                    </div>
                </div>
            )}

            {showToolbar && (
                <Toolbar hotRef={hotRef} activeSheet={sheet}
                    onMetaChange={(_, m) => handleMetaChange(m)} disabled={!canEditCells} />
            )}

            <div className={CSS.GRID_WRAPPER}>
                {/* Loading state while waiting for real data */}
                {!gridReady && (
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: gridHeight, color: "var(--eww-color-text-muted)", fontSize: 13 }}>
                        Loading…
                    </div>
                )}

                {/* SheetGrid only mounts after real sheetJson has arrived */}
                {gridReady && (
                    <SheetGrid
                        key={sheetIdValue}
                        sheet={sheet}
                        isEditable={canEditCells}
                        isAdmin={canEditColumns}
                        height={gridHeight}
                        rowHeaders={rowHeaders}
                        colHeaders={colHeaders}
                        hotRef={hotRef}
                        hfRef={hfRef}
                        onCellChange={(id, data) => handleCellChange(id, data)}
                        onMetaChange={(_, m) => handleMetaChange(m)}
                        onDimensionChange={(_, d) => handleDimensionChange(d)}
                        onAuditLog={onAuditLog}
                        auditJson={auditJson}
                    />
                )}
            </div>

            {showColumnPanel && canEditColumns && (
                <ColumnSettingsPanel sheet={sheet} isAdmin={canEditColumns}
                    onAddColumn={() => handleAddColumn()} onUpdateColumn={(_, k, c) => handleUpdateColumn(k, c)}
                    onDeleteColumn={(_, k) => handleDeleteColumn(k)} onReorderColumn={(_, f, t) => handleReorderColumn(f, t)}
                    onClose={() => setShowColumnPanel(false)} />
            )}
            {showRowPanel && canEditColumns && (
                <RowSettingsPanel sheet={sheet} isAdmin={canEditColumns}
                    onAddRow={handleAddRow} onUpdateRow={handleUpdateRow}
                    onDeleteRow={handleDeleteRow} onReorderRow={handleReorderRow}
                    onClose={() => setShowRowPanel(false)} />
            )}
        </div>
    );
}

function SavingIndicator({ status }) {
    if (status === "idle") return null;
    if (status === "saving") return (
        <div className="eww-save-indicator eww-save-indicator--saving">
            <div className="eww-save-indicator__spinner" /><span>Saving</span>
        </div>
    );
    return (
        <div className="eww-save-indicator eww-save-indicator--saved">
            <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                <path d="M2 6l3 3 5-5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
            <span>Saved</span>
        </div>
    );
}

function resolveAttr(prop) {
    if (prop === null || prop === undefined) return undefined;
    if (typeof prop === "object" && "status" in prop) return prop.status === "available" ? prop.value : undefined;
    return prop;
}
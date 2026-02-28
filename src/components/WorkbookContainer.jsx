/**
 * WorkbookContainer.jsx
 * Single sheet per widget instance.
 * No JSX fragments anywhere — Mendix compatibility.
 */

import { createElement, useRef, useCallback, useState, useEffect } from "react";

import { SheetGrid }           from "./SheetGrid";
import { Toolbar }             from "./Toolbar";
import { ColumnSettingsPanel } from "./ColumnSettingsPanel";
import { RowSettingsPanel }    from "./RowSettingsPanel";
import { ReadOnlyBadge }       from "./ReadOnlyBadge";

import { parseSheetJson, serializeSheet } from "../services/dataService";
import { triggerSheetChange }             from "../services/mendixBridge";
import { CSS, AUTOSAVE_DEBOUNCE_MS }      from "../utils/constants";

export function WorkbookContainer(props) {
    const {
        sheetId, sheetName, sheetJson,
        currentUserId, accessUserId, permissionType, isAdmin,
        onSheetChange,
        gridHeight = 600, rowCount = 50,
        showToolbar = true, showSheetName = true,
        rowHeaders = true, colHeaders = true,
    } = props;

    const sheetIdValue      = resolveAttr(sheetId)        ?? "";
    const sheetNameValue    = resolveAttr(sheetName)      ?? "Sheet";
    const sheetJsonValue    = resolveAttr(sheetJson);
    const isAdminValue      = resolveAttr(isAdmin)        ?? false;
    const currentUserValue  = resolveAttr(currentUserId)  ?? "";
    const accessUserValue   = resolveAttr(accessUserId)   ?? "";
    const permissionValue   = resolveAttr(permissionType) ?? "View";

    const isUserMatch  = currentUserValue && accessUserValue
        && currentUserValue.trim() === accessUserValue.trim();

    const canEditCells   = isAdminValue || (isUserMatch && permissionValue === "Edit");
    const canEditColumns = isAdminValue;

    console.info("[ExcelWidget] Access Debug:", {
        currentUserValue, accessUserValue, permissionValue,
        isAdminValue, isUserMatch, canEditCells, canEditColumns,
    });

    const [sheetData, setSheetData]             = useState(() => parseSheetJson(sheetJsonValue, rowCount));
    const [savingStatus, setSavingStatus]       = useState("idle");
    const [showColumnPanel, setShowColumnPanel] = useState(false);
    const [showRowPanel, setShowRowPanel]       = useState(false);

    const hotRef        = useRef(null);
    const debounceTimer = useRef(null);
    const savedTimer    = useRef(null);
    const isFirstLoad   = useRef(true);

    useEffect(() => {
        const parsed = parseSheetJson(sheetJsonValue, rowCount);
        setSheetData(parsed);
        isFirstLoad.current = true;
    }, [sheetIdValue]);

    useEffect(() => {
        if (isFirstLoad.current) { isFirstLoad.current = false; return; }
        setSavingStatus("saving");
        clearTimeout(debounceTimer.current);
        debounceTimer.current = setTimeout(() => { performSave(); }, AUTOSAVE_DEBOUNCE_MS);
        return () => clearTimeout(debounceTimer.current);
    }, [sheetData]);

    useEffect(() => {
        return () => { clearTimeout(debounceTimer.current); clearTimeout(savedTimer.current); };
    }, []);

    const performSave = useCallback(() => {
        try {
            const newJson = serializeSheet(sheetData);
            const success = triggerSheetChange(sheetJson, newJson, onSheetChange);
            if (!success) { setSavingStatus("idle"); return; }
            setSavingStatus("saved");
            clearTimeout(savedTimer.current);
            savedTimer.current = setTimeout(() => setSavingStatus("idle"), 2500);
        } catch (err) {
            console.error("[ExcelWidget] Auto-save failed:", err.message);
            setSavingStatus("idle");
        }
    }, [sheetData, sheetJson, onSheetChange]);

    const handleCellChange      = useCallback((newData)    => setSheetData(prev => ({ ...prev, data: newData })), []);
    const handleMetaChange      = useCallback((newMeta)    => setSheetData(prev => ({ ...prev, cellMeta: newMeta })), []);
    const handleDimensionChange = useCallback((dimensions) => setSheetData(prev => ({ ...prev, ...dimensions })), []);

    const handleAddColumn = useCallback(() => {
        setSheetData(prev => {
            const cols   = prev.columns || [];
            const newCol = { key: `col-${Date.now()}`, header: `Column ${cols.length + 1}`, type: "text", width: 120, source: [], format: "", readOnly: false };
            return { ...prev, columns: [...cols, newCol], data: (prev.data || []).map(row => [...row, null]) };
        });
    }, []);

    const handleUpdateColumn = useCallback((colKey, changes) => {
        setSheetData(prev => ({ ...prev, columns: (prev.columns || []).map(c => c.key === colKey ? { ...c, ...changes } : c) }));
    }, []);

    const handleDeleteColumn = useCallback((colKey) => {
        setSheetData(prev => {
            const idx = (prev.columns || []).findIndex(c => c.key === colKey);
            if (idx === -1) return prev;
            return {
                ...prev,
                columns: prev.columns.filter(c => c.key !== colKey),
                data: (prev.data || []).map(row => { const r = [...row]; r.splice(idx, 1); return r; }),
            };
        });
    }, []);

    const handleReorderColumn = useCallback((fromIndex, toIndex) => {
        setSheetData(prev => {
            const cols = [...(prev.columns || [])];
            const [moved] = cols.splice(fromIndex, 1);
            cols.splice(toIndex, 0, moved);
            const newData = (prev.data || []).map(row => {
                const r = [...row];
                const [mc] = r.splice(fromIndex, 1);
                r.splice(toIndex, 0, mc);
                return r;
            });
            return { ...prev, columns: cols, data: newData };
        });
    }, []);

    const handleAddRow = useCallback(() => {
        setSheetData(prev => ({ ...prev, rowLabels: [...(prev.rowLabels || []), ""] }));
    }, []);

    const handleUpdateRow = useCallback((rowIndex, newLabel) => {
        setSheetData(prev => {
            const labels = [...(prev.rowLabels || [])];
            while (labels.length <= rowIndex) labels.push("");
            labels[rowIndex] = newLabel;
            return { ...prev, rowLabels: labels };
        });
    }, []);

    const handleDeleteRow = useCallback((rowIndex) => {
        setSheetData(prev => ({ ...prev, rowLabels: (prev.rowLabels || []).filter((_, i) => i !== rowIndex) }));
    }, []);

    const handleReorderRow = useCallback((fromIndex, toIndex) => {
        setSheetData(prev => {
            const labels = [...(prev.rowLabels || [])];
            const [moved] = labels.splice(fromIndex, 1);
            labels.splice(toIndex, 0, moved);
            return { ...prev, rowLabels: labels };
        });
    }, []);

    const sheet = {
        sheetId: sheetIdValue, sheetName: sheetNameValue, isEditable: canEditCells,
        data: sheetData.data || [], columns: sheetData.columns || [],
        rowLabels: sheetData.rowLabels || [], cellMeta: sheetData.cellMeta || {},
        colWidths: sheetData.colWidths || [], rowHeights: sheetData.rowHeights || [],
        mergedCells: sheetData.mergedCells || [],
    };

    const hasCustomColumns = sheet.columns.length > 0;
    const hasCustomRows    = sheet.rowLabels.length > 0;

    return (
        <div className={CSS.WORKBOOK_ROOT}>

            {showSheetName && (
                <div className={CSS.HEADER}>
                    <div className="eww-header__left">
                        <span className="eww-header__sheet-icon">📄</span>
                        <span className="eww-header__title">{sheetNameValue}</span>

                        {canEditColumns && (
                            <div className="eww-header__config-group">

                                <button
                                    className={["eww-col-config-btn", hasCustomColumns ? "eww-col-config-btn--active" : ""].filter(Boolean).join(" ")}
                                    onClick={() => setShowColumnPanel(true)}
                                    title={hasCustomColumns ? `${sheet.columns.length} columns configured` : "Configure column headers"}
                                >
                                    <span className="eww-col-config-btn__icon">⊞</span>
                                    <span className="eww-col-config-btn__label">
                                        {hasCustomColumns ? `${sheet.columns.length} Column${sheet.columns.length !== 1 ? "s" : ""}` : "Columns"}
                                    </span>
                                    {hasCustomColumns && (
                                        <span className="eww-col-config-btn__badge">{sheet.columns.length}</span>
                                    )}
                                </button>

                                <button
                                    className={["eww-col-config-btn", "eww-row-config-btn", hasCustomRows ? "eww-col-config-btn--active eww-row-config-btn--active" : ""].filter(Boolean).join(" ")}
                                    onClick={() => setShowRowPanel(true)}
                                    title={hasCustomRows ? `${sheet.rowLabels.length} row labels configured` : "Configure row labels"}
                                >
                                    <span className="eww-col-config-btn__icon">☰</span>
                                    <span className="eww-col-config-btn__label">
                                        {hasCustomRows ? `${sheet.rowLabels.length} Row${sheet.rowLabels.length !== 1 ? "s" : ""}` : "Rows"}
                                    </span>
                                    {hasCustomRows && (
                                        <span className="eww-col-config-btn__badge eww-row-config-btn__badge">{sheet.rowLabels.length}</span>
                                    )}
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
                <Toolbar
                    hotRef={hotRef} activeSheet={sheet}
                    onMetaChange={(_, newMeta) => handleMetaChange(newMeta)}
                    disabled={!canEditCells}
                />
            )}

            <div className={CSS.GRID_WRAPPER}>
                <SheetGrid
                    key={sheetIdValue} sheet={sheet}
                    isEditable={canEditCells} isAdmin={canEditColumns}
                    height={gridHeight} rowHeaders={rowHeaders} colHeaders={colHeaders}
                    hotRef={hotRef}
                    onCellChange={(_, newData) => handleCellChange(newData)}
                    onMetaChange={(_, newMeta) => handleMetaChange(newMeta)}
                    onDimensionChange={(_, dims) => handleDimensionChange(dims)}
                />
            </div>

            {showColumnPanel && canEditColumns && (
                <ColumnSettingsPanel
                    sheet={sheet} isAdmin={canEditColumns}
                    onAddColumn={() => handleAddColumn()}
                    onUpdateColumn={(_, colKey, changes) => handleUpdateColumn(colKey, changes)}
                    onDeleteColumn={(_, colKey) => handleDeleteColumn(colKey)}
                    onReorderColumn={(_, from, to) => handleReorderColumn(from, to)}
                    onClose={() => setShowColumnPanel(false)}
                />
            )}

            {showRowPanel && canEditColumns && (
                <RowSettingsPanel
                    sheet={sheet} isAdmin={canEditColumns}
                    onAddRow={handleAddRow}
                    onUpdateRow={handleUpdateRow}
                    onDeleteRow={handleDeleteRow}
                    onReorderRow={handleReorderRow}
                    onClose={() => setShowRowPanel(false)}
                />
            )}

        </div>
    );
}

// ── SavingIndicator ───────────────────────────────────────────────────────────
// No fragments — each branch returns a plain div wrapper

function SavingIndicator({ status }) {
    if (status === "idle") return null;

    const isSaving = status === "saving";

    if (isSaving) {
        return (
            <div className="eww-save-indicator eww-save-indicator--saving">
                <div className="eww-save-indicator__spinner" />
                <span>Saving</span>
            </div>
        );
    }

    return (
        <div className="eww-save-indicator eww-save-indicator--saved">
            <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                <path d="M2 6l3 3 5-5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
            <span>Saved</span>
        </div>
    );
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function resolveAttr(prop) {
    if (prop === null || prop === undefined) return undefined;
    if (typeof prop === "object" && "status" in prop) {
        return prop.status === "available" ? prop.value : undefined;
    }
    return prop;
}
/**
 * WorkbookContainer.jsx
 * Single sheet per widget instance.
 * No JSX fragments anywhere — Mendix compatibility.
 *
 * FIX: sheetData now re-parses whenever sheetJsonValue changes from Mendix
 * (e.g. after a Java action writes a formula result back into sheetJson).
 *
 * Two refs work together to prevent an infinite save loop:
 *
 *   isPendingEdit   — true while the user has an unsaved edit in the debounce
 *                     window. Blocks Mendix-pushed sheetJsonValue from
 *                     overwriting in-flight user changes.
 *
 *   isMendixUpdate  — set to true just before we call setSheetData from the
 *                     Mendix-push effect. The auto-save effect checks this flag
 *                     and skips saving (and immediately resets the flag to false)
 *                     so that a Mendix-originated data update never triggers
 *                     an unnecessary write back to Mendix.
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

    const sheetIdValue        = resolveAttr(sheetId)        ?? "";
    const sheetNameValue      = resolveAttr(sheetName)      ?? "Sheet";
    const sheetJsonValue      = resolveAttr(sheetJson);
    const isAdminValue        = resolveAttr(isAdmin)        ?? false;
    const currentUserValue    = resolveAttr(currentUserId)  ?? "";
    const accessUserValue     = resolveAttr(accessUserId)   ?? "";
    const permissionValue     = resolveAttr(permissionType) ?? "View";
    const allSheetsJsonValue  = resolveAttr(allSheetsJson)  ?? "";

    const isUserMatch  = currentUserValue && accessUserValue
        && currentUserValue.trim() === accessUserValue.trim();

    const canEditCells   = isAdminValue || (isUserMatch && permissionValue === "Edit");
    const canEditColumns = isAdminValue;

    const [sheetData, setSheetData]             = useState(() => parseSheetJson(sheetJsonValue, rowCount));
    const [savingStatus, setSavingStatus]       = useState("idle");
    const [showColumnPanel, setShowColumnPanel] = useState(false);
    const [showRowPanel, setShowRowPanel]       = useState(false);

    const hotRef        = useRef(null);
    const debounceTimer = useRef(null);
    const savedTimer    = useRef(null);
    const isFirstLoad   = useRef(true);

    // ── isPendingEdit: true while user has an unsaved edit in the debounce window.
    // Set when the auto-save effect fires. Cleared when performSave succeeds.
    // While true, Mendix-pushed sheetJsonValue updates are ignored so an
    // in-flight user edit is never overwritten by a concurrent Mendix refresh.
    const isPendingEdit = useRef(false);

    // ── isMendixUpdate: set to true immediately before calling setSheetData
    // from the Mendix-push effect. The auto-save effect reads this flag and
    // skips saving (then immediately resets the flag to false) so that data
    // arriving FROM Mendix never bounces straight back as a save to Mendix,
    // which would create an infinite microflow loop.
    const isMendixUpdate = useRef(false);

    // ── Parse allSheetsJson → allSheets array ─────────────────────────────
    const [allSheets, setAllSheets] = useState(() => parseAllSheetsJson(allSheetsJsonValue));

    useEffect(() => {
        setAllSheets(parseAllSheetsJson(allSheetsJsonValue));
    }, [allSheetsJsonValue]);

    // ── HyperFormula instance ─────────────────────────────────────────────
    const { hfRef, hfReady } = useHyperformula(
        allSheets,
        sheetNameValue,
        hotRef
    );

    // ── Reset on sheet switch ─────────────────────────────────────────────
    // When sheetId changes (user navigated to a different sheet), always
    // re-parse regardless of pending edits — the previous sheet's edit state
    // is irrelevant to the newly loaded sheet.
    // Also reset both guard refs so neither carries stale state across sheets.
    useEffect(() => {
        isPendingEdit.current  = false;
        isMendixUpdate.current = true;   // this setSheetData is not a user edit
        const parsed = parseSheetJson(sheetJsonValue, rowCount);
        setSheetData(parsed);
        isFirstLoad.current = true;
    }, [sheetIdValue]); // eslint-disable-line react-hooks/exhaustive-deps
    // sheetJsonValue intentionally excluded — the effect below handles
    // same-sheet updates from Mendix.

    // ── Accept Mendix-pushed sheetJson updates (formula results, etc.) ────
    // Fires when Mendix writes a new value into sheetJson on the SAME sheet
    // (e.g. a Java action computed a formula and committed the result).
    // Guarded by isPendingEdit so a user's uncommitted edit is never lost.
    // Sets isMendixUpdate before calling setSheetData so the auto-save
    // effect below knows NOT to treat this as a user-initiated change.
    useEffect(() => {
        // Skip the very first render — initial value already handled by useState.
        if (isFirstLoad.current) return;

        // Skip if the user has typed something that hasn't saved yet.
        // performSave will clear this flag once the commit succeeds.
        if (isPendingEdit.current) return;

        const parsed = parseSheetJson(sheetJsonValue, rowCount);
        isMendixUpdate.current = true;   // tell auto-save effect: not a user edit
        setSheetData(parsed);
    }, [sheetJsonValue]); // eslint-disable-line react-hooks/exhaustive-deps
    // rowCount intentionally excluded — it does not change at runtime.

    // ── Auto-save on user edits ───────────────────────────────────────────
    useEffect(() => {
        // First load: mark done and exit without saving.
        if (isFirstLoad.current) { isFirstLoad.current = false; return; }

        // If this sheetData change was pushed FROM Mendix (not typed by the user),
        // reset the flag and bail out — we must not save it back to Mendix.
        if (isMendixUpdate.current) { isMendixUpdate.current = false; return; }

        // Real user edit — mark pending and schedule debounced save.
        isPendingEdit.current = true;
        setSavingStatus("saving");
        clearTimeout(debounceTimer.current);
        debounceTimer.current = setTimeout(() => { performSave(); }, AUTOSAVE_DEBOUNCE_MS);
        return () => clearTimeout(debounceTimer.current);
    }, [sheetData]); // eslint-disable-line react-hooks/exhaustive-deps

    // ── Cleanup on unmount ────────────────────────────────────────────────
    useEffect(() => {
        return () => { clearTimeout(debounceTimer.current); clearTimeout(savedTimer.current); };
    }, []);

    const performSave = useCallback(() => {
        try {
            const newJson = serializeSheet(sheetData);
            const success = triggerSheetChange(sheetJson, newJson, onSheetChange);
            if (!success) { setSavingStatus("idle"); return; }

            // Save succeeded — Mendix now holds the latest value.
            // Clear isPendingEdit so the next Mendix-pushed update (e.g. the
            // formula result written by the Java action that our save triggered)
            // is accepted and rendered by the effect above.
            isPendingEdit.current = false;

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
        sheetId:     sheetIdValue,
        sheetName:   sheetNameValue,
        isEditable:  canEditCells,
        data:        sheetData.data        || [],
        columns:     sheetData.columns     || [],
        rowLabels:   sheetData.rowLabels   || [],
        cellMeta:    sheetData.cellMeta    || {},
        colWidths:   sheetData.colWidths   || [],
        rowHeights:  sheetData.rowHeights  || [],
        mergedCells: sheetData.mergedCells || [],
        lockedCells: sheetData.lockedCells || [],
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
                {hfReady && (
                    <SheetGrid
                        key={sheetIdValue} sheet={sheet}
                        isEditable={canEditCells} isAdmin={canEditColumns}
                        height={gridHeight} rowHeaders={rowHeaders} colHeaders={colHeaders}
                        hotRef={hotRef}
                        hfRef={hfRef}
                        onCellChange={(_, newData) => handleCellChange(newData)}
                        onMetaChange={(_, newMeta) => handleMetaChange(newMeta)}
                        onDimensionChange={(_, dims) => handleDimensionChange(dims)}
                        onAuditLog={onAuditLog}
                        auditJson={auditJson}
                    />
                )}
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
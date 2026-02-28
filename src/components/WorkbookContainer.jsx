/**
 * WorkbookContainer.jsx
 *
 * NEW ARCHITECTURE (single sheet per widget instance):
 *   - Widget sits inside DataView of Spreadsheet entity
 *   - Reads ONE sheet's data from sheetJson attribute
 *   - Writes cell changes back to sheetJson via setValue()
 *   - Mendix ListView handles sheet tabs (Sheet1, Sheet2...)
 *   - No sheet add/delete/rename in widget — Mendix owns that
 *
 * DATA FLOW:
 *   Mendix DataView(Spreadsheet) → sheetJson prop
 *     → parseSheetJson() → grid data
 *       → user edits cell
 *         → serializeSheet() → setValue(newJson)
 *           → onSheetChange.execute() → Mendix commits
 */

import { createElement, useRef, useCallback, useState, useEffect } from "react";

import { SheetGrid }           from "./SheetGrid";
import { Toolbar }             from "./Toolbar";
import { ColumnSettingsPanel } from "./ColumnSettingsPanel";
import { ReadOnlyBadge }       from "./ReadOnlyBadge";

import { parseSheetJson, serializeSheet } from "../services/dataService";
import { triggerSheetChange }             from "../services/mendixBridge";
import { CSS, AUTOSAVE_DEBOUNCE_MS }      from "../utils/constants";

export function WorkbookContainer(props) {
    const {
        sheetId,
        sheetName,
        sheetJson,
        currentUserId,
        accessUserId,
        permissionType,
        isAdmin,
        onSheetChange,
        gridHeight     = 600,
        rowCount       = 50,
        showToolbar    = true,
        showSheetName  = true,
        rowHeaders     = true,
        colHeaders     = true,
    } = props;

    // ── Resolve Mendix attribute / expression values ───────────────────────
    const sheetIdValue      = resolveAttr(sheetId)       ?? "";
    const sheetNameValue    = resolveAttr(sheetName)     ?? "Sheet";
    const sheetJsonValue    = resolveAttr(sheetJson);
    const isAdminValue      = resolveAttr(isAdmin)       ?? false;
    const currentUserValue  = resolveAttr(currentUserId) ?? "";
    const accessUserValue   = resolveAttr(accessUserId)  ?? "";
    const permissionValue   = resolveAttr(permissionType)?? "View";

    // ── Resolve access level ───────────────────────────────────────────────
    const isUserMatch  = currentUserValue
        && accessUserValue
        && currentUserValue.trim() === accessUserValue.trim();

    const canEditCells   = isAdminValue || (isUserMatch && permissionValue === "Edit");
    const canEditColumns = isAdminValue;

    // ── DEBUG (remove after testing) ──────────────────────────────────────
    console.info("[ExcelWidget] Access Debug:", {
        currentUserValue, accessUserValue, permissionValue,
        isAdminValue, isUserMatch, canEditCells, canEditColumns,
    });

    // ── Parse sheet data from JSON ─────────────────────────────────────────
    const [sheetData, setSheetData] = useState(() => parseSheetJson(sheetJsonValue, rowCount));
    const [savingStatus, setSavingStatus] = useState("idle");
    const [showColumnPanel, setShowColumnPanel] = useState(false);

    const hotRef        = useRef(null);
    const debounceTimer = useRef(null);
    const savedTimer    = useRef(null);
    const isFirstLoad   = useRef(true);

    // ── Re-parse when sheet ID changes (user clicked different tab) ────────
    useEffect(() => {
        const parsed = parseSheetJson(sheetJsonValue, rowCount);
        setSheetData(parsed);
        isFirstLoad.current = true;
    }, [sheetIdValue]);

    // ── Auto-save when sheetData changes ──────────────────────────────────
    useEffect(() => {
        if (isFirstLoad.current) {
            isFirstLoad.current = false;
            return;
        }
        setSavingStatus("saving");
        clearTimeout(debounceTimer.current);
        debounceTimer.current = setTimeout(() => { performSave(); }, AUTOSAVE_DEBOUNCE_MS);
        return () => clearTimeout(debounceTimer.current);
    }, [sheetData]);

    // ── Cleanup timers on unmount ──────────────────────────────────────────
    useEffect(() => {
        return () => {
            clearTimeout(debounceTimer.current);
            clearTimeout(savedTimer.current);
        };
    }, []);

    // ── Save ──────────────────────────────────────────────────────────────
    const performSave = useCallback(() => {
        try {
            const newJson = serializeSheet(sheetData);
            const success = triggerSheetChange(sheetJson, newJson, onSheetChange);
            if (!success) { setSavingStatus("idle"); return; }
            setSavingStatus("saved");
            clearTimeout(savedTimer.current);
            savedTimer.current = setTimeout(() => setSavingStatus("idle"), 2000);
        } catch (err) {
            console.error("[ExcelWidget] Auto-save failed:", err.message);
            setSavingStatus("idle");
        }
    }, [sheetData, sheetJson, onSheetChange]);

    // ── Cell / meta / dimension change handlers ────────────────────────────
    const handleCellChange = useCallback((newData) => {
        setSheetData(prev => ({ ...prev, data: newData }));
    }, []);

    const handleMetaChange = useCallback((newMeta) => {
        setSheetData(prev => ({ ...prev, cellMeta: newMeta }));
    }, []);

    const handleDimensionChange = useCallback((dimensions) => {
        setSheetData(prev => ({ ...prev, ...dimensions }));
    }, []);

    // ── Column handlers (admin only) ───────────────────────────────────────
    const handleAddColumn = useCallback(() => {
        setSheetData(prev => {
            const cols   = prev.columns || [];
            const newCol = {
                key: `col-${Date.now()}`, header: `Column ${cols.length + 1}`,
                type: "text", width: 120, source: [], format: "", readOnly: false,
            };
            const newData = (prev.data || []).map(row => [...row, null]);
            return { ...prev, columns: [...cols, newCol], data: newData };
        });
    }, []);

    const handleUpdateColumn = useCallback((colKey, changes) => {
        setSheetData(prev => ({
            ...prev,
            columns: (prev.columns || []).map(c => c.key === colKey ? { ...c, ...changes } : c),
        }));
    }, []);

    const handleDeleteColumn = useCallback((colKey) => {
        setSheetData(prev => {
            const idx = (prev.columns || []).findIndex(c => c.key === colKey);
            if (idx === -1) return prev;
            const newCols = prev.columns.filter(c => c.key !== colKey);
            const newData = (prev.data || []).map(row => { const r = [...row]; r.splice(idx, 1); return r; });
            return { ...prev, columns: newCols, data: newData };
        });
    }, []);

    const handleReorderColumn = useCallback((fromIndex, toIndex) => {
        setSheetData(prev => {
            const cols = [...(prev.columns || [])];
            const [moved] = cols.splice(fromIndex, 1);
            cols.splice(toIndex, 0, moved);
            const newData = (prev.data || []).map(row => {
                const r = [...row];
                const [movedCell] = r.splice(fromIndex, 1);
                r.splice(toIndex, 0, movedCell);
                return r;
            });
            return { ...prev, columns: cols, data: newData };
        });
    }, []);

    // ── Build sheet object for SheetGrid ──────────────────────────────────
    const sheet = {
        sheetId:     sheetIdValue,
        sheetName:   sheetNameValue,
        isEditable:  canEditCells,
        data:        sheetData.data        || [],
        columns:     sheetData.columns     || [],
        cellMeta:    sheetData.cellMeta    || {},
        colWidths:   sheetData.colWidths   || [],
        rowHeights:  sheetData.rowHeights  || [],
        mergedCells: sheetData.mergedCells || [],
    };

    const hasCustomColumns = sheet.columns.length > 0;

    return (
        <div className={CSS.WORKBOOK_ROOT}>

            {/* ── Sheet name header ─────────────────────────────────── */}
            {showSheetName && (
                <div className={CSS.HEADER}>

                    {/* Left: sheet icon + name */}
                    <div className="eww-header__left">
                        <span className="eww-header__sheet-icon">📄</span>
                        <span className="eww-header__title">{sheetNameValue}</span>

                        {/* Admin: Column config button — lives right beside the sheet name */}
                        {canEditColumns && (
                            <button
                                className={[
                                    "eww-col-config-btn",
                                    hasCustomColumns ? "eww-col-config-btn--active" : "",
                                ].filter(Boolean).join(" ")}
                                onClick={() => setShowColumnPanel(true)}
                                title={
                                    hasCustomColumns
                                        ? `${sheet.columns.length} custom column${sheet.columns.length !== 1 ? "s" : ""} configured — click to edit`
                                        : "Configure custom column headers (Admin)"
                                }
                            >
                                <span className="eww-col-config-btn__icon">⚙</span>
                                <span className="eww-col-config-btn__label">
                                    {hasCustomColumns
                                        ? `${sheet.columns.length} Column${sheet.columns.length !== 1 ? "s" : ""}`
                                        : "Columns"
                                    }
                                </span>
                                {hasCustomColumns && (
                                    <span className="eww-col-config-btn__badge">
                                        {sheet.columns.length}
                                    </span>
                                )}
                            </button>
                        )}
                    </div>

                    {/* Right: saving indicator + read-only badge */}
                    <div className="eww-header__meta">
                        <SavingIndicator status={savingStatus} />
                        {!canEditCells && <ReadOnlyBadge />}
                    </div>

                </div>
            )}

            {/* ── Toolbar ───────────────────────────────────────────── */}
            {showToolbar && (
                <Toolbar
                    hotRef={hotRef}
                    activeSheet={sheet}
                    onMetaChange={(_, newMeta) => handleMetaChange(newMeta)}
                    disabled={!canEditCells}
                />
            )}

            {/* ── Grid ──────────────────────────────────────────────── */}
            <div className={CSS.GRID_WRAPPER}>
                <SheetGrid
                    key={sheetIdValue}
                    sheet={sheet}
                    isEditable={canEditCells}
                    isAdmin={canEditColumns}
                    height={gridHeight}
                    rowHeaders={rowHeaders}
                    colHeaders={colHeaders}
                    hotRef={hotRef}
                    onCellChange={(_, newData) => handleCellChange(newData)}
                    onMetaChange={(_, newMeta) => handleMetaChange(newMeta)}
                    onDimensionChange={(_, dims) => handleDimensionChange(dims)}
                />
            </div>

            {/* ── Column settings panel (admin only) ────────────────── */}
            {showColumnPanel && canEditColumns && (
                <ColumnSettingsPanel
                    sheet={sheet}
                    isAdmin={canEditColumns}
                    onAddColumn={() => handleAddColumn()}
                    onUpdateColumn={(_, colKey, changes) => handleUpdateColumn(colKey, changes)}
                    onDeleteColumn={(_, colKey) => handleDeleteColumn(colKey)}
                    onReorderColumn={(_, from, to) => handleReorderColumn(from, to)}
                    onClose={() => setShowColumnPanel(false)}
                />
            )}

        </div>
    );
}

// ── SavingIndicator ───────────────────────────────────────────────────────────

function SavingIndicator({ status }) {
    if (status === "idle") return null;
    const isSaving  = status === "saving";
    const className = [
        CSS.SAVING_INDICATOR,
        isSaving ? "eww-saving-indicator--saving" : "eww-saving-indicator--saved",
    ].join(" ");
    return (
        <span className={className}>
            {isSaving
                ? createElement("span", null,
                    createElement("span", { className: "eww-saving-indicator__dot" }),
                    "Saving…")
                : "✓ Saved"
            }
        </span>
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
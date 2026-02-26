/**
 * WorkbookContainer.jsx - Step A
 * Clean version — no HyperFormula, no CDN, no themes, no permissions panel.
 */

import { createElement, useRef, useCallback, useState } from "react";

import { useWorkbookState }    from "../hooks/useWorkbookState";
import { useAutoSave }         from "../hooks/useAutoSave";
import { usePermissions }      from "../hooks/usePermissions";

import { SheetGrid }           from "./SheetGrid";
import { SheetTabBar }         from "./SheetTabBar";
import { Toolbar }             from "./Toolbar";
import { ReadOnlyBadge }       from "./ReadOnlyBadge";
import { ColumnSettingsPanel } from "./ColumnSettingsPanel";

import {
    updateSheetData, updateSheetMeta, updateSheetDimensions,
    addSheet, deleteSheet, renameSheet,
    addColumn, updateColumn, deleteColumn, reorderColumn,
} from "../services/dataService";
import { CSS }                   from "../utils/constants";
import { triggerSheetTabChange } from "../services/mendixBridge";

export function WorkbookContainer(props) {
    const {
        workbookId,
        workbookName,
        sheetsJson,
        isReadOnly,
        isAdmin,
        onSheetChange,
        onSheetTabChange,
        gridHeight         = 600,
        showToolbar        = true,
        showWorkbookHeader = true,
        rowHeaders         = true,
        colHeaders         = true,
    } = props;

    // ── Resolve Mendix datasource attributes ──────────────────────────────
    const workbookItem     = props.workbookSource?.items?.[0];
    const sheetsJsonAttr   = workbookItem && sheetsJson   ? sheetsJson.get(workbookItem)   : sheetsJson;
    const isReadOnlyAttr   = workbookItem && isReadOnly   ? isReadOnly.get(workbookItem)   : isReadOnly;
    const workbookIdAttr   = workbookItem && workbookId   ? workbookId.get(workbookItem)   : workbookId;
    const workbookNameAttr = workbookItem && workbookName ? workbookName.get(workbookItem) : workbookName;
    const isAdminAttr      = workbookItem && isAdmin      ? isAdmin.get(workbookItem)      : isAdmin;

    const sheetsJsonValue   = resolveAttr(sheetsJsonAttr);
    const isReadOnlyValue   = resolveAttr(isReadOnlyAttr)   ?? false;
    const workbookIdValue   = resolveAttr(workbookIdAttr)   ?? "";
    const workbookNameValue = resolveAttr(workbookNameAttr) ?? "Workbook";
    const isAdminValue      = resolveAttr(isAdminAttr)      ?? false;

    // ── UI state ──────────────────────────────────────────────────────────
    const [showColumnPanel, setShowColumnPanel] = useState(false);
    const hotRef = useRef(null);

    // ── Core state ─────────────────────────────────────────────────────────
    const {
        sheets,
        setSheets,
        activeSheet,
        activeSheetIndex,
        setActiveSheetIndex,
        isLoading,
        parseError,
        markPendingEdits,
        clearPendingEdits,
    } = useWorkbookState(sheetsJsonValue);

    // ── Permissions ────────────────────────────────────────────────────────
    const { canEditSheet } = usePermissions(isReadOnlyValue);

    // ── Auto-save ──────────────────────────────────────────────────────────
    const { savingStatus } = useAutoSave({
        sheets,
        onSheetChange,
        sheetsJson: sheetsJsonAttr,
        clearPendingEdits,
    });

    // ── Handlers ──────────────────────────────────────────────────────────
    const handleCellChange = useCallback((sheetId, newData) => {
        markPendingEdits();
        setSheets(prev => updateSheetData(prev, sheetId, newData));
    }, [markPendingEdits, setSheets]);

    const handleMetaChange = useCallback((sheetId, newMeta) => {
        markPendingEdits();
        setSheets(prev => updateSheetMeta(prev, sheetId, newMeta));
    }, [markPendingEdits, setSheets]);

    const handleDimensionChange = useCallback((sheetId, dimensions) => {
        markPendingEdits();
        setSheets(prev => updateSheetDimensions(prev, sheetId, dimensions));
    }, [markPendingEdits, setSheets]);

    const handleAddSheet = useCallback(() => {
        markPendingEdits();
        setSheets(prev => addSheet(prev));
        setActiveSheetIndex(sheets.length);
    }, [markPendingEdits, setSheets, setActiveSheetIndex, sheets.length]);

    const handleDeleteSheet = useCallback((sheetId) => {
        markPendingEdits();
        setSheets(prev => deleteSheet(prev, sheetId));
        setActiveSheetIndex(Math.max(0, activeSheetIndex - 1));
    }, [markPendingEdits, setSheets, setActiveSheetIndex, activeSheetIndex]);

    const handleRenameSheet = useCallback((sheetId, newName) => {
        markPendingEdits();
        setSheets(prev => renameSheet(prev, sheetId, newName));
    }, [markPendingEdits, setSheets]);

    const handleAddColumn = useCallback((sheetId) => {
        markPendingEdits();
        setSheets(prev => addColumn(prev, sheetId));
    }, [markPendingEdits, setSheets]);

    const handleUpdateColumn = useCallback((sheetId, colKey, changes) => {
        markPendingEdits();
        setSheets(prev => updateColumn(prev, sheetId, colKey, changes));
    }, [markPendingEdits, setSheets]);

    const handleDeleteColumn = useCallback((sheetId, colKey) => {
        markPendingEdits();
        setSheets(prev => deleteColumn(prev, sheetId, colKey));
    }, [markPendingEdits, setSheets]);

    const handleReorderColumn = useCallback((sheetId, fromIndex, toIndex) => {
        markPendingEdits();
        setSheets(prev => reorderColumn(prev, sheetId, fromIndex, toIndex));
    }, [markPendingEdits, setSheets]);

    const handleTabChange = useCallback((index) => {
        setActiveSheetIndex(index);
        triggerSheetTabChange(onSheetTabChange);
    }, [setActiveSheetIndex, onSheetTabChange]);

    // ── Render: error ──────────────────────────────────────────────────────
    if (parseError) {
        return (
            <div className={CSS.WORKBOOK_ROOT} style={styles.errorBox}>
                ⚠ Failed to load workbook data. Please check the sheetsJson configuration.
                <br />
                <small style={{ color: "#999", marginTop: 4, display: "block" }}>{parseError}</small>
            </div>
        );
    }

    // ── Render: loading ────────────────────────────────────────────────────
    if (isLoading) {
        return (
            <div className={CSS.WORKBOOK_ROOT} style={styles.loadingBox}>
                <div style={styles.spinner} />
                <span>Loading workbook…</span>
            </div>
        );
    }

    // ── Render: empty ──────────────────────────────────────────────────────
    if (!sheets.length) {
        return (
            <div className={CSS.WORKBOOK_ROOT} style={styles.emptyBox}>
                <span style={styles.emptyIcon}>📋</span>
                <span>No sheets found in this workbook.</span>
                <small style={{ color: "#999" }}>
                    Ask your administrator to configure sheets for this workbook.
                </small>
            </div>
        );
    }

    const activeSheetEditable = activeSheet ? canEditSheet(activeSheet.isEditable) : false;

    return (
        <div className={CSS.WORKBOOK_ROOT}>

            {showWorkbookHeader && (
                <div className={CSS.HEADER}>
                    <span className="eww-header__title">📊 {workbookNameValue}</span>
                    <div className="eww-header__meta">
                        <SavingIndicator status={savingStatus} />
                        {!activeSheetEditable && activeSheet && <ReadOnlyBadge />}
                    </div>
                </div>
            )}

            {showToolbar && (
                <Toolbar
                    hotRef={hotRef}
                    activeSheet={activeSheet}
                    onMetaChange={handleMetaChange}
                    disabled={!activeSheetEditable}
                    isAdmin={isAdminValue}
                    onOpenColumnSettings={() => setShowColumnPanel(true)}
                />
            )}

            <div className={CSS.GRID_WRAPPER}>
                {activeSheet && (
                    <SheetGrid
                        key={activeSheet.sheetId}
                        sheet={activeSheet}
                        isEditable={activeSheetEditable}
                        isAdmin={isAdminValue}
                        height={gridHeight}
                        rowHeaders={rowHeaders}
                        colHeaders={colHeaders}
                        hotRef={hotRef}
                        onCellChange={handleCellChange}
                        onMetaChange={handleMetaChange}
                        onDimensionChange={handleDimensionChange}
                    />
                )}
            </div>

            <SheetTabBar
                sheets={sheets}
                activeIndex={activeSheetIndex}
                isWorkbookEditable={!isReadOnlyValue}
                canEditSheet={canEditSheet}
                onTabChange={handleTabChange}
                onAddSheet={handleAddSheet}
                onDeleteSheet={handleDeleteSheet}
                onRenameSheet={handleRenameSheet}
            />

            {showColumnPanel && isAdminValue && (
                <ColumnSettingsPanel
                    sheet={activeSheet}
                    isAdmin={isAdminValue}
                    onAddColumn={handleAddColumn}
                    onUpdateColumn={handleUpdateColumn}
                    onDeleteColumn={handleDeleteColumn}
                    onReorderColumn={handleReorderColumn}
                    onClose={() => setShowColumnPanel(false)}
                />
            )}

        </div>
    );
}

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

function resolveAttr(prop) {
    if (prop === null || prop === undefined) return undefined;
    if (typeof prop === "object" && "status" in prop) {
        return prop.status === "available" ? prop.value : undefined;
    }
    return prop;
}

const styles = {
    errorBox:   { padding: 16, background: "#fce8e6", border: "1px solid #f5c6c6", borderRadius: 6, color: "#c5221f", fontSize: 13 },
    loadingBox: { display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, padding: 40, color: "#5f6368", fontSize: 14 },
    emptyBox:   { display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 8, padding: 48, color: "#5f6368", fontSize: 13 },
    emptyIcon:  { fontSize: 32 },
    spinner:    { width: 28, height: 28, border: "3px solid #e0e0e0", borderTopColor: "#1a73e8", borderRadius: "50%", animation: "eww-spin 0.75s linear infinite" },
};
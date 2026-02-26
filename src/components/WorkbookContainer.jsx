/**
 * WorkbookContainer.jsx
 *
 * Root component of the widget. Owns all state and orchestrates
 * all child components.
 *
 * RENDERS:
 *   ┌────────────────────────────────────┐
 *   │ WorkbookHeader (name + save status)│  ← only if showWorkbookHeader
 *   ├────────────────────────────────────┤
 *   │ Toolbar (bold, color, merge...)    │  ← only if showToolbar + editable
 *   ├────────────────────────────────────┤
 *   │                                    │
 *   │ SheetGrid (HotTable)               │  ← active sheet only
 *   │                                    │
 *   ├────────────────────────────────────┤
 *   │ SheetTabBar (Sheet1 | Sheet2 | ...) │
 *   └────────────────────────────────────┘
 *
 * DATA FLOW (read):
 *   Mendix sheetsJson prop
 *     → useWorkbookState (parse + state)
 *       → SheetGrid (renders active sheet)
 *
 * DATA FLOW (write):
 *   User edits cell in SheetGrid
 *     → onCellChange callback
 *       → useWorkbookState.setSheets (update local state)
 *         → useAutoSave (debounce → serialize → fire Mendix action)
 */

import { createElement, useRef, useCallback } from "react";

import { useWorkbookState }  from "../hooks/useWorkbookState";
import { useAutoSave }       from "../hooks/useAutoSave";
import { usePermissions }    from "../hooks/usePermissions";

import { SheetGrid }         from "./SheetGrid";
import { SheetTabBar }       from "./SheetTabBar";
import { Toolbar }           from "./Toolbar";
import { ReadOnlyBadge }     from "./ReadOnlyBadge";

import { updateSheetData, updateSheetMeta, updateSheetDimensions, addSheet, deleteSheet, renameSheet } from "../services/dataService";
import { CSS }               from "../utils/constants";
import { triggerSheetTabChange } from "../services/mendixBridge";

/**
 * @param {object} props - all props from ExcelWorkbookWidget.jsx (from Mendix XML)
 */
export function WorkbookContainer(props) {
    const {
        // Workbook identity
        workbookId,
        workbookName,

        // Data
        sheetsJson,

        // Permissions
        isReadOnly,

        // Actions
        onSheetChange,
        onSheetTabChange,

        // Display settings
        gridHeight       = 600,
        showToolbar      = true,
        showWorkbookHeader = true,
        rowHeaders       = true,
        colHeaders       = true,
    } = props;

    // ── Resolve Mendix attribute values from datasource ───────────────────
    // With datasource pattern, attributes are ListAttributeValue objects.
    // We call .get(item) on them to get the EditableValue for that item.
    const workbookItem   = props.workbookSource?.items?.[0];

    const sheetsJsonAttr    = workbookItem && sheetsJson    ? sheetsJson.get(workbookItem)    : sheetsJson;
    const isReadOnlyAttr    = workbookItem && isReadOnly    ? isReadOnly.get(workbookItem)    : isReadOnly;
    const workbookIdAttr    = workbookItem && workbookId    ? workbookId.get(workbookItem)    : workbookId;
    const workbookNameAttr  = workbookItem && workbookName  ? workbookName.get(workbookItem)  : workbookName;

    const sheetsJsonValue   = resolveAttr(sheetsJsonAttr);
    const isReadOnlyValue   = resolveAttr(isReadOnlyAttr)   ?? false;
    const workbookIdValue   = resolveAttr(workbookIdAttr)   ?? "";
    const workbookNameValue = resolveAttr(workbookNameAttr) ?? "Workbook";

    // ── Ref to SheetGrid (for toolbar commands) ────────────────────────────
    // Toolbar buttons need to call HotTable methods directly (bold, color etc.)
    // We pass this ref down to SheetGrid and it attaches to the HotTable instance.
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
        sheetsJson: sheetsJsonAttr,   // EditableValue with setValue — needed by mendixBridge
        clearPendingEdits,
    });

    // ── Cell change handler ────────────────────────────────────────────────
    const handleCellChange = useCallback((sheetId, newData) => {
        markPendingEdits();
        setSheets(prev => updateSheetData(prev, sheetId, newData));
    }, [markPendingEdits, setSheets]);

    // ── Cell meta change handler (formatting from Toolbar) ─────────────────
    const handleMetaChange = useCallback((sheetId, newMeta) => {
        markPendingEdits();
        setSheets(prev => updateSheetMeta(prev, sheetId, newMeta));
    }, [markPendingEdits, setSheets]);

    // ── Dimension change handler (col/row resize) ──────────────────────────
    const handleDimensionChange = useCallback((sheetId, dimensions) => {
        markPendingEdits();
        setSheets(prev => updateSheetDimensions(prev, sheetId, dimensions));
    }, [markPendingEdits, setSheets]);

    // ── Add new empty sheet ────────────────────────────────────────────────
    const handleAddSheet = useCallback(() => {
        markPendingEdits();
        setSheets(prev => addSheet(prev));
        // Switch to the new tab (it will be the last one)
        setActiveSheetIndex(sheets.length); // current length = new sheet's index
    }, [markPendingEdits, setSheets, setActiveSheetIndex, sheets.length]);

    // ── Delete a sheet ─────────────────────────────────────────────────────
    const handleDeleteSheet = useCallback((sheetId) => {
        markPendingEdits();
        setSheets(prev => deleteSheet(prev, sheetId));
        // If we deleted the active sheet, move to the previous tab
        setActiveSheetIndex(Math.max(0, activeSheetIndex - 1));
    }, [markPendingEdits, setSheets, setActiveSheetIndex, activeSheetIndex]);

    // ── Rename a sheet ─────────────────────────────────────────────────────
    const handleRenameSheet = useCallback((sheetId, newName) => {
        markPendingEdits();
        setSheets(prev => renameSheet(prev, sheetId, newName));
    }, [markPendingEdits, setSheets]);

    // ── Tab change handler ─────────────────────────────────────────────────
    const handleTabChange = useCallback((index) => {
        setActiveSheetIndex(index);
        triggerSheetTabChange(onSheetTabChange);
    }, [setActiveSheetIndex, onSheetTabChange]);

    // ── Render: error state ────────────────────────────────────────────────
    if (parseError) {
        return (
            <div className={CSS.WORKBOOK_ROOT} style={styles.errorBox}>
                ⚠ Failed to load workbook data. Please check the sheetsJson configuration.
                <br />
                <small style={{ color: "#999", marginTop: 4, display: "block" }}>
                    {parseError}
                </small>
            </div>
        );
    }

    // ── Render: loading state ──────────────────────────────────────────────
    if (isLoading) {
        return (
            <div className={CSS.WORKBOOK_ROOT} style={styles.loadingBox}>
                <div style={styles.spinner} />
                <span>Loading workbook…</span>
            </div>
        );
    }

    // ── Render: no sheets configured ──────────────────────────────────────
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

    // ── Resolve active sheet editability ──────────────────────────────────
    const activeSheetEditable = activeSheet
        ? canEditSheet(activeSheet.isEditable)
        : false;

    // ── Main render ───────────────────────────────────────────────────────
    return (
        <div className={CSS.WORKBOOK_ROOT}>

            {/* ── Workbook Header ──────────────────────────────────────── */}
            {showWorkbookHeader && (
                <div className={CSS.HEADER}>
                    <span className="eww-header__title">
                        📊 {workbookNameValue}
                    </span>
                    <div className="eww-header__meta">
                        <SavingIndicator status={savingStatus} />
                        {!activeSheetEditable && activeSheet && (
                            <ReadOnlyBadge />
                        )}
                    </div>
                </div>
            )}

            {/* ── Toolbar ──────────────────────────────────────────────── */}
            {showToolbar && (
                <Toolbar
                    hotRef={hotRef}
                    activeSheet={activeSheet}
                    onMetaChange={handleMetaChange}
                    disabled={!activeSheetEditable}
                />
            )}

            {/* ── Sheet Grid ───────────────────────────────────────────── */}
            <div className={CSS.GRID_WRAPPER}>
                {activeSheet && (
                    <SheetGrid
                        key={activeSheet.sheetId}
                        sheet={activeSheet}
                        isEditable={activeSheetEditable}
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

            {/* ── Sheet Tab Bar ─────────────────────────────────────────── */}
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

        </div>
    );
}

// ─── Saving Indicator (inline — small enough to not need its own file) ────────

function SavingIndicator({ status }) {
    if (status === "idle") return null;

    const isSaving = status === "saving";
    const className = [
        CSS.SAVING_INDICATOR,
        isSaving
            ? "eww-saving-indicator--saving"
            : "eww-saving-indicator--saved",
    ].join(" ");

    return (
        <span className={className}>
            {isSaving
                ? createElement("span", null,
                    createElement("span", { className: "eww-saving-indicator__dot" }),
                    "Saving…"
                  )
                : "✓ Saved"
            }
        </span>
    );
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

/**
 * resolveAttr
 *
 * Mendix passes props as EditableValue objects ({ status, value, setValue }).
 * This helper extracts the plain .value safely.
 * If the prop is already a plain value (string, boolean, number), returns as-is.
 *
 * @param   {any} prop  - Mendix EditableValue or plain value
 * @returns {any}       - the resolved plain value, or undefined
 */
function resolveAttr(prop) {
    if (prop === null || prop === undefined) return undefined;
    // Mendix EditableValue shape
    if (typeof prop === "object" && "status" in prop) {
        return prop.status === "available" ? prop.value : undefined;
    }
    // Already a plain value (e.g. integer/boolean props from XML)
    return prop;
}

// ─── Inline styles (structural only — visual styles are in CSS) ───────────────

const styles = {
    errorBox: {
        padding:      16,
        background:   "#fce8e6",
        border:       "1px solid #f5c6c6",
        borderRadius: 6,
        color:        "#c5221f",
        fontSize:     13,
    },
    loadingBox: {
        display:        "flex",
        flexDirection:  "column",
        alignItems:     "center",
        justifyContent: "center",
        gap:            10,
        padding:        40,
        color:          "#5f6368",
        fontSize:       14,
    },
    emptyBox: {
        display:        "flex",
        flexDirection:  "column",
        alignItems:     "center",
        justifyContent: "center",
        gap:            8,
        padding:        48,
        color:          "#5f6368",
        fontSize:       13,
    },
    emptyIcon: {
        fontSize: 32,
    },
    spinner: {
        width:        28,
        height:       28,
        border:       "3px solid #e0e0e0",
        borderTopColor: "#1a73e8",
        borderRadius: "50%",
        animation:    "eww-spin 0.75s linear infinite",
    },
};
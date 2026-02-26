/**
 * ColumnSettingsPanel.jsx
 *
 * Admin-only panel for managing column definitions.
 * Appears as a slide-in panel triggered by a "⚙ Columns" button in the Toolbar.
 *
 * WHAT ADMIN CAN DO:
 *   - Add new column
 *   - Edit column header name
 *   - Change column data type (text, numeric, date, checkbox, dropdown, time)
 *   - For dropdown type: edit the source options list
 *   - Delete a column
 *   - Reorder columns (drag not needed — up/down arrows)
 *
 * WHAT USERS SEE:
 *   - Nothing — this panel is only rendered when isAdmin = true
 *
 * DATA FLOW:
 *   Admin edits → onColumnChange(sheetId, updatedColumns) 
 *   → WorkbookContainer.setSheets → auto-save persists to Mendix
 */

import { createElement, useState, useCallback } from "react";
import { COLUMN_TYPE_META } from "../utils/constants";

export function ColumnSettingsPanel({
    sheet,
    isAdmin,
    onAddColumn,
    onUpdateColumn,
    onDeleteColumn,
    onReorderColumn,
    onClose,
}) {
    if (!isAdmin || !sheet) return null;

    const columns    = sheet.columns || [];
    // Which column is expanded (showing type/source editor)
    const [expandedKey, setExpandedKey] = useState(null);

    const toggleExpand = useCallback((key) => {
        setExpandedKey(prev => prev === key ? null : key);
    }, []);

    return (
        <div style={panelStyles.overlay} onClick={onClose}>
            <div style={panelStyles.panel} onClick={e => e.stopPropagation()}>

                {/* ── Header ───────────────────────────────────────────── */}
                <div style={panelStyles.header}>
                    <span style={panelStyles.title}>⚙ Column Settings</span>
                    <span style={panelStyles.subtitle}>
                        {sheet.sheetName} — Admin only
                    </span>
                    <button onClick={onClose} style={panelStyles.closeBtn} title="Close">✕</button>
                </div>

                {/* ── Column List ──────────────────────────────────────── */}
                <div style={panelStyles.columnList}>
                    {columns.length === 0 && (
                        <div style={panelStyles.emptyHint}>
                            No columns defined yet. Click "+ Add Column" to start.
                            <br />
                            <small style={{ color: "#999" }}>
                                Until columns are defined, the grid uses default A, B, C headers.
                            </small>
                        </div>
                    )}

                    {columns.map((col, index) => (
                        <ColumnRow
                            key={col.key}
                            col={col}
                            index={index}
                            total={columns.length}
                            isExpanded={expandedKey === col.key}
                            onToggleExpand={() => toggleExpand(col.key)}
                            onUpdate={(changes) => onUpdateColumn(sheet.sheetId, col.key, changes)}
                            onDelete={() => onDeleteColumn(sheet.sheetId, col.key)}
                            onMoveUp={() => index > 0 && onReorderColumn(sheet.sheetId, index, index - 1)}
                            onMoveDown={() => index < columns.length - 1 && onReorderColumn(sheet.sheetId, index, index + 1)}
                        />
                    ))}
                </div>

                {/* ── Add Column Button ────────────────────────────────── */}
                <div style={panelStyles.footer}>
                    <button
                        style={panelStyles.addBtn}
                        onClick={() => onAddColumn(sheet.sheetId)}
                    >
                        + Add Column
                    </button>
                </div>

            </div>
        </div>
    );
}

// ─── ColumnRow ────────────────────────────────────────────────────────────────

function ColumnRow({ col, index, total, isExpanded, onToggleExpand, onUpdate, onDelete, onMoveUp, onMoveDown }) {

    const typeMeta = COLUMN_TYPE_META.find(t => t.value === col.type) || COLUMN_TYPE_META[0];

    return (
        <div style={rowStyles.wrapper}>

            {/* ── Row summary line ──────────────────────────────────── */}
            <div style={rowStyles.summary}>

                {/* Order buttons */}
                <div style={rowStyles.orderBtns}>
                    <button
                        onClick={onMoveUp}
                        disabled={index === 0}
                        style={rowStyles.arrowBtn}
                        title="Move up"
                    >▲</button>
                    <button
                        onClick={onMoveDown}
                        disabled={index === total - 1}
                        style={rowStyles.arrowBtn}
                        title="Move down"
                    >▼</button>
                </div>

                {/* Type icon */}
                <span style={rowStyles.typeIcon} title={typeMeta.label}>
                    {typeMeta.icon}
                </span>

                {/* Header name (inline edit) */}
                <input
                    value={col.header}
                    onChange={e => onUpdate({ header: e.target.value })}
                    style={rowStyles.headerInput}
                    placeholder="Column name"
                    maxLength={60}
                    onClick={e => e.stopPropagation()}
                />

                {/* Type selector */}
                <select
                    value={col.type}
                    onChange={e => onUpdate({ type: e.target.value, source: [], format: "" })}
                    style={rowStyles.typeSelect}
                >
                    {COLUMN_TYPE_META.map(t => (
                        <option key={t.value} value={t.value}>{t.label}</option>
                    ))}
                </select>

                {/* Expand for advanced settings */}
                {(typeMeta.hasSource || typeMeta.hasFormat) && (
                    <button
                        onClick={onToggleExpand}
                        style={rowStyles.expandBtn}
                        title={isExpanded ? "Collapse" : "More settings"}
                    >
                        {isExpanded ? "▴" : "▾"}
                    </button>
                )}

                {/* Delete */}
                <button
                    onClick={onDelete}
                    style={rowStyles.deleteBtn}
                    title={`Delete column "${col.header}"`}
                >
                    ✕
                </button>
            </div>

            {/* ── Expanded settings ─────────────────────────────────── */}
            {isExpanded && (
                <div style={rowStyles.expanded}>

                    {/* Dropdown source list */}
                    {typeMeta.hasSource && (
                        <div style={rowStyles.fieldGroup}>
                            <label style={rowStyles.fieldLabel}>
                                Dropdown options (one per line):
                            </label>
                            <textarea
                                value={(col.source || []).join("\n")}
                                onChange={e => onUpdate({
                                    source: e.target.value
                                        .split("\n")
                                        .map(s => s.trim())
                                        .filter(Boolean)
                                })}
                                style={rowStyles.textarea}
                                placeholder={"Option 1\nOption 2\nOption 3"}
                                rows={4}
                            />
                        </div>
                    )}

                    {/* Format string */}
                    {typeMeta.hasFormat && (
                        <div style={rowStyles.fieldGroup}>
                            <label style={rowStyles.fieldLabel}>
                                {col.type === "date" ? "Date format:" : "Number format:"}
                            </label>
                            <input
                                value={col.format || ""}
                                onChange={e => onUpdate({ format: e.target.value })}
                                style={rowStyles.formatInput}
                                placeholder={col.type === "date" ? "DD/MM/YYYY" : "0,0.00"}
                            />
                            <small style={{ color: "#999", marginTop: 2, display: "block" }}>
                                {col.type === "date"
                                    ? "e.g. DD/MM/YYYY or MM-DD-YYYY or YYYY/MM/DD"
                                    : "e.g. 0,0.00 or 0% or $0,0"}
                            </small>
                        </div>
                    )}

                    {/* Read-only toggle */}
                    <div style={rowStyles.fieldGroup}>
                        <label style={{ ...rowStyles.fieldLabel, display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
                            <input
                                type="checkbox"
                                checked={col.readOnly || false}
                                onChange={e => onUpdate({ readOnly: e.target.checked })}
                            />
                            Lock this column (users cannot edit values)
                        </label>
                    </div>

                </div>
            )}
        </div>
    );
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const panelStyles = {
    overlay: {
        position:   "fixed",
        top: 0, left: 0, right: 0, bottom: 0,
        background: "rgba(0,0,0,0.35)",
        zIndex:     9999,
        display:    "flex",
        alignItems: "flex-start",
        justifyContent: "flex-end",
        paddingTop: 60,
        paddingRight: 16,
    },
    panel: {
        background:   "#fff",
        borderRadius: 8,
        boxShadow:    "0 8px 32px rgba(0,0,0,0.18)",
        width:        420,
        maxHeight:    "80vh",
        display:      "flex",
        flexDirection:"column",
        overflow:     "hidden",
    },
    header: {
        padding:        "14px 16px",
        borderBottom:   "1px solid #e0e0e0",
        display:        "flex",
        alignItems:     "center",
        gap:            8,
        background:     "#f8f9fa",
        flexShrink:     0,
    },
    title: {
        fontWeight: 600,
        fontSize:   14,
        color:      "#1a1a1a",
        flexShrink: 0,
    },
    subtitle: {
        fontSize:  12,
        color:     "#888",
        flex:      1,
        overflow:  "hidden",
        textOverflow: "ellipsis",
        whiteSpace:   "nowrap",
    },
    closeBtn: {
        background: "none",
        border:     "none",
        cursor:     "pointer",
        fontSize:   14,
        color:      "#5f6368",
        padding:    "0 4px",
        flexShrink: 0,
    },
    columnList: {
        overflowY:  "auto",
        flex:       1,
        padding:    "8px 0",
    },
    emptyHint: {
        padding:   "20px 16px",
        color:     "#5f6368",
        fontSize:  13,
        lineHeight: 1.5,
    },
    footer: {
        padding:     "10px 16px",
        borderTop:   "1px solid #e0e0e0",
        flexShrink:  0,
        background:  "#f8f9fa",
    },
    addBtn: {
        background:   "#1a73e8",
        color:        "#fff",
        border:       "none",
        borderRadius: 4,
        padding:      "7px 16px",
        cursor:       "pointer",
        fontSize:     13,
        fontWeight:   500,
        width:        "100%",
    },
};

const rowStyles = {
    wrapper: {
        borderBottom: "1px solid #f0f0f0",
        padding:      "0 8px",
    },
    summary: {
        display:    "flex",
        alignItems: "center",
        gap:        6,
        padding:    "6px 0",
    },
    orderBtns: {
        display:       "flex",
        flexDirection: "column",
        gap:           1,
        flexShrink:    0,
    },
    arrowBtn: {
        background: "none",
        border:     "none",
        cursor:     "pointer",
        fontSize:   8,
        padding:    "1px 3px",
        color:      "#888",
        lineHeight: 1,
    },
    typeIcon: {
        fontSize:   12,
        width:      18,
        textAlign:  "center",
        flexShrink: 0,
        color:      "#1a73e8",
        fontWeight: 700,
    },
    headerInput: {
        flex:         1,
        border:       "1px solid #e0e0e0",
        borderRadius: 3,
        padding:      "3px 6px",
        fontSize:     13,
        outline:      "none",
        minWidth:     0,
    },
    typeSelect: {
        border:       "1px solid #e0e0e0",
        borderRadius: 3,
        padding:      "3px 4px",
        fontSize:     12,
        flexShrink:   0,
        background:   "#fff",
        cursor:       "pointer",
    },
    expandBtn: {
        background: "none",
        border:     "none",
        cursor:     "pointer",
        fontSize:   12,
        color:      "#1a73e8",
        padding:    "0 4px",
        flexShrink: 0,
    },
    deleteBtn: {
        background: "none",
        border:     "none",
        cursor:     "pointer",
        fontSize:   11,
        color:      "#c5221f",
        padding:    "0 4px",
        flexShrink: 0,
    },
    expanded: {
        padding:      "8px 8px 12px 32px",
        background:   "#fafafa",
        borderTop:    "1px solid #f0f0f0",
        borderRadius: "0 0 4px 4px",
    },
    fieldGroup: {
        marginBottom: 10,
    },
    fieldLabel: {
        fontSize:     12,
        color:        "#5f6368",
        marginBottom: 4,
        display:      "block",
        fontWeight:   500,
    },
    textarea: {
        width:        "100%",
        border:       "1px solid #e0e0e0",
        borderRadius: 3,
        padding:      "4px 6px",
        fontSize:     12,
        fontFamily:   "inherit",
        resize:       "vertical",
        boxSizing:    "border-box",
        outline:      "none",
    },
    formatInput: {
        width:        "100%",
        border:       "1px solid #e0e0e0",
        borderRadius: 3,
        padding:      "4px 6px",
        fontSize:     12,
        outline:      "none",
        boxSizing:    "border-box",
    },
};
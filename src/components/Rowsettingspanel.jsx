/**
 * RowSettingsPanel.jsx
 *
 * Admin-only panel for managing custom row labels.
 * Mirrors ColumnSettingsPanel exactly:
 *   - Starts EMPTY — no rows pre-populated
 *   - Admin adds labels one by one via "+ Add Row"
 *   - Deleting all labels → grid falls back to default row numbers (1, 2, 3…)
 */

import { createElement, useCallback } from "react";

export function RowSettingsPanel({
    sheet,
    isAdmin,
    onAddRow,
    onUpdateRow,
    onDeleteRow,
    onReorderRow,
    onClose,
}) {
    if (!isAdmin || !sheet) return null;

    // Only show labels that have been explicitly added by the admin.
    // We do NOT auto-populate from data rows — admin adds them manually,
    // exactly like column config.
    const rowLabels = sheet.rowLabels || [];

    return (
        <div style={panelStyles.overlay} onClick={onClose}>
            <div style={panelStyles.panel} onClick={e => e.stopPropagation()}>

                {/* ── Header ───────────────────────────────────────────── */}
                <div style={panelStyles.header}>
                    <span style={panelStyles.title}>☰ Row Labels</span>
                    <span style={panelStyles.subtitle}>
                        {sheet.sheetName} — Admin only
                    </span>
                    <button onClick={onClose} style={panelStyles.closeBtn} title="Close">✕</button>
                </div>

                {/* ── Info hint ────────────────────────────────────────── */}
                <div style={panelStyles.hint}>
                    <span style={panelStyles.hintIcon}>ℹ</span>
                    <span>
                        Add a label for each row you want to name.
                        Rows with no label will show their row number.
                        Delete all labels to restore default numbering.
                    </span>
                </div>

                {/* ── Label List ───────────────────────────────────────── */}
                <div style={panelStyles.labelList}>
                    {rowLabels.length === 0 && (
                        <div style={panelStyles.emptyHint}>
                            No row labels defined yet. Click "+ Add Row Label" to start.
                            <br />
                            <small style={{ color: "#999", marginTop: 4, display: "block" }}>
                                Until labels are added, the grid shows default row numbers.
                            </small>
                        </div>
                    )}

                    {rowLabels.map((label, index) => (
                        <RowLabelItem
                            key={index}
                            index={index}
                            label={label}
                            total={rowLabels.length}
                            onUpdate={(newLabel) => onUpdateRow(index, newLabel)}
                            onDelete={() => onDeleteRow(index)}
                            onMoveUp={() => index > 0 && onReorderRow(index, index - 1)}
                            onMoveDown={() => index < rowLabels.length - 1 && onReorderRow(index, index + 1)}
                        />
                    ))}
                </div>

                {/* ── Footer ───────────────────────────────────────────── */}
                <div style={panelStyles.footer}>
                    <button style={panelStyles.addBtn} onClick={onAddRow}>
                        + Add Row Label
                    </button>
                </div>

            </div>
        </div>
    );
}

// ─── RowLabelItem ─────────────────────────────────────────────────────────────

function RowLabelItem({ index, label, total, onUpdate, onDelete, onMoveUp, onMoveDown }) {
    return (
        <div style={itemStyles.wrapper}>

            {/* Row position badge */}
            <span style={itemStyles.rowNumber} title={`Applies to row ${index + 1}`}>
                {index + 1}
            </span>

            {/* Up / Down */}
            <div style={itemStyles.orderBtns}>
                <button onClick={onMoveUp}   disabled={index === 0}           style={itemStyles.arrowBtn} title="Move up">▲</button>
                <button onClick={onMoveDown} disabled={index === total - 1}   style={itemStyles.arrowBtn} title="Move down">▼</button>
            </div>

            {/* Label — plain string input only */}
            <input
                value={label}
                onChange={e => onUpdate(e.target.value)}
                placeholder={`Row ${index + 1} label…`}
                style={itemStyles.labelInput}
                maxLength={80}
                type="text"
            />

            {/* Delete */}
            <button onClick={onDelete} style={itemStyles.deleteBtn} title={`Remove label for row ${index + 1}`}>
                ✕
            </button>

        </div>
    );
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const panelStyles = {
    overlay: {
        position: "fixed", top: 0, left: 0, right: 0, bottom: 0,
        background: "rgba(0,0,0,0.35)", zIndex: 9999,
        display: "flex", alignItems: "flex-start", justifyContent: "flex-end",
        paddingTop: 60, paddingRight: 16,
    },
    panel: {
        background: "#fff", borderRadius: 8,
        boxShadow: "0 8px 32px rgba(0,0,0,0.18)",
        width: 380, maxHeight: "80vh",
        display: "flex", flexDirection: "column", overflow: "hidden",
    },
    header: {
        padding: "14px 16px", borderBottom: "1px solid #e0e0e0",
        display: "flex", alignItems: "center", gap: 8,
        background: "#f8f9fa", flexShrink: 0,
    },
    title:    { fontWeight: 600, fontSize: 14, color: "#1a1a1a", flexShrink: 0 },
    subtitle: { fontSize: 12, color: "#888", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
    closeBtn: { background: "none", border: "none", cursor: "pointer", fontSize: 14, color: "#5f6368", padding: "0 4px", flexShrink: 0 },
    hint: {
        display: "flex", alignItems: "flex-start", gap: 6,
        padding: "8px 16px", background: "#f0f4ff",
        borderBottom: "1px solid #dde6ff",
        fontSize: 12, color: "#3c4ca0", flexShrink: 0, lineHeight: 1.5,
    },
    hintIcon:  { fontWeight: 700, flexShrink: 0, marginTop: 1 },
    labelList: { overflowY: "auto", flex: 1, padding: "8px 0" },
    emptyHint: { padding: "20px 16px", color: "#5f6368", fontSize: 13, lineHeight: 1.6 },
    footer:    { padding: "10px 16px", borderTop: "1px solid #e0e0e0", flexShrink: 0, background: "#f8f9fa" },
    addBtn: {
        background: "#1a73e8", color: "#fff", border: "none",
        borderRadius: 4, padding: "7px 16px", cursor: "pointer",
        fontSize: 13, fontWeight: 500, width: "100%",
    },
};

const itemStyles = {
    wrapper:   { display: "flex", alignItems: "center", gap: 6, padding: "5px 12px", borderBottom: "1px solid #f0f0f0" },
    rowNumber: {
        width: 24, height: 24, display: "inline-flex",
        alignItems: "center", justifyContent: "center",
        background: "#f1f3f4", borderRadius: "50%",
        fontSize: 11, fontWeight: 600, color: "#5f6368", flexShrink: 0,
    },
    orderBtns: { display: "flex", flexDirection: "column", gap: 1, flexShrink: 0 },
    arrowBtn:  { background: "none", border: "none", cursor: "pointer", fontSize: 8, padding: "1px 3px", color: "#888", lineHeight: 1 },
    labelInput: {
        flex: 1, border: "1px solid #e0e0e0", borderRadius: 3,
        padding: "4px 8px", fontSize: 13, outline: "none",
        minWidth: 0, fontFamily: "inherit",
    },
    deleteBtn: { background: "none", border: "none", cursor: "pointer", fontSize: 11, color: "#c5221f", padding: "0 4px", flexShrink: 0 },
};
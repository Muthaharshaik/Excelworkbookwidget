/**
 * RowSettingsPanel.jsx — Redesigned UI
 * No JSX fragments — Mendix compatibility.
 */

import { createElement, useCallback } from "react";

export function RowSettingsPanel({
    sheet, isAdmin,
    onAddRow, onUpdateRow, onDeleteRow, onReorderRow, onClose,
}) {
    if (!isAdmin || !sheet) return null;

    const rowLabels = sheet.rowLabels || [];

    // Wrapper div has no visual effect — both children are position:fixed
    return (
        <div style={{ display: "contents" }}>

            {/* Backdrop */}
            <div style={S.backdrop} onClick={onClose} />

            {/* Panel */}
            <div style={S.panel}>

                {/* Header */}
                <div style={S.header}>
                    <div style={S.headerLeft}>
                        <div style={S.headerIcon}>
                            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                                <rect x="1" y="1"  width="14" height="4" rx="1.5" fill="#059669"/>
                                <rect x="1" y="6"  width="14" height="4" rx="1.5" fill="#059669" opacity="0.35"/>
                                <rect x="1" y="11" width="14" height="4" rx="1.5" fill="#059669" opacity="0.15"/>
                            </svg>
                        </div>
                        <div>
                            <div style={S.headerTitle}>Row Labels</div>
                            <div style={S.headerSub}>{sheet.sheetName} · Admin only</div>
                        </div>
                    </div>
                    <button onClick={onClose} style={S.closeBtn} title="Close">
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                            <path d="M1 1l12 12M13 1L1 13" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round"/>
                        </svg>
                    </button>
                </div>

                {/* Info bar */}
                <div style={S.infoBar}>
                    <svg width="14" height="14" viewBox="0 0 14 14" fill="none" style={{ flexShrink: 0, marginTop: 1 }}>
                        <circle cx="7" cy="7" r="6" stroke="#0d9488" strokeWidth="1.5"/>
                        <path d="M7 6v4M7 4.5v.5" stroke="#0d9488" strokeWidth="1.5" strokeLinecap="round"/>
                    </svg>
                    <span>
                        Add labels for rows you want to name. Only labelled rows are shown in the grid.
                        Delete all to restore default row numbers.
                    </span>
                </div>

                {/* Count bar */}
                <div style={S.countBar}>
                    <span style={S.countPill}>
                        {rowLabels.length} {rowLabels.length === 1 ? "row" : "rows"}
                    </span>
                    <span style={S.countHint}>
                        {rowLabels.length === 0 ? "Default 1, 2, 3… numbering active" : "Custom row labels active"}
                    </span>
                </div>

                {/* List */}
                <div style={S.list}>
                    {rowLabels.length === 0 && (
                        <div style={S.empty}>
                            <div style={S.emptyIcon}>☰</div>
                            <div style={S.emptyTitle}>No row labels yet</div>
                            <div style={S.emptyDesc}>
                                Add row labels to replace the default 1, 2, 3 numbers
                                with meaningful names like "Q1", "Revenue", "Jan".
                            </div>
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

                {/* Footer */}
                <div style={S.footer}>
                    <button style={S.addBtn} onClick={onAddRow}>
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                            <path d="M7 1v12M1 7h12" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                        </svg>
                        Add Row Label
                    </button>
                </div>

            </div>
        </div>
    );
}

function RowLabelItem({ index, label, total, onUpdate, onDelete, onMoveUp, onMoveDown }) {
    return (
        <div style={R.wrapper}>
            <div style={R.row}>

                <div style={R.orderCol}>
                    <button onClick={onMoveUp}   disabled={index === 0}         style={R.arrow} title="Move up">▲</button>
                    <span style={R.indexNum}>{index + 1}</span>
                    <button onClick={onMoveDown} disabled={index === total - 1} style={R.arrow} title="Move down">▼</button>
                </div>

                <div style={R.rowBadge}>
                    <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                        <rect x="1" y="4" width="10" height="4" rx="1" fill="#059669" opacity="0.7"/>
                    </svg>
                </div>

                <input
                    value={label}
                    onChange={e => onUpdate(e.target.value)}
                    placeholder={`Row ${index + 1} label…`}
                    style={R.labelInput}
                    maxLength={80}
                    type="text"
                    autoFocus={label === "" && index === total - 1}
                />

                <button onClick={onDelete} style={R.deleteBtn} title={`Remove row ${index + 1} label`}>
                    <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                        <path d="M1 1l10 10M11 1L1 11" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round"/>
                    </svg>
                </button>

            </div>
        </div>
    );
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const S = {
    backdrop: {
        position: "fixed", inset: 0,
        background: "rgba(15,23,42,0.3)",
        backdropFilter: "blur(2px)",
        zIndex: 9998,
    },
    panel: {
        position: "fixed",
        top: 56, right: 16,
        width: 400,
        maxHeight: "calc(100vh - 80px)",
        background: "#ffffff",
        borderRadius: 12,
        boxShadow: "0 20px 60px rgba(15,23,42,0.18), 0 4px 16px rgba(15,23,42,0.08)",
        border: "1px solid rgba(226,232,240,0.8)",
        display: "flex", flexDirection: "column",
        overflow: "hidden",
        zIndex: 9999,
        animation: "eww-slideIn 0.2s cubic-bezier(0.16,1,0.3,1)",
    },
    header: {
        display: "flex", alignItems: "center", justifyContent: "space-between",
        padding: "16px 18px", borderBottom: "1px solid #f1f5f9",
        background: "linear-gradient(135deg, #f0fdf8 0%, #ecfdf5 100%)",
        flexShrink: 0,
    },
    headerLeft:  { display: "flex", alignItems: "center", gap: 12 },
    headerIcon:  {
        width: 36, height: 36, background: "#d1fae5",
        borderRadius: 8, border: "1px solid #a7f3d0",
        display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
    },
    headerTitle: { fontSize: 14, fontWeight: 700, color: "#0f172a", letterSpacing: "-0.01em" },
    headerSub:   { fontSize: 11, color: "#94a3b8", marginTop: 1 },
    closeBtn: {
        width: 28, height: 28, background: "#f1f5f9",
        border: "1px solid #e2e8f0", borderRadius: 6, cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center",
        color: "#64748b", flexShrink: 0,
    },
    infoBar: {
        display: "flex", alignItems: "flex-start", gap: 8,
        padding: "10px 18px", background: "#f0fdfa",
        borderBottom: "1px solid #ccfbf1",
        fontSize: 11, color: "#0f766e", lineHeight: 1.6, flexShrink: 0,
    },
    countBar: {
        display: "flex", alignItems: "center", gap: 10,
        padding: "10px 18px", borderBottom: "1px solid #f1f5f9",
        background: "#fafbfc", flexShrink: 0,
    },
    countPill: {
        display: "inline-flex", alignItems: "center",
        padding: "3px 10px", background: "#d1fae5",
        color: "#059669", border: "1px solid #a7f3d0",
        borderRadius: 20, fontSize: 11, fontWeight: 700,
    },
    countHint: { fontSize: 11, color: "#94a3b8" },
    list:      { overflowY: "auto", flex: 1 },
    empty: {
        display: "flex", flexDirection: "column",
        alignItems: "center", justifyContent: "center",
        padding: "36px 24px", textAlign: "center", gap: 8,
    },
    emptyIcon:  { fontSize: 32, marginBottom: 4 },
    emptyTitle: { fontSize: 14, fontWeight: 600, color: "#334155" },
    emptyDesc:  { fontSize: 12, color: "#94a3b8", lineHeight: 1.6, maxWidth: 260 },
    footer: {
        padding: "14px 18px", borderTop: "1px solid #f1f5f9",
        background: "#fafbfc", flexShrink: 0,
    },
    addBtn: {
        display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
        width: "100%", padding: "10px 0",
        background: "linear-gradient(135deg, #059669 0%, #047857 100%)",
        color: "#fff", border: "none", borderRadius: 8, cursor: "pointer",
        fontSize: 13, fontWeight: 600, letterSpacing: "0.01em",
        boxShadow: "0 2px 8px rgba(5,150,105,0.3)",
    },
};

const R = {
    wrapper:  { borderBottom: "1px solid #f8fafc" },
    row:      { display: "flex", alignItems: "center", gap: 8, padding: "10px 14px" },
    orderCol: { display: "flex", flexDirection: "column", alignItems: "center", gap: 1, flexShrink: 0 },
    arrow:    { background: "none", border: "none", cursor: "pointer", fontSize: 7, padding: "2px 3px", color: "#cbd5e1", lineHeight: 1 },
    indexNum: { fontSize: 9, fontWeight: 700, color: "#cbd5e1", lineHeight: 1 },
    rowBadge: {
        width: 28, height: 28, flexShrink: 0,
        background: "#d1fae5", border: "1px solid #a7f3d0",
        borderRadius: 6,
        display: "flex", alignItems: "center", justifyContent: "center",
    },
    labelInput: {
        flex: 1, border: "1px solid #e2e8f0", borderRadius: 6,
        padding: "6px 10px", fontSize: 13, color: "#0f172a",
        outline: "none", minWidth: 0, background: "#fff", fontFamily: "inherit",
    },
    deleteBtn: {
        width: 28, height: 28, flexShrink: 0,
        background: "#fff5f5", border: "1px solid #fecaca",
        borderRadius: 6, cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center", color: "#ef4444",
    },
};
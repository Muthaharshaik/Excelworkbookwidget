/**
 * ColumnSettingsPanel.jsx — Redesigned UI
 * No JSX fragments — Mendix compatibility.
 */

import { createElement, useState, useCallback } from "react";
import { COLUMN_TYPE_META } from "../utils/constants";

const TYPE_COLORS = {
    text:     { bg: "#eff6ff", color: "#2563eb", border: "#bfdbfe" },
    numeric:  { bg: "#f0fdf4", color: "#16a34a", border: "#bbf7d0" },
    date:     { bg: "#fff7ed", color: "#ea580c", border: "#fed7aa" },
    time:     { bg: "#fdf4ff", color: "#9333ea", border: "#e9d5ff" },
    checkbox: { bg: "#f0fdfa", color: "#0d9488", border: "#99f6e4" },
    dropdown: { bg: "#fefce8", color: "#ca8a04", border: "#fde68a" },
};

export function ColumnSettingsPanel({
    sheet, isAdmin,
    onAddColumn, onUpdateColumn, onDeleteColumn, onReorderColumn, onClose,
}) {
    if (!isAdmin || !sheet) return null;

    const columns = sheet.columns || [];
    const [expandedKey, setExpandedKey] = useState(null);

    const toggleExpand = useCallback((key) => {
        setExpandedKey(prev => prev === key ? null : key);
    }, []);

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
                                <rect x="1" y="1" width="6" height="14" rx="1.5" fill="#2563eb" opacity="0.15"/>
                                <rect x="9" y="1" width="6" height="14" rx="1.5" fill="#2563eb" opacity="0.15"/>
                                <rect x="1" y="1" width="6" height="4" rx="1.5" fill="#2563eb"/>
                                <rect x="9" y="1" width="6" height="4" rx="1.5" fill="#2563eb"/>
                            </svg>
                        </div>
                        <div>
                            <div style={S.headerTitle}>Column Settings</div>
                            <div style={S.headerSub}>{sheet.sheetName} · Admin only</div>
                        </div>
                    </div>
                    <button onClick={onClose} style={S.closeBtn} title="Close">
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                            <path d="M1 1l12 12M13 1L1 13" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round"/>
                        </svg>
                    </button>
                </div>

                {/* Count bar */}
                <div style={S.countBar}>
                    <span style={S.countPill}>
                        {columns.length} {columns.length === 1 ? "column" : "columns"}
                    </span>
                    <span style={S.countHint}>
                        {columns.length === 0 ? "Default A–Z headers active" : "Custom headers active"}
                    </span>
                </div>

                {/* List */}
                <div style={S.list}>
                    {columns.length === 0 && (
                        <div style={S.empty}>
                            <div style={S.emptyIcon}>⊞</div>
                            <div style={S.emptyTitle}>No columns configured</div>
                            <div style={S.emptyDesc}>Add columns to replace the default A, B, C headers with custom names and types.</div>
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

                {/* Footer */}
                <div style={S.footer}>
                    <button style={S.addBtn} onClick={() => onAddColumn(sheet.sheetId)}>
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                            <path d="M7 1v12M1 7h12" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                        </svg>
                        Add Column
                    </button>
                </div>

            </div>
        </div>
    );
}

function ColumnRow({ col, index, total, isExpanded, onToggleExpand, onUpdate, onDelete, onMoveUp, onMoveDown }) {
    const typeMeta   = COLUMN_TYPE_META.find(t => t.value === col.type) || COLUMN_TYPE_META[0];
    const typeColors = TYPE_COLORS[col.type] || TYPE_COLORS.text;

    return (
        <div style={R.wrapper}>
            <div style={R.row}>

                <div style={R.orderCol}>
                    <button onClick={onMoveUp}   disabled={index === 0}         style={R.arrow} title="Move up">▲</button>
                    <span style={R.indexNum}>{index + 1}</span>
                    <button onClick={onMoveDown} disabled={index === total - 1} style={R.arrow} title="Move down">▼</button>
                </div>

                <div style={{ ...R.typeBadge, background: typeColors.bg, color: typeColors.color, border: `1px solid ${typeColors.border}` }}>
                    {typeMeta.icon}
                </div>

                <input
                    value={col.header}
                    onChange={e => onUpdate({ header: e.target.value })}
                    style={R.headerInput}
                    placeholder="Column name"
                    maxLength={60}
                />

                <select
                    value={col.type}
                    onChange={e => onUpdate({ type: e.target.value, source: [], format: "" })}
                    style={{ ...R.typeSelect, color: typeColors.color, borderColor: typeColors.border, background: typeColors.bg }}
                >
                    {COLUMN_TYPE_META.map(t => (
                        <option key={t.value} value={t.value}>{t.label}</option>
                    ))}
                </select>

                {(typeMeta.hasSource || typeMeta.hasFormat) && (
                    <button onClick={onToggleExpand} style={R.expandBtn} title="More settings">
                        <svg width="10" height="10" viewBox="0 0 10 10" fill="none"
                            style={{ transform: isExpanded ? "rotate(180deg)" : "none", transition: "transform 0.2s" }}>
                            <path d="M1 3l4 4 4-4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                    </button>
                )}

                <button onClick={onDelete} style={R.deleteBtn} title={`Delete "${col.header}"`}>
                    <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                        <path d="M1 1l10 10M11 1L1 11" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round"/>
                    </svg>
                </button>
            </div>

            {isExpanded && (
                <div style={R.expanded}>
                    {typeMeta.hasSource && (
                        <div style={R.field}>
                            <label style={R.label}>Options <span style={R.labelHint}>(one per line)</span></label>
                            <textarea
                                value={(col.source || []).join("\n")}
                                onChange={e => onUpdate({ source: e.target.value.split("\n").map(s => s.trim()).filter(Boolean) })}
                                style={R.textarea}
                                placeholder={"Option A\nOption B\nOption C"}
                                rows={4}
                            />
                        </div>
                    )}
                    {typeMeta.hasFormat && (
                        <div style={R.field}>
                            <label style={R.label}>
                                {col.type === "date" ? "Date format" : "Number format"}
                                <span style={R.labelHint}> e.g. {col.type === "date" ? "DD/MM/YYYY" : "0,0.00"}</span>
                            </label>
                            <input
                                value={col.format || ""}
                                onChange={e => onUpdate({ format: e.target.value })}
                                style={R.formatInput}
                                placeholder={col.type === "date" ? "DD/MM/YYYY" : "0,0.00"}
                            />
                        </div>
                    )}
                    <label style={{ ...R.label, display: "flex", alignItems: "center", gap: 8, cursor: "pointer", marginTop: 6 }}>
                        <input
                            type="checkbox"
                            checked={col.readOnly || false}
                            onChange={e => onUpdate({ readOnly: e.target.checked })}
                            style={{ width: 14, height: 14, accentColor: "#2563eb" }}
                        />
                        Lock this column (read-only for users)
                    </label>
                </div>
            )}
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
        width: 440,
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
        padding: "16px 18px",
        borderBottom: "1px solid #f1f5f9",
        background: "linear-gradient(135deg, #f8faff 0%, #f0f4ff 100%)",
        flexShrink: 0,
    },
    headerLeft:  { display: "flex", alignItems: "center", gap: 12 },
    headerIcon:  {
        width: 36, height: 36, background: "#eff6ff",
        borderRadius: 8, border: "1px solid #bfdbfe",
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
    countBar: {
        display: "flex", alignItems: "center", gap: 10,
        padding: "10px 18px", borderBottom: "1px solid #f1f5f9",
        background: "#fafbfc", flexShrink: 0,
    },
    countPill: {
        display: "inline-flex", alignItems: "center",
        padding: "3px 10px", background: "#eff6ff",
        color: "#2563eb", border: "1px solid #bfdbfe",
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
        background: "linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%)",
        color: "#fff", border: "none", borderRadius: 8, cursor: "pointer",
        fontSize: 13, fontWeight: 600, letterSpacing: "0.01em",
        boxShadow: "0 2px 8px rgba(37,99,235,0.3)",
    },
};

const R = {
    wrapper:  { borderBottom: "1px solid #f8fafc" },
    row:      { display: "flex", alignItems: "center", gap: 8, padding: "10px 14px" },
    orderCol: { display: "flex", flexDirection: "column", alignItems: "center", gap: 1, flexShrink: 0 },
    arrow:    { background: "none", border: "none", cursor: "pointer", fontSize: 7, padding: "2px 3px", color: "#cbd5e1", lineHeight: 1 },
    indexNum: { fontSize: 9, fontWeight: 700, color: "#cbd5e1", lineHeight: 1 },
    typeBadge: {
        width: 28, height: 28, flexShrink: 0, borderRadius: 6,
        display: "flex", alignItems: "center", justifyContent: "center",
        fontSize: 12, fontWeight: 700,
    },
    headerInput: {
        flex: 1, border: "1px solid #e2e8f0", borderRadius: 6,
        padding: "6px 10px", fontSize: 13, color: "#0f172a",
        outline: "none", minWidth: 0, background: "#fff", fontFamily: "inherit",
    },
    typeSelect: {
        height: 32, padding: "0 8px", border: "1px solid",
        borderRadius: 6, fontSize: 11, fontWeight: 600,
        cursor: "pointer", outline: "none", flexShrink: 0,
    },
    expandBtn: {
        width: 28, height: 28, flexShrink: 0,
        background: "#f8fafc", border: "1px solid #e2e8f0",
        borderRadius: 6, cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center", color: "#64748b",
    },
    deleteBtn: {
        width: 28, height: 28, flexShrink: 0,
        background: "#fff5f5", border: "1px solid #fecaca",
        borderRadius: 6, cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center", color: "#ef4444",
    },
    expanded: {
        padding: "12px 14px 14px 50px",
        background: "#f8fafc", borderTop: "1px solid #f1f5f9",
    },
    field:     { marginBottom: 10 },
    label:     { display: "block", fontSize: 11, fontWeight: 600, color: "#475569", marginBottom: 5 },
    labelHint: { fontWeight: 400, color: "#94a3b8" },
    textarea: {
        width: "100%", border: "1px solid #e2e8f0", borderRadius: 6,
        padding: "7px 10px", fontSize: 12, fontFamily: "inherit",
        resize: "vertical", boxSizing: "border-box", outline: "none", background: "#fff",
    },
    formatInput: {
        width: "100%", border: "1px solid #e2e8f0", borderRadius: 6,
        padding: "7px 10px", fontSize: 12, outline: "none",
        boxSizing: "border-box", background: "#fff",
    },
};
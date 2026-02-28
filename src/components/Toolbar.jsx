/**
 * Toolbar.jsx
 *
 * Formatting toolbar rendered above the grid.
 * Commands the active HotTable instance directly via hotRef.
 *
 * NOTE: Column Settings button has been moved to the sheet name header
 * in WorkbookContainer.jsx — admin clicks it there instead.
 */

import { createElement, useState, useCallback, useEffect } from "react";
import { CSS } from "../utils/constants";
import { cellKey } from "../utils/helpers";

export function Toolbar({ hotRef, activeSheet, onMetaChange, disabled }) {

    const [activeFmt, setActiveFmt] = useState({
        bold:      false,
        italic:    false,
        underline: false,
        fontColor: "#000000",
        bgColor:   "#ffffff",
        align:     "left",
    });

    useEffect(() => {
        const hot = hotRef?.current?.hotInstance;
        if (!hot) return;

        const syncSelection = () => {
            const selected = hot.getSelected();
            if (!selected?.length) return;

            const [row, col] = selected[0];
            const key  = cellKey(row, col);
            const meta = activeSheet?.cellMeta?.[key] ?? {};

            setActiveFmt({
                bold:      meta.bold      ?? false,
                italic:    meta.italic    ?? false,
                underline: meta.underline ?? false,
                fontColor: meta.fontColor ?? "#000000",
                bgColor:   meta.bgColor   ?? "#ffffff",
                align:     meta.align     ?? "left",
            });
        };

        hot.addHook("afterSelectionEnd", syncSelection);
        return () => {
            try { hot.removeHook("afterSelectionEnd", syncSelection); } catch {}
        };
    }, [hotRef, activeSheet?.cellMeta]);

    const applyFormat = useCallback((property, value) => {
        const hot = hotRef?.current?.hotInstance;
        if (!hot || !activeSheet || disabled) return;

        const selected = hot.getSelected();
        if (!selected?.length) return;

        const newMeta = { ...(activeSheet.cellMeta ?? {}) };

        selected.forEach(([r1, c1, r2, c2]) => {
            const rowStart = Math.min(r1, r2);
            const rowEnd   = Math.max(r1, r2);
            const colStart = Math.min(c1, c2);
            const colEnd   = Math.max(c1, c2);

            for (let r = rowStart; r <= rowEnd; r++) {
                for (let c = colStart; c <= colEnd; c++) {
                    const key      = cellKey(r, c);
                    const existing = newMeta[key] ?? {};
                    newMeta[key]   = { ...existing, [property]: value };
                }
            }
        });

        onMetaChange?.(activeSheet.sheetId, newMeta);
        hot.render();
    }, [hotRef, activeSheet, onMetaChange, disabled]);

    const toggleFormat = useCallback((property) => {
        const newValue = !activeFmt[property];
        setActiveFmt(prev => ({ ...prev, [property]: newValue }));
        applyFormat(property, newValue);
    }, [activeFmt, applyFormat]);

    const setAlignment = useCallback((align) => {
        setActiveFmt(prev => ({ ...prev, align }));
        applyFormat("align", align);
    }, [applyFormat]);

    const mergeCells = useCallback(() => {
        const hot = hotRef?.current?.hotInstance;
        if (!hot || disabled) return;
        const plugin = hot.getPlugin("mergeCells");
        if (!plugin) return;
        const selected = hot.getSelected();
        if (!selected?.length) return;
        const [r1, c1, r2, c2] = selected[0];
        if (r1 === r2 && c1 === c2) return;
        plugin.merge(Math.min(r1,r2), Math.min(c1,c2), Math.max(r1,r2), Math.max(c1,c2));
        hot.render();
    }, [hotRef, disabled]);

    const unmergeCells = useCallback(() => {
        const hot = hotRef?.current?.hotInstance;
        if (!hot || disabled) return;
        const plugin = hot.getPlugin("mergeCells");
        if (!plugin) return;
        const selected = hot.getSelected();
        if (!selected?.length) return;
        const [r1, c1, r2, c2] = selected[0];
        plugin.unmerge(Math.min(r1,r2), Math.min(c1,c2), Math.max(r1,r2), Math.max(c1,c2));
        hot.render();
    }, [hotRef, disabled]);

    const clearFormatting = useCallback(() => {
        const hot = hotRef?.current?.hotInstance;
        if (!hot || !activeSheet || disabled) return;
        const selected = hot.getSelected();
        if (!selected?.length) return;
        const newMeta = { ...(activeSheet.cellMeta ?? {}) };
        selected.forEach(([r1, c1, r2, c2]) => {
            for (let r = Math.min(r1,r2); r <= Math.max(r1,r2); r++) {
                for (let c = Math.min(c1,c2); c <= Math.max(c1,c2); c++) {
                    delete newMeta[cellKey(r, c)];
                }
            }
        });
        setActiveFmt({ bold: false, italic: false, underline: false, fontColor: "#000000", bgColor: "#ffffff", align: "left" });
        onMetaChange?.(activeSheet.sheetId, newMeta);
        hot.render();
    }, [hotRef, activeSheet, onMetaChange, disabled]);

    return (
        <div className={CSS.TOOLBAR} role="toolbar" aria-label="Formatting toolbar">

            {/* ── Text Formatting ───────────────────────────────────── */}
            <ToolbarBtn title="Bold (Ctrl+B)" active={activeFmt.bold} disabled={disabled} onClick={() => toggleFormat("bold")}>
                <strong>B</strong>
            </ToolbarBtn>
            <ToolbarBtn title="Italic (Ctrl+I)" active={activeFmt.italic} disabled={disabled} onClick={() => toggleFormat("italic")}>
                <em>I</em>
            </ToolbarBtn>
            <ToolbarBtn title="Underline (Ctrl+U)" active={activeFmt.underline} disabled={disabled} onClick={() => toggleFormat("underline")}>
                <span style={{ textDecoration: "underline" }}>U</span>
            </ToolbarBtn>

            <Divider />

            {/* ── Colors ────────────────────────────────────────────── */}
            <ColorBtn
                title="Font Color" value={activeFmt.fontColor} disabled={disabled}
                label="A" labelStyle={{ color: activeFmt.fontColor, fontWeight: 700 }}
                onChange={(color) => { setActiveFmt(prev => ({ ...prev, fontColor: color })); applyFormat("fontColor", color); }}
            />
            <ColorBtn
                title="Background Color" value={activeFmt.bgColor} disabled={disabled}
                label={<span style={{ background: activeFmt.bgColor === "#ffffff" ? "#e0e0e0" : activeFmt.bgColor, padding: "2px 4px", borderRadius: 2 }}>⬛</span>}
                onChange={(color) => { setActiveFmt(prev => ({ ...prev, bgColor: color })); applyFormat("bgColor", color); }}
            />

            <Divider />

            {/* ── Alignment ─────────────────────────────────────────── */}
            <ToolbarBtn title="Align Left"   active={activeFmt.align === "left"}   disabled={disabled} onClick={() => setAlignment("left")}>≡</ToolbarBtn>
            <ToolbarBtn title="Align Center" active={activeFmt.align === "center"} disabled={disabled} onClick={() => setAlignment("center")}>☰</ToolbarBtn>
            <ToolbarBtn title="Align Right"  active={activeFmt.align === "right"}  disabled={disabled} onClick={() => setAlignment("right")}>
                <span style={{ display: "flex", flexDirection: "column", gap: 1, alignItems: "flex-end" }}>
                    <span style={{ width: 10, height: 1.5, background: "currentColor", display: "block" }} />
                    <span style={{ width: 14, height: 1.5, background: "currentColor", display: "block" }} />
                    <span style={{ width: 10, height: 1.5, background: "currentColor", display: "block" }} />
                </span>
            </ToolbarBtn>

            <Divider />

            {/* ── Merge ─────────────────────────────────────────────── */}
            <ToolbarBtn title="Merge selected cells"   disabled={disabled} onClick={mergeCells}>⊞</ToolbarBtn>
            <ToolbarBtn title="Unmerge selected cells" disabled={disabled} onClick={unmergeCells}>⊟</ToolbarBtn>

            <Divider />

            {/* ── Clear Formatting ──────────────────────────────────── */}
            <ToolbarBtn title="Clear formatting from selected cells" disabled={disabled} onClick={clearFormatting}>
                <span style={{ fontSize: 11, fontFamily: "sans-serif" }}>✕fmt</span>
            </ToolbarBtn>

        </div>
    );
}

// ─── Sub-components ───────────────────────────────────────────────────────────

function ToolbarBtn({ children, title, active, disabled, onClick }) {
    const className = [
        `${CSS.TOOLBAR}__btn`,
        active   ? `${CSS.TOOLBAR}__btn--active`  : "",
        disabled ? `${CSS.TOOLBAR}__btn--disabled` : "",
    ].filter(Boolean).join(" ");

    return (
        <button className={className} title={title} disabled={disabled}
            onClick={disabled ? undefined : onClick} type="button">
            {children}
        </button>
    );
}

function ColorBtn({ title, value, disabled, label, labelStyle, onChange }) {
    return (
        <label
            className={`${CSS.TOOLBAR}__btn`}
            title={title}
            style={{ cursor: disabled ? "default" : "pointer", opacity: disabled ? 0.4 : 1, position: "relative" }}
        >
            <span style={labelStyle}>{label}</span>
            <input
                type="color" value={value} disabled={disabled}
                onChange={e => onChange(e.target.value)}
                style={{ position: "absolute", width: 1, height: 1, opacity: 0, pointerEvents: disabled ? "none" : "auto" }}
                aria-label={title}
            />
        </label>
    );
}

function Divider() {
    return <div className={`${CSS.TOOLBAR}__divider`} aria-hidden="true" />;
}
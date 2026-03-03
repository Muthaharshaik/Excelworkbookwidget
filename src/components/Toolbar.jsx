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

            {/* ── Text style group ──────────────────────────────────── */}
            <ToolbarGroup>
                <ToolbarBtn title="Bold (Ctrl+B)" active={activeFmt.bold} disabled={disabled} onClick={() => toggleFormat("bold")}>
                    <IconBold />
                </ToolbarBtn>
                <ToolbarBtn title="Italic (Ctrl+I)" active={activeFmt.italic} disabled={disabled} onClick={() => toggleFormat("italic")}>
                    <IconItalic />
                </ToolbarBtn>
                <ToolbarBtn title="Underline (Ctrl+U)" active={activeFmt.underline} disabled={disabled} onClick={() => toggleFormat("underline")}>
                    <IconUnderline />
                </ToolbarBtn>
            </ToolbarGroup>

            <Divider />

            {/* ── Color group ───────────────────────────────────────── */}
            <ToolbarGroup>
                <FontColorBtn
                    title="Font Color"
                    value={activeFmt.fontColor}
                    disabled={disabled}
                    onChange={(color) => { setActiveFmt(prev => ({ ...prev, fontColor: color })); applyFormat("fontColor", color); }}
                />
                <BgColorBtn
                    title="Background Color"
                    value={activeFmt.bgColor}
                    disabled={disabled}
                    onChange={(color) => { setActiveFmt(prev => ({ ...prev, bgColor: color })); applyFormat("bgColor", color); }}
                />
            </ToolbarGroup>

            <Divider />

            {/* ── Alignment group ───────────────────────────────────── */}
            <ToolbarGroup>
                <ToolbarBtn title="Align Left"   active={activeFmt.align === "left"}   disabled={disabled} onClick={() => setAlignment("left")}>
                    <IconAlignLeft />
                </ToolbarBtn>
                <ToolbarBtn title="Align Center" active={activeFmt.align === "center"} disabled={disabled} onClick={() => setAlignment("center")}>
                    <IconAlignCenter />
                </ToolbarBtn>
                <ToolbarBtn title="Align Right"  active={activeFmt.align === "right"}  disabled={disabled} onClick={() => setAlignment("right")}>
                    <IconAlignRight />
                </ToolbarBtn>
            </ToolbarGroup>

            <Divider />

            {/* ── Merge group ───────────────────────────────────────── */}
            {/* <ToolbarGroup>
                <ToolbarBtn title="Merge selected cells"   disabled={disabled} onClick={mergeCells}>
                    <IconMerge />
                </ToolbarBtn>
                <ToolbarBtn title="Unmerge selected cells" disabled={disabled} onClick={unmergeCells}>
                    <IconUnmerge />
                </ToolbarBtn>
            </ToolbarGroup>

            <Divider /> */}

            {/* ── Clear formatting ──────────────────────────────────── */}
            <ToolbarGroup>
                <ToolbarBtn title="Clear formatting from selected cells" disabled={disabled} onClick={clearFormatting}>
                    <IconClearFormat />
                </ToolbarBtn>
            </ToolbarGroup>

        </div>
    );
}

// ─── Sub-components ───────────────────────────────────────────────────────────

function ToolbarGroup({ children }) {
    return (
        <div className="eww-toolbar__group">
            {children}
        </div>
    );
}

function ToolbarBtn({ children, title, active, disabled, onClick }) {
    const cls = [
        "eww-toolbar__btn",
        active   ? "eww-toolbar__btn--active"   : "",
        disabled ? "eww-toolbar__btn--disabled"  : "",
    ].filter(Boolean).join(" ");

    return (
        <button className={cls} title={title} disabled={disabled}
            onClick={disabled ? undefined : onClick} type="button">
            {children}
        </button>
    );
}

/** Font color button — "A" with a live color bar underneath */
function FontColorBtn({ title, value, disabled, onChange }) {
    return (
        <label className={["eww-toolbar__btn", "eww-toolbar__color-btn", disabled ? "eww-toolbar__btn--disabled" : ""].filter(Boolean).join(" ")}
            title={title} style={{ cursor: disabled ? "default" : "pointer" }}>
            <span className="eww-toolbar__color-btn__inner">
                <svg width="13" height="13" viewBox="0 0 13 13" fill="none" aria-hidden="true">
                    <text x="1" y="11" fontFamily="-apple-system, sans-serif" fontWeight="700" fontSize="12" fill="currentColor">A</text>
                </svg>
                <span className="eww-toolbar__color-swatch" style={{ background: value }} />
            </span>
            <input type="color" value={value} disabled={disabled}
                onChange={e => onChange(e.target.value)}
                style={{ position: "absolute", width: 1, height: 1, opacity: 0, pointerEvents: disabled ? "none" : "auto" }}
                aria-label={title} />
        </label>
    );
}

/** Background color button — paint bucket icon with a live color swatch */
function BgColorBtn({ title, value, disabled, onChange }) {
    const swatchColor = value === "#ffffff" ? "#e2e8f0" : value;
    return (
        <label className={["eww-toolbar__btn", "eww-toolbar__color-btn", disabled ? "eww-toolbar__btn--disabled" : ""].filter(Boolean).join(" ")}
            title={title} style={{ cursor: disabled ? "default" : "pointer" }}>
            <span className="eww-toolbar__color-btn__inner">
                <svg width="13" height="13" viewBox="0 0 14 14" fill="none" aria-hidden="true">
                    <path d="M3 10.5c0-.8.7-1.5 1.5-1.5s1.5.7 1.5 1.5S5.3 12 4.5 12 3 11.3 3 10.5z" fill="currentColor" opacity=".7"/>
                    <path d="M8 2L2 8l2.5 2.5 6-6L8 2z" stroke="currentColor" strokeWidth="1.2" strokeLinejoin="round" fill="none"/>
                    <path d="M10.5 4.5l1.5-1.5-1-1-1.5 1.5" stroke="currentColor" strokeWidth="1.1" strokeLinejoin="round" fill="none"/>
                </svg>
                <span className="eww-toolbar__color-swatch" style={{ background: swatchColor }} />
            </span>
            <input type="color" value={value} disabled={disabled}
                onChange={e => onChange(e.target.value)}
                style={{ position: "absolute", width: 1, height: 1, opacity: 0, pointerEvents: disabled ? "none" : "auto" }}
                aria-label={title} />
        </label>
    );
}

function Divider() {
    return <div className="eww-toolbar__divider" aria-hidden="true" />;
}

// ─── SVG Icons ────────────────────────────────────────────────────────────────

function IconBold() {
    return (
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <path d="M4 3h4a2.5 2.5 0 0 1 0 5H4V3z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" fill="none"/>
            <path d="M4 8h4.5a2.5 2.5 0 0 1 0 5H4V8z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" fill="none"/>
        </svg>
    );
}

function IconItalic() {
    return (
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <line x1="8.5" y1="2.5" x2="5.5" y2="11.5" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round"/>
            <line x1="6" y1="2.5" x2="11" y2="2.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="3" y1="11.5" x2="8" y2="11.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        </svg>
    );
}

function IconUnderline() {
    return (
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <path d="M3.5 2.5v5a3.5 3.5 0 0 0 7 0v-5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" fill="none"/>
            <line x1="2" y1="13" x2="12" y2="13" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        </svg>
    );
}

function IconAlignLeft() {
    return (
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <line x1="2" y1="3.5" x2="12" y2="3.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="2" y1="6.5" x2="9"  y2="6.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="2" y1="9.5" x2="12" y2="9.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="2" y1="12.5" x2="8" y2="12.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        </svg>
    );
}

function IconAlignCenter() {
    return (
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <line x1="2"   y1="3.5" x2="12" y2="3.5"  stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="3.5" y1="6.5" x2="10.5" y2="6.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="2"   y1="9.5" x2="12" y2="9.5"  stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="3.5" y1="12.5" x2="10.5" y2="12.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        </svg>
    );
}

function IconAlignRight() {
    return (
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <line x1="2"  y1="3.5" x2="12" y2="3.5"  stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="5"  y1="6.5" x2="12" y2="6.5"  stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="2"  y1="9.5" x2="12" y2="9.5"  stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="6"  y1="12.5" x2="12" y2="12.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        </svg>
    );
}

function IconMerge() {
    return (
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <rect x="1.5" y="1.5" width="11" height="11" rx="1" stroke="currentColor" strokeWidth="1.3" fill="none"/>
            <line x1="7" y1="1.5" x2="7" y2="12.5" stroke="currentColor" strokeWidth="1.1" strokeDasharray="2 1.5"/>
            <path d="M4.5 7H9.5M8 5.5L9.5 7 8 8.5" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
        </svg>
    );
}

function IconUnmerge() {
    return (
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <rect x="1.5" y="1.5" width="11" height="11" rx="1" stroke="currentColor" strokeWidth="1.3" fill="none"/>
            <line x1="7" y1="1.5" x2="7" y2="12.5" stroke="currentColor" strokeWidth="1.4"/>
            <path d="M5.5 5.5L4 7l1.5 1.5M8.5 5.5L10 7l-1.5 1.5" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
        </svg>
    );
}

function IconClearFormat() {
    return (
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <path d="M3 3.5h8M5.5 3.5V12M7 7.5H10.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="9" y1="9" x2="12.5" y2="12.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
            <line x1="12.5" y1="9" x2="9" y2="12.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        </svg>
    );
}
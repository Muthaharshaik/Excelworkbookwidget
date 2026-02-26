/**
 * Toolbar.jsx
 *
 * Formatting toolbar rendered above the grid.
 * Commands the active HotTable instance directly via hotRef.
 *
 * WHAT IT DOES:
 *   - Reads the currently selected cells from HotTable
 *   - Toggles / applies formatting (bold, italic, underline,
 *     font color, bg color, alignment, merge)
 *   - Stores formatting in cellMeta via onMetaChange callback
 *     (NOT in HotTable's own meta system — ours is serialized to JSON)
 *   - Refreshes HotTable so the custom renderer picks up the new meta
 *
 * WHAT IT DOES NOT DO:
 *   - Does not manage sheet state (WorkbookContainer owns that)
 *   - Does not save (useAutoSave owns that)
 *   - Does not appear when the active sheet is read-only
 *
 * FORMATTING ARCHITECTURE:
 *   We store formatting in sheet.cellMeta as:
 *     { "row,col": { bold, italic, underline, fontColor, bgColor, align } }
 *
 *   SheetGrid's custom renderer reads this on every cell paint.
 *   When Toolbar changes meta → onMetaChange → setSheets → SheetGrid
 *   re-renders → renderer picks up new styles. Clean one-way data flow.
 */

import { createElement, useState, useCallback, useEffect } from "react";
import { CSS } from "../utils/constants";
import { cellKey } from "../utils/helpers";

/**
 * @param {object}   props
 * @param {object}   props.hotRef          - ref to HotTable instance
 * @param {object}   props.activeSheet     - current sheet object (for reading cellMeta)
 * @param {Function} props.onMetaChange    - (sheetId, newMeta) => void
 * @param {boolean}  props.disabled        - true when sheet is read-only
 */
export function Toolbar({ hotRef, activeSheet, onMetaChange, disabled, isAdmin, onOpenColumnSettings }) {

    // ── Track active formatting state of selected cells ───────────────────
    // When user clicks a cell, we read its meta and update these states
    // so toolbar buttons show as active/inactive correctly.
    const [activeFmt, setActiveFmt] = useState({
        bold:      false,
        italic:    false,
        underline: false,
        fontColor: "#000000",
        bgColor:   "#ffffff",
        align:     "left",
    });

    // ── Sync toolbar state when selection changes ──────────────────────────
    // HotTable fires afterSelectionEnd when user clicks/selects cells.
    // We read the top-left cell's meta to reflect its formatting in toolbar.
    useEffect(() => {
        const hot = hotRef?.current?.hotInstance;
        if (!hot) return;

        const syncSelection = () => {
            const selected = hot.getSelected(); // [[r1,c1,r2,c2], ...]
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
            // Safe remove — HT may already be destroyed
            try { hot.removeHook("afterSelectionEnd", syncSelection); } catch {}
        };
    }, [hotRef, activeSheet?.cellMeta]);

    // ── Core: apply a formatting property to all selected cells ───────────
    const applyFormat = useCallback((property, value) => {
        const hot = hotRef?.current?.hotInstance;
        if (!hot || !activeSheet || disabled) return;

        const selected = hot.getSelected(); // [[r1,c1,r2,c2], ...]
        if (!selected?.length) return;

        // Deep-clone existing meta so we don't mutate state directly
        const newMeta = { ...(activeSheet.cellMeta ?? {}) };

        // Expand all selection ranges and apply to each cell
        selected.forEach(([r1, c1, r2, c2]) => {
            const rowStart = Math.min(r1, r2);
            const rowEnd   = Math.max(r1, r2);
            const colStart = Math.min(c1, c2);
            const colEnd   = Math.max(c1, c2);

            for (let r = rowStart; r <= rowEnd; r++) {
                for (let c = colStart; c <= colEnd; c++) {
                    const key        = cellKey(r, c);
                    const existing   = newMeta[key] ?? {};
                    newMeta[key]     = { ...existing, [property]: value };
                }
            }
        });

        // Fire up to WorkbookContainer → setSheets → SheetGrid re-renders
        onMetaChange?.(activeSheet.sheetId, newMeta);

        // Force HT to repaint cells so the custom renderer picks up new meta
        hot.render();

    }, [hotRef, activeSheet, onMetaChange, disabled]);

    // ── Toggle helpers (bold, italic, underline) ───────────────────────────
    const toggleFormat = useCallback((property) => {
        const currentValue = activeFmt[property];
        const newValue     = !currentValue;
        setActiveFmt(prev => ({ ...prev, [property]: newValue }));
        applyFormat(property, newValue);
    }, [activeFmt, applyFormat]);

    // ── Alignment ─────────────────────────────────────────────────────────
    const setAlignment = useCallback((align) => {
        setActiveFmt(prev => ({ ...prev, align }));
        applyFormat("align", align);
    }, [applyFormat]);

    // ── Merge cells ───────────────────────────────────────────────────────
    const mergeCells = useCallback(() => {
        const hot = hotRef?.current?.hotInstance;
        if (!hot || disabled) return;
        const plugin = hot.getPlugin("mergeCells");
        if (!plugin) return;

        const selected = hot.getSelected();
        if (!selected?.length) return;

        const [r1, c1, r2, c2] = selected[0];
        if (r1 === r2 && c1 === c2) return; // single cell — nothing to merge

        plugin.merge(
            Math.min(r1, r2), Math.min(c1, c2),
            Math.max(r1, r2), Math.max(c1, c2)
        );
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
        plugin.unmerge(
            Math.min(r1, r2), Math.min(c1, c2),
            Math.max(r1, r2), Math.max(c1, c2)
        );
        hot.render();
    }, [hotRef, disabled]);

    // ── Clear formatting ───────────────────────────────────────────────────
    const clearFormatting = useCallback(() => {
        const hot = hotRef?.current?.hotInstance;
        if (!hot || !activeSheet || disabled) return;

        const selected = hot.getSelected();
        if (!selected?.length) return;

        const newMeta = { ...(activeSheet.cellMeta ?? {}) };

        selected.forEach(([r1, c1, r2, c2]) => {
            for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
                for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
                    delete newMeta[cellKey(r, c)];
                }
            }
        });

        setActiveFmt({ bold: false, italic: false, underline: false, fontColor: "#000000", bgColor: "#ffffff", align: "left" });
        onMetaChange?.(activeSheet.sheetId, newMeta);
        hot.render();
    }, [hotRef, activeSheet, onMetaChange, disabled]);

    // ── Render ─────────────────────────────────────────────────────────────
    return (
        <div className={CSS.TOOLBAR} role="toolbar" aria-label="Formatting toolbar">

            {/* ── Text Formatting Group ─────────────────────────────── */}
            <ToolbarBtn
                title="Bold (Ctrl+B)"
                active={activeFmt.bold}
                disabled={disabled}
                onClick={() => toggleFormat("bold")}
            >
                <strong>B</strong>
            </ToolbarBtn>

            <ToolbarBtn
                title="Italic (Ctrl+I)"
                active={activeFmt.italic}
                disabled={disabled}
                onClick={() => toggleFormat("italic")}
            >
                <em>I</em>
            </ToolbarBtn>

            <ToolbarBtn
                title="Underline (Ctrl+U)"
                active={activeFmt.underline}
                disabled={disabled}
                onClick={() => toggleFormat("underline")}
            >
                <span style={{ textDecoration: "underline" }}>U</span>
            </ToolbarBtn>

            <Divider />

            {/* ── Color Group ───────────────────────────────────────── */}
            <ColorBtn
                title="Font Color"
                value={activeFmt.fontColor}
                disabled={disabled}
                label="A"
                labelStyle={{ color: activeFmt.fontColor, fontWeight: 700 }}
                onChange={(color) => {
                    setActiveFmt(prev => ({ ...prev, fontColor: color }));
                    applyFormat("fontColor", color);
                }}
            />

            <ColorBtn
                title="Background Color"
                value={activeFmt.bgColor}
                disabled={disabled}
                label={<span style={{ background: activeFmt.bgColor === "#ffffff" ? "#e0e0e0" : activeFmt.bgColor, padding: "2px 4px", borderRadius: 2 }}>⬛</span>}
                onChange={(color) => {
                    setActiveFmt(prev => ({ ...prev, bgColor: color }));
                    applyFormat("bgColor", color);
                }}
            />

            <Divider />

            {/* ── Alignment Group ───────────────────────────────────── */}
            <ToolbarBtn
                title="Align Left"
                active={activeFmt.align === "left"}
                disabled={disabled}
                onClick={() => setAlignment("left")}
            >
                ≡
            </ToolbarBtn>

            <ToolbarBtn
                title="Align Center"
                active={activeFmt.align === "center"}
                disabled={disabled}
                onClick={() => setAlignment("center")}
            >
                ☰
            </ToolbarBtn>

            <ToolbarBtn
                title="Align Right"
                active={activeFmt.align === "right"}
                disabled={disabled}
                onClick={() => setAlignment("right")}
            >
                <span style={{ display: "flex", flexDirection: "column", gap: 1, alignItems: "flex-end" }}>
                    <span style={{ width: 10, height: 1.5, background: "currentColor", display: "block" }} />
                    <span style={{ width: 14, height: 1.5, background: "currentColor", display: "block" }} />
                    <span style={{ width: 10, height: 1.5, background: "currentColor", display: "block" }} />
                </span>
            </ToolbarBtn>

            <Divider />

            {/* ── Merge Group ───────────────────────────────────────── */}
            <ToolbarBtn
                title="Merge selected cells"
                disabled={disabled}
                onClick={mergeCells}
            >
                ⊞
            </ToolbarBtn>

            <ToolbarBtn
                title="Unmerge selected cells"
                disabled={disabled}
                onClick={unmergeCells}
            >
                ⊟
            </ToolbarBtn>

            <Divider />

            {/* ── Clear ─────────────────────────────────────────────── */}
            <ToolbarBtn
                title="Clear formatting from selected cells"
                disabled={disabled}
                onClick={clearFormatting}
            >
                <span style={{ fontSize: 11, fontFamily: "sans-serif" }}>✕fmt</span>
            </ToolbarBtn>

            {/* ── Admin: Column Settings ────────────────────────────── */}
            {isAdmin && (
                <span>
                    <Divider />
                    <ToolbarBtn
                        title="Edit column headers and data types (Admin only)"
                        disabled={false}
                        onClick={onOpenColumnSettings}
                    >
                        <span style={{ fontSize: 11 }}>⚙ Columns</span>
                    </ToolbarBtn>
                </span>
            )}

        </div>
    );
}

// ─── Small reusable sub-components ───────────────────────────────────────────

function ToolbarBtn({ children, title, active, disabled, onClick }) {
    const className = [
        `${CSS.TOOLBAR}__btn`,
        active   ? `${CSS.TOOLBAR}__btn--active`   : "",
        disabled ? `${CSS.TOOLBAR}__btn--disabled`  : "",
    ].filter(Boolean).join(" ");

    return (
        <button
            className={className}
            title={title}
            disabled={disabled}
            onClick={disabled ? undefined : onClick}
            type="button"
        >
            {children}
        </button>
    );
}

/**
 * ColorBtn
 * A toolbar button that opens a native color picker via a hidden <input type="color">.
 */
function ColorBtn({ title, value, disabled, label, labelStyle, onChange }) {
    return (
        <label
            className={`${CSS.TOOLBAR}__btn`}
            title={title}
            style={{ cursor: disabled ? "default" : "pointer", opacity: disabled ? 0.4 : 1, position: "relative" }}
        >
            <span style={labelStyle}>{label}</span>
            <input
                type="color"
                value={value}
                disabled={disabled}
                onChange={e => onChange(e.target.value)}
                style={{
                    position: "absolute",
                    width:    1,
                    height:   1,
                    opacity:  0,
                    pointerEvents: disabled ? "none" : "auto",
                }}
                aria-label={title}
            />
        </label>
    );
}

function Divider() {
    return <div className={`${CSS.TOOLBAR}__divider`} aria-hidden="true" />;
}
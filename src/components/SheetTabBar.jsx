/**
 * SheetTabBar.jsx
 *
 * Sheet tab strip at the bottom of the workbook.
 *
 * FEATURES:
 *   - Click tab          → switch active sheet
 *   - Double-click tab   → rename sheet (inline edit)
 *   - Click ✕ on tab     → delete sheet
 *   - Click + button     → add new empty sheet
 *   - Lock icon          → sheet is view-only for this user
 *
 * PERMISSIONS:
 *   isWorkbookEditable = true  → shows + button and ✕ on tabs
 *   isWorkbookEditable = false → view-only user, no add/delete/rename
 *
 * All mutations update local state in WorkbookContainer which triggers
 * auto-save — no Mendix microflow needed for sheet management.
 */

import { createElement, useState, useRef, useEffect, useCallback } from "react";
import { CSS }      from "../utils/constants";
import { truncate } from "../utils/helpers";

export function SheetTabBar({
    sheets,
    activeIndex,
    isWorkbookEditable,
    canEditSheet,
    onTabChange,
    onAddSheet,
    onDeleteSheet,
    onRenameSheet,
}) {
    const [renamingId, setRenamingId]   = useState(null);
    const [renameValue, setRenameValue] = useState("");
    const renameInputRef                = useRef(null);

    // Focus rename input when it appears
    useEffect(() => {
        if (renamingId && renameInputRef.current) {
            renameInputRef.current.focus();
            renameInputRef.current.select();
        }
    }, [renamingId]);

    const startRename = useCallback((e, sheet) => {
        e.stopPropagation();
        if (!isWorkbookEditable) return;
        setRenamingId(sheet.sheetId);
        setRenameValue(sheet.sheetName);
    }, [isWorkbookEditable]);

    const commitRename = useCallback(() => {
        if (!renamingId) return;
        const trimmed = renameValue.trim();
        if (trimmed) onRenameSheet?.(renamingId, trimmed);
        setRenamingId(null);
        setRenameValue("");
    }, [renamingId, renameValue, onRenameSheet]);

    const handleRenameKeyDown = useCallback((e) => {
        if (e.key === "Enter")  { e.preventDefault(); commitRename(); }
        if (e.key === "Escape") { setRenamingId(null); setRenameValue(""); }
    }, [commitRename]);

    const handleDelete = useCallback((e, sheet) => {
        e.stopPropagation();
        if (!isWorkbookEditable) return;
        if (sheets.length <= 1) { alert("A workbook must have at least one sheet."); return; }
        onDeleteSheet?.(sheet.sheetId);
    }, [isWorkbookEditable, sheets.length, onDeleteSheet]);

    if (!sheets || sheets.length === 0) return null;

    return (
        <div className={CSS.TAB_BAR} role="tablist" aria-label="Sheets">

            {sheets.map((sheet, index) => {
                const isActive   = index === activeIndex;
                const isEditable = canEditSheet(sheet.isEditable);
                const isRenaming = renamingId === sheet.sheetId;

                const tabClassName = [
                    CSS.TAB,
                    isActive    ? CSS.TAB_ACTIVE   : "",
                    !isEditable ? CSS.TAB_READONLY : "",
                ].filter(Boolean).join(" ");

                return (
                    <div key={sheet.sheetId} style={{ position: "relative", display: "flex", flexShrink: 0 }}>
                        <button
                            role="tab"
                            aria-selected={isActive}
                            className={tabClassName}
                            onClick={() => !isRenaming && onTabChange(index)}
                            onDoubleClick={(e) => startRename(e, sheet)}
                            title={isWorkbookEditable ? "Click to switch • Double-click to rename" : sheet.sheetName}
                            style={{ paddingRight: isWorkbookEditable ? 20 : 12 }}
                        >
                            {!isEditable && <span style={{ fontSize: 10, marginRight: 3 }}>🔒</span>}

                            {isRenaming ? (
                                <input
                                    ref={renameInputRef}
                                    value={renameValue}
                                    onChange={e => setRenameValue(e.target.value)}
                                    onBlur={commitRename}
                                    onKeyDown={handleRenameKeyDown}
                                    onClick={e => e.stopPropagation()}
                                    maxLength={50}
                                    style={renameInputStyle}
                                />
                            ) : (
                                <span className="eww-tab__name">{truncate(sheet.sheetName, 20)}</span>
                            )}
                        </button>

                        {/* ✕ delete button — only for editors */}
                        {isWorkbookEditable && !isRenaming && (
                            <button
                                onClick={(e) => handleDelete(e, sheet)}
                                title={`Delete "${sheet.sheetName}"`}
                                className="eww-tab__delete-btn"
                                style={{ ...deleteButtonStyle, opacity: isActive ? 1 : 0 }}
                            >
                                ✕
                            </button>
                        )}
                    </div>
                );
            })}

            {/* + add sheet button */}
            {isWorkbookEditable && (
                <button
                    onClick={onAddSheet}
                    title="Add new sheet"
                    className="eww-tab eww-tab--add"
                    style={addButtonStyle}
                >
                    +
                </button>
            )}

        </div>
    );
}

const renameInputStyle = {
    background: "transparent", border: "none",
    borderBottom: "1px solid #1a73e8", outline: "none",
    fontSize: 12, fontFamily: "inherit", color: "inherit",
    width: 80, padding: 0, margin: 0,
};

const deleteButtonStyle = {
    position: "absolute", right: 2, top: "50%", transform: "translateY(-50%)",
    background: "none", border: "none", cursor: "pointer",
    fontSize: 9, color: "#5f6368", padding: "0 2px", lineHeight: 1,
    borderRadius: 2, transition: "opacity 0.15s",
};

const addButtonStyle = {
    fontSize: 18, fontWeight: 300, padding: "0 10px",
    color: "#5f6368", flexShrink: 0,
};
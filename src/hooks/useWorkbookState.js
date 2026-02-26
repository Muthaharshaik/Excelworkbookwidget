/**
 * useWorkbookState.js
 *
 * Owns the core workbook state:
 *   - sheets[]         : parsed array of sheet objects
 *   - activeSheetIndex : which tab is currently visible
 *
 * RESPONSIBILITIES:
 *   1. Parse sheetsJson from Mendix into sheets[] on first load
 *   2. Re-parse when Mendix refreshes sheetsJson (after a DB commit)
 *      BUT only if the update came from Mendix (not from the user typing) —
 *      we guard against overwriting unsaved local edits
 *   3. Expose setSheets so other hooks (useAutoSave) can update cell data
 *   4. Expose activeSheetIndex + setActiveSheetIndex for tab switching
 *
 * WHAT THIS HOOK DOES NOT DO:
 *   - Does not call any Mendix actions (that's mendixBridge.js)
 *   - Does not debounce or save (that's useAutoSave.js)
 *   - Does not know about permissions (that's usePermissions.js)
 */

import { useState, useEffect, useRef, useCallback } from "react";
import { parseSheets } from "../services/dataService";
import { clampIndex } from "../utils/helpers";

/**
 * @param {string}  sheetsJsonProp  - raw value from Mendix sheetsJson attribute
 *                                    (changes whenever Mendix refreshes the entity)
 */
export function useWorkbookState(sheetsJsonProp) {

    // ── Core state ─────────────────────────────────────────────────────────
    const [sheets, setSheets]                   = useState([]);
    const [activeSheetIndex, setActiveSheetIndex] = useState(0);

    // ── Loading / error state ───────────────────────────────────────────────
    const [isLoading, setIsLoading]   = useState(true);
    const [parseError, setParseError] = useState(null);

    // ── Guard: track the last JSON we loaded from Mendix ──────────────────
    // We use a ref (not state) so comparing it doesn't trigger re-renders.
    // When Mendix sends a new sheetsJson, we compare against this.
    // If they're the same string, we skip re-parsing (user is mid-edit).
    const lastLoadedJsonRef = useRef(null);

    // ── Track whether the user has made unsaved local edits ───────────────
    // If true, we do NOT overwrite sheets[] when Mendix refreshes the prop.
    // This is cleared by useAutoSave after a successful save + DB commit.
    const hasPendingEditsRef = useRef(false);

    // ── Initial load + Mendix-driven refresh ──────────────────────────────
    useEffect(() => {
        // No prop yet — Mendix datasource still loading
        if (sheetsJsonProp === undefined || sheetsJsonProp === null) {
            setIsLoading(true);
            return;
        }

        // Prop hasn't actually changed since last load — skip
        if (sheetsJsonProp === lastLoadedJsonRef.current) {
            setIsLoading(false);
            return;
        }

        // User has unsaved edits — DO NOT overwrite their work with
        // a Mendix refresh. Wait until after the save cycle clears the flag.
        if (hasPendingEditsRef.current) {
            console.info(
                "[ExcelWidget] Mendix sheetsJson changed but user has pending edits. " +
                "Skipping reload to protect unsaved data."
            );
            setIsLoading(false);
            return;
        }

        // ── Parse the new JSON ────────────────────────────────────────────
        setIsLoading(true);
        setParseError(null);

        try {
            const parsed = parseSheets(sheetsJsonProp);
            setSheets(parsed);

            // Keep activeSheetIndex in bounds if sheet count changed
            setActiveSheetIndex(prev => clampIndex(prev, 0, Math.max(0, parsed.length - 1)));

            // Remember what we just loaded so we can detect future changes
            lastLoadedJsonRef.current = sheetsJsonProp;

        } catch (err) {
            // parseSheets handles its own errors internally, but belt-and-suspenders
            console.error("[ExcelWidget] useWorkbookState parse error:", err.message);
            setParseError(err.message);
            setSheets([]);
        } finally {
            setIsLoading(false);
        }

    }, [sheetsJsonProp]);

    // ── Safe active sheet index setter ────────────────────────────────────
    // Clamps the index to valid range so callers don't have to.
    const safeSetActiveSheetIndex = useCallback((index) => {
        setActiveSheetIndex(prev => {
            const clamped = clampIndex(index, 0, Math.max(0, sheets.length - 1));
            return clamped;
        });
    }, [sheets.length]);

    // ── Mark pending edits ────────────────────────────────────────────────
    // Called by useAutoSave whenever the user makes a cell change.
    // Prevents Mendix prop refresh from overwriting local state.
    const markPendingEdits = useCallback(() => {
        hasPendingEditsRef.current = true;
    }, []);

    // ── Clear pending edits ───────────────────────────────────────────────
    // Called by useAutoSave AFTER a successful save + DB commit cycle.
    // After this, the next Mendix prop refresh WILL reload sheets.
    const clearPendingEdits = useCallback(() => {
        hasPendingEditsRef.current = false;
        // Also update lastLoadedJsonRef so the next Mendix refresh
        // (which will contain the just-saved data) is recognised as new
        lastLoadedJsonRef.current = null;
    }, []);

    // ── Active sheet object (convenience) ─────────────────────────────────
    const activeSheet = sheets[activeSheetIndex] ?? null;

    return {
        // State
        sheets,
        setSheets,
        activeSheet,
        activeSheetIndex,
        setActiveSheetIndex: safeSetActiveSheetIndex,

        // Loading / error
        isLoading,
        parseError,

        // Pending edit guards (used by useAutoSave)
        markPendingEdits,
        clearPendingEdits,
    };
}
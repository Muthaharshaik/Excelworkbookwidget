/**
 * useAutoSave.js
 *
 * Watches sheets[] state for changes and auto-saves to Mendix.
 *
 * FLOW:
 *   sheets[] state changes (user edited a cell)
 *       ↓
 *   savingStatus → "saving"
 *       ↓
 *   wait AUTOSAVE_DEBOUNCE_MS (800ms) — resets on every new change
 *       ↓
 *   serializeSheets(sheets) → JSON string
 *       ↓
 *   sheetsJson.setValue(newJson)   ← writes back into Mendix attribute
 *       ↓
 *   onSheetChange.execute()        ← fires Mendix microflow to commit
 *       ↓
 *   savingStatus → "saved" (for 2s) → "idle"
 *       ↓
 *   clearPendingEdits()            ← allows next Mendix refresh to reload
 */

import { useState, useEffect, useRef, useCallback } from "react";
import { serializeSheets }      from "../services/dataService";
import { triggerSheetChange }   from "../services/mendixBridge";
import { AUTOSAVE_DEBOUNCE_MS } from "../utils/constants";

/**
 * @param {object} params
 * @param {SheetObject[]}  params.sheets          - current sheets state (changes on every edit)
 * @param {object}         params.onSheetChange   - Mendix action prop
 * @param {object}         params.sheetsJson      - Mendix attribute prop (EditableValue with setValue)
 * @param {Function}       params.clearPendingEdits - from useWorkbookState
 */
export function useAutoSave({ sheets, onSheetChange, sheetsJson, clearPendingEdits }) {

    // "idle" | "saving" | "saved"
    const [savingStatus, setSavingStatus] = useState("idle");

    // Timer refs
    const debounceTimer  = useRef(null);
    const savedTimer     = useRef(null);

    // Track if we've done at least one load (skip saving on initial parse)
    const isInitialLoad  = useRef(true);

    // ── Mark first load done ───────────────────────────────────────────────
    // After the first time sheets arrives (parsed from Mendix), we mark
    // initial load as done. Subsequent changes are user edits → save them.
    useEffect(() => {
        if (sheets.length > 0 && isInitialLoad.current) {
            isInitialLoad.current = false;
        }
    }, [sheets]);

    // ── Auto-save effect ───────────────────────────────────────────────────
    useEffect(() => {
        // Skip: nothing loaded yet
        if (isInitialLoad.current) return;
        // Skip: no sheets
        if (!sheets.length) return;

        setSavingStatus("saving");

        // Clear any existing debounce
        clearTimeout(debounceTimer.current);

        debounceTimer.current = setTimeout(() => {
            performSave();
        }, AUTOSAVE_DEBOUNCE_MS);

        // Cleanup on unmount or before next effect run
        return () => clearTimeout(debounceTimer.current);

    }, [sheets]); // Re-runs whenever sheets state changes

    // ── Perform the actual save ────────────────────────────────────────────
    const performSave = useCallback(() => {
        try {
            // 1. Serialize current sheets state to JSON string
            const newJson = serializeSheets(sheets);

            // 2. Write into Mendix attribute + fire commit microflow
            //    Both steps handled by mendixBridge.triggerSheetChange
            const success = triggerSheetChange(sheetsJson, newJson, onSheetChange);

            if (!success) {
                // Bridge already logged the specific reason — just reset status
                setSavingStatus("idle");
                return;
            }

            // 3. Update status → "saved" for 2s then back to "idle"
            setSavingStatus("saved");
            clearPendingEdits();

            clearTimeout(savedTimer.current);
            savedTimer.current = setTimeout(() => {
                setSavingStatus("idle");
            }, 2000);

        } catch (err) {
            console.error("[ExcelWidget] Auto-save failed:", err.message);
            setSavingStatus("idle");
        }
    }, [sheets, sheetsJson, onSheetChange, clearPendingEdits]);

    // ── Cleanup on unmount ─────────────────────────────────────────────────
    useEffect(() => {
        return () => {
            clearTimeout(debounceTimer.current);
            clearTimeout(savedTimer.current);
        };
    }, []);

    return { savingStatus };
}
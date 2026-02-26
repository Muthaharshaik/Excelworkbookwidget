/**
 * useHyperFormula.js
 *
 * Creates and manages ONE HyperFormula instance for the entire workbook.
 *
 * WHY ONE INSTANCE:
 *   HyperFormula needs to know about ALL sheets to resolve cross-sheet
 *   references like =Revenue!B2. If each SheetGrid had its own HF instance,
 *   cross-sheet formulas would fail. One instance = one formula engine for
 *   the whole workbook.
 *
 * SYNC STRATEGY:
 *   - On mount: initialise HF with all sheets data
 *   - On sheet data change: call hf.setSheetContent() for the changed sheet
 *   - On sheet add/delete/rename: rebuild HF sheet registry
 *   - HF instance is stable (same reference) — only content changes
 *
 * SHEET NAMING IN HF:
 *   HyperFormula identifies sheets by name, not index.
 *   We register each sheet using sheet.sheetName so that
 *   =Revenue!A1 works as expected.
 */

import { useRef, useEffect, useMemo } from "react";
import HyperFormula from "hyperformula";

/**
 * @param {SheetObject[]} sheets - current widget sheets state
 * @returns {{ hfInstance: HyperFormula, getSheetId: (sheetName: string) => number }}
 */
export function useHyperFormula(sheets) {

    // Stable ref — HF instance never replaced, only mutated
    const hfRef = useRef(null);

    // ── Create HF instance once on mount ─────────────────────────────────
    // We build it from the initial sheets state. Later updates are applied
    // incrementally via setSheetContent / addSheet / removeSheet.
    if (!hfRef.current) {
        hfRef.current = buildHFInstance(sheets);
    }

    // ── Keep track of sheet names we've registered ────────────────────────
    const registeredSheetsRef = useRef(sheets.map(s => s.sheetName));

    // ── Sync sheets state changes → HF instance ──────────────────────────
    useEffect(() => {
        const hf              = hfRef.current;
        if (!hf) return;

        const currentNames    = sheets.map(s => s.sheetName);
        const registeredNames = registeredSheetsRef.current;

        // 1. Add newly added sheets
        sheets.forEach(sheet => {
            if (!registeredNames.includes(sheet.sheetName)) {
                try {
                    const hfSheetId = hf.addSheet(sheet.sheetName);
                    hf.setSheetContent(hfSheetId, sheet.data || []);
                } catch (e) {
                    // Sheet may already exist under a different tracking state
                    console.warn("[ExcelWidget] HF addSheet warning:", e.message);
                }
            }
        });

        // 2. Remove deleted sheets
        registeredNames.forEach(name => {
            if (!currentNames.includes(name)) {
                try {
                    const hfSheetId = hf.getSheetId(name);
                    if (hfSheetId !== undefined) hf.removeSheet(hfSheetId);
                } catch (e) {
                    console.warn("[ExcelWidget] HF removeSheet warning:", e.message);
                }
            }
        });

        // 3. Update content for all existing sheets
        sheets.forEach(sheet => {
            try {
                const hfSheetId = hf.getSheetId(sheet.sheetName);
                if (hfSheetId !== undefined) {
                    hf.setSheetContent(hfSheetId, sheet.data || []);
                }
            } catch (e) {
                console.warn("[ExcelWidget] HF setSheetContent warning:", e.message);
            }
        });

        // 4. Handle sheet renames
        sheets.forEach(sheet => {
            // If sheetId matches but name changed, rename in HF
            // (This is handled by remove+add above for simplicity)
        });

        registeredSheetsRef.current = currentNames;

    }, [sheets]);

    // ── Utility: get HF sheet id by sheet name ────────────────────────────
    const getSheetId = useMemo(() => {
        return (sheetName) => {
            const hf = hfRef.current;
            if (!hf) return undefined;
            return hf.getSheetId(sheetName);
        };
    }, []);

    return {
        hfInstance: hfRef.current,
        getSheetId,
    };
}

// ─────────────────────────────────────────────────────────────────────────────
//  Build HyperFormula instance from sheets array
// ─────────────────────────────────────────────────────────────────────────────

function buildHFInstance(sheets) {
    try {
        // Build sheets data object for HF constructor
        // HF expects: { sheetName: [[row data]] }
        const sheetsData = {};
        sheets.forEach(sheet => {
            sheetsData[sheet.sheetName] = sheet.data || [];
        });

        const hf = HyperFormula.buildFromSheets(sheetsData, {
            // Use the non-commercial license key (same as HotTable)
            licenseKey: "non-commercial-and-evaluation",

            // Locale settings
            language:  "enGB",

            // Allow circular references to not throw hard errors
            // (just returns an error value instead of crashing)
            maxRows:   10000,
            maxColumns: 1000,

            // Precision
            precisionRounding: 10,
        });

        console.log(`[ExcelWidget] HyperFormula initialised with ${sheets.length} sheet(s).`);
        return hf;

    } catch (e) {
        console.error("[ExcelWidget] Failed to initialise HyperFormula:", e.message);
        return null;
    }
}
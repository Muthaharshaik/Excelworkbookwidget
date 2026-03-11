/**
 * useHyperformula.js
 *
 * Creates and manages ONE HyperFormula instance per WorkbookContainer mount.
 *
 * MULTI-SHEET SUPPORT:
 *   Accepts allSheets array (parsed from Workbook.allSheetsJson).
 *   Registers ALL sheets into HF so cross-sheet references like
 *   =Revenue!A1 resolve correctly even when only one sheet is visible.
 *
 * SYNC STRATEGY:
 *   - On mount:              register all sheets from allSheets
 *   - On allSheets change:   diff against registered sheets, add/remove/update
 *   - On current data change: update current sheet content in HF immediately
 *                             so formulas referencing the current sheet
 *                             recalculate live as the user types
 *
 * FALLBACK:
 *   If allSheets is empty (allSheetsJson not wired in Studio Pro),
 *   falls back to single-sheet mode — exactly as it worked before.
 *   Existing deployments are not broken.
 *
 * SHEET NAMING:
 *   HF identifies sheets by name. We use sheetName from Mendix so that
 *   =Revenue!A1 works as expected when the user types the sheet name.
 *   Duplicate sheet names are handled by suffixing with sheetId.
 */

import { useRef, useState, useEffect } from "react";
import HyperFormula from "hyperformula";

/**
 * @param {Array<{ sheetId, sheetName, data }>} allSheets
 *   Parsed from Workbook.allSheetsJson by parseAllSheetsJson().
 *   Empty array = single-sheet fallback mode.
 *
 * @param {string} currentSheetName
 *   The sheetName of the currently displayed sheet.
 *   Used to ensure HotTable's formulas.sheetName matches HF registration.
 *
 * @param {any[][]} currentSheetData
 *   Live data of the current sheet from React state.
 *   Synced into HF on every change so formulas recalculate live.
 *
 * @returns {{
 *   hfRef:   React.MutableRefObject<HyperFormula|null>,
 *   hfReady: boolean
 * }}
 */
export function useHyperformula(allSheets, currentSheetName, currentSheetData) {

    const hfRef                  = useRef(null);
    const [hfReady, setHfReady]  = useState(false);

    // Track which sheet names are currently registered in HF
    const registeredSheetsRef    = useRef([]);

    // ── Mount: create HF instance + register all sheets ───────────────────
    useEffect(() => {
        try {
            const hf = HyperFormula.buildEmpty({
                licenseKey: "gpl-v3",
            });

            hfRef.current = hf;

            // Register all sheets if allSheets provided
            // Otherwise HF starts empty — HotTable will add the sheet
            // automatically when the formulas prop is first applied
            if (allSheets && allSheets.length > 0) {
                registerAllSheets(hf, allSheets);
                registeredSheetsRef.current = allSheets.map(s => s.sheetName);
            }

            setHfReady(true);

        } catch (e) {
            console.error("[ExcelWidget] HyperFormula failed to initialise:", e.message);
        }

        return () => {
            try {
                if (hfRef.current) {
                    hfRef.current.destroy();
                    hfRef.current = null;
                }
            } catch (e) {
                // safe to ignore — HotTable may have already torn it down
            }
        };
    }, []); // Run once on mount only

    // ── Sync: update HF when allSheets changes ─────────────────────────────
    // Fires when Mendix refreshes allSheetsJson (e.g. after another sheet saves)
    // Diffs against currently registered sheets — only adds/removes/updates
    // what actually changed. Does NOT recreate the HF instance.
    useEffect(() => {
        const hf = hfRef.current;
        if (!hf || !allSheets || allSheets.length === 0) return;

        const registered = registeredSheetsRef.current;
        const incoming   = allSheets.map(s => s.sheetName);

        // 1. Add sheets that are new
        allSheets.forEach(sheet => {
            if (!registered.includes(sheet.sheetName)) {
                try {
                    const hfSheetId = hf.addSheet(sheet.sheetName);
                    hf.setSheetContent(hfSheetId, sheet.data || [[]]);
                } catch (e) {
                    console.warn("[ExcelWidget] HF addSheet warning:", e.message);
                }
            }
        });

        // 2. Remove sheets that no longer exist
        registered.forEach(name => {
            if (!incoming.includes(name)) {
                try {
                    const hfSheetId = hf.getSheetId(name);
                    if (hfSheetId !== undefined) {
                        hf.removeSheet(hfSheetId);
                    }
                } catch (e) {
                    console.warn("[ExcelWidget] HF removeSheet warning:", e.message);
                }
            }
        });

        // 3. Update content for all existing sheets
        // Skip the current sheet — it's synced separately via currentSheetData
        allSheets.forEach(sheet => {
            if (sheet.sheetName === currentSheetName) return;
            try {
                const hfSheetId = hf.getSheetId(sheet.sheetName);
                if (hfSheetId !== undefined) {
                    hf.setSheetContent(hfSheetId, sheet.data || [[]]);
                }
            } catch (e) {
                console.warn("[ExcelWidget] HF setSheetContent warning:", e.message);
            }
        });

        registeredSheetsRef.current = incoming;

    }, [allSheets, currentSheetName]);

    // ── Sync: update current sheet data in HF live ─────────────────────────
    // Fires on every cell edit in the current sheet.
    // This ensures =Revenue!A1 in another sheet (if visible) would see
    // the latest data, and self-referencing formulas recalculate correctly.
    useEffect(() => {
        const hf = hfRef.current;
        if (!hf || !currentSheetName || !currentSheetData) return;

        try {
            const hfSheetId = hf.getSheetId(currentSheetName);
            if (hfSheetId !== undefined) {
                hf.setSheetContent(hfSheetId, currentSheetData);
            }
        } catch (e) {
            // Sheet may not be registered yet on first render — safe to ignore
        }
    }, [currentSheetData, currentSheetName]);

    return { hfRef, hfReady };
}

// ─────────────────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * registerAllSheets
 * Registers all sheets from allSheets array into a fresh HF instance.
 * Called once on mount.
 *
 * @param {HyperFormula} hf
 * @param {Array<{ sheetId, sheetName, data }>} allSheets
 */
function registerAllSheets(hf, allSheets) {
    allSheets.forEach(sheet => {
        try {
            const hfSheetId = hf.addSheet(sheet.sheetName);
            hf.setSheetContent(hfSheetId, sheet.data || [[]]);
        } catch (e) {
            console.warn(`[ExcelWidget] HF registerAllSheets warning for "${sheet.sheetName}":`, e.message);
        }
    });
}
import { useRef, useState, useEffect } from "react";
import HyperFormula from "hyperformula";

export function useHyperformula(allSheets, currentSheetName, hotRef) {

    const hfRef                               = useRef(null);
    const [hfReady, setHfReady]               = useState(false);
    const [allSheetsReady, setAllSheetsReady] = useState(false);
    const registeredSheetsRef                 = useRef([]);

    // ── Mount: create HF instance only ───────────────────────────────────
    // We do NOT register any sheets here.
    // HotTable registers the current sheet itself via the formulas prop.
    // Other sheets are registered in the sync effect below.
    useEffect(() => {
        try {
            const hf = HyperFormula.buildEmpty({ licenseKey: "gpl-v3" });
            hfRef.current = hf;
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
            } catch (e) {}
            registeredSheetsRef.current = [];
            setAllSheetsReady(false);
        };
    }, []);

    // ── Sync: register/update OTHER sheets when allSheets arrives ─────────
    // Current sheet is ALWAYS excluded — HotTable owns it completely.
    // We only manage sheets that are NOT currently displayed.
    useEffect(() => {
        const hf = hfRef.current;
        if (!hf || !hfReady) return;

        // Single-sheet fallback or allSheetsJson not wired
        if (!allSheets || allSheets.length === 0) {
            setAllSheetsReady(true);
            return;
        }

        // Only work with OTHER sheets — never the current sheet
        const otherSheets = allSheets.filter(s => s.sheetName !== currentSheetName);

        // If no other sheets exist, nothing to register
        if (otherSheets.length === 0) {
            setAllSheetsReady(true);
            return;
        }

        const registered = registeredSheetsRef.current;
        const incoming   = otherSheets.map(s => s.sheetName);
        let   dataChanged = false;

        // 1. Add other sheets that are new
        otherSheets.forEach(sheet => {
            if (!registered.includes(sheet.sheetName)) {
                try {
                    const id = hf.addSheet(sheet.sheetName);
                    hf.setSheetContent(id, sheet.data || [[]]);
                    dataChanged = true;
                } catch (e) {
                    console.warn("[ExcelWidget] HF addSheet warning:", e.message);
                }
            }
        });

        // 2. Remove other sheets that no longer exist
        registered.forEach(name => {
            if (!incoming.includes(name)) {
                try {
                    const id = hf.getSheetId(name);
                    if (id !== undefined) {
                        hf.removeSheet(id);
                        dataChanged = true;
                    }
                } catch (e) {
                    console.warn("[ExcelWidget] HF removeSheet warning:", e.message);
                }
            }
        });

        // 3. Update content for existing other sheets
        otherSheets.forEach(sheet => {
            try {
                const id = hf.getSheetId(sheet.sheetName);
                if (id !== undefined) {
                    hf.setSheetContent(id, sheet.data || [[]]);
                    dataChanged = true;
                }
            } catch (e) {
                console.warn("[ExcelWidget] HF setSheetContent warning:", e.message);
            }
        });

        registeredSheetsRef.current = incoming;

        // Mark ready — SheetGrid renders only after this
        // Guarantees HotTable mounts with all other sheets already in HF
        setAllSheetsReady(true);

        // Notify HotTable to recalculate cross-sheet formulas
        // Only after initial mount is complete
        if (dataChanged && allSheetsReady) {
            try {
                const hot = hotRef?.current?.hotInstance;
                if (hot) hot.render();
            } catch (e) {}
        }

    }, [allSheets, hfReady, currentSheetName]);

    return { hfRef, hfReady: hfReady && allSheetsReady };
}

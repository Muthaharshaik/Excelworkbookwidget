import { useRef, useState, useEffect } from "react";
import { HyperFormula } from "hyperformula";

export function useHyperformula(allSheets, currentSheetName, hotRef) {

    const hfRef                               = useRef(null);
    const [hfReady, setHfReady]               = useState(false);
    const [allSheetsReady, setAllSheetsReady] = useState(false);

    // Track sheets currently registered inside HF
    const registeredSheetsRef = useRef(new Set());

    // ─────────────────────────────────────────────────────────────
    // Create HF engine once
    // ─────────────────────────────────────────────────────────────
    useEffect(() => {
        try {
            const hf = HyperFormula.buildEmpty({
                licenseKey: "gpl-v3"
            });

            hfRef.current = hf;
            setHfReady(true);

        } catch (e) {
            console.error("[ExcelWidget] HyperFormula failed to initialise:", e.message);
        }

        return () => {
            try {
                hfRef.current?.destroy();
            } catch {}

            hfRef.current = null;
            registeredSheetsRef.current.clear();
            setAllSheetsReady(false);
        };
    }, []);

    // ─────────────────────────────────────────────────────────────
    // Sync OTHER sheets into HyperFormula
    // (current sheet is always owned by Handsontable)
    // ─────────────────────────────────────────────────────────────
    useEffect(() => {

        const hf = hfRef.current;
        if (!hf || !hfReady) return;

        if (!Array.isArray(allSheets) || allSheets.length === 0) {
            setAllSheetsReady(true);
            return;
        }

        const registered = registeredSheetsRef.current;

        // Only sync sheets that are NOT currently displayed
        const otherSheets = allSheets.filter(
            s => s.sheetName !== currentSheetName
        );

        let recalculationNeeded = false;

        // ── Ensure all other sheets exist and update their data
        otherSheets.forEach(sheet => {

            try {

                const sheetName = sheet.sheetName;
                const existingId = hf.getSheetId(sheetName);

                // Sheet not yet registered → create
                if (existingId === undefined) {

                    const newId = hf.addSheet(sheetName);
                    hf.setSheetContent(newId, sheet.data || [[]]);

                    registered.add(sheetName);
                    recalculationNeeded = true;
                    return;
                }

                // Sheet exists → update content
                hf.setSheetContent(existingId, sheet.data || [[]]);
                registered.add(sheetName);
                recalculationNeeded = true;

            } catch (err) {
                console.warn("[ExcelWidget] HF sync warning:", err.message);
            }

        });

        // ── Remove sheets that no longer exist in Mendix
        registered.forEach(sheetName => {

            const stillExists = otherSheets.some(
                s => s.sheetName === sheetName
            );

            if (!stillExists) {

                try {
                    const id = hf.getSheetId(sheetName);

                    if (id !== undefined) {
                        hf.removeSheet(id);
                    }

                } catch {}

                registered.delete(sheetName);
                recalculationNeeded = true;
            }

        });

        setAllSheetsReady(true);

        // Trigger Handsontable re-render if HF graph changed
        if (recalculationNeeded) {

            try {
                const hot = hotRef?.current?.hotInstance;

                if (hot) {
                    hot.render();
                }

            } catch {}

        }

    }, [allSheets, currentSheetName, hfReady]);

    return {
        hfRef,
        hfReady: hfReady && allSheetsReady
    };
}

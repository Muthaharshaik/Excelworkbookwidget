/**
 * useHyperformula.js — FINAL
 *
 * Manages HF engine and other-sheet data for cross-sheet formula resolution.
 * No sheet removal needed — the bleed fix is entirely in SheetGrid's data prop.
 */

import { useRef, useState, useEffect } from "react";
import { HyperFormula } from "hyperformula";

export function useHyperformula(allSheets, currentSheetName, hotRef) {

    const hfRef                 = useRef(null);
    const [hfReady, setHfReady] = useState(false);
    const registeredSheetsRef   = useRef(new Set());
    const currentSheetNameRef   = useRef(currentSheetName);
    currentSheetNameRef.current = currentSheetName;

    useEffect(() => {
        try {
            const hf = HyperFormula.buildEmpty({ licenseKey: "gpl-v3" });
            hfRef.current = hf;
            setHfReady(true);
        } catch (e) {
            console.info("[HF] Engine creation failed:", e.message);
        }
        return () => {
            try { hfRef.current?.destroy(); } catch {}
            hfRef.current = null;
            registeredSheetsRef.current.clear();
            setHfReady(false);
        };
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    useEffect(() => {
        const hf = hfRef.current;
        if (!hf || !hfReady) return;
        if (!Array.isArray(allSheets) || allSheets.length === 0) return;

        const current    = currentSheetNameRef.current;
        const registered = registeredSheetsRef.current;
        const others     = allSheets.filter(s => s.sheetName !== current);

        others.forEach(sheet => {
            try {
                const safeData = (Array.isArray(sheet.data) && sheet.data.length > 0) ? sheet.data : [[]];
                if (hf.doesSheetExist(sheet.sheetName)) {
                    hf.setSheetContent(hf.getSheetId(sheet.sheetName), safeData);
                } else {
                    hf.addSheet(sheet.sheetName);
                    hf.setSheetContent(hf.getSheetId(sheet.sheetName), safeData);
                }
                registered.add(sheet.sheetName);
            } catch (err) {
                console.info(`[HF] sync error "${sheet.sheetName}":`, err.message);
            }
        });

        registered.forEach(sheetName => {
            if (sheetName === current) return;
            if (!others.some(s => s.sheetName === sheetName)) {
                try {
                    if (hf.doesSheetExist(sheetName)) hf.removeSheet(hf.getSheetId(sheetName));
                } catch {}
                registered.delete(sheetName);
            }
        });

        try {
            const hot = hotRef?.current?.hotInstance;
            if (hot) hot.render();
        } catch {}

    }, [allSheets, hfReady, hotRef]);

    return { hfRef, hfReady };
}
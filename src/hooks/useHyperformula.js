/**
 * useHyperformula.js
 *
 * Creates ONE HyperFormula instance per WorkbookContainer mount.
 *
 * WHY useEffect + hfReady STATE:
 *   HotTable only reads the formulas prop on initial mount.
 *   We use hfReady state to gate SheetGrid rendering — SheetGrid
 *   is not rendered until HF is confirmed ready, guaranteeing
 *   HotTable always mounts with a live HF engine attached.
 *
 * WHY ONE INSTANCE PER WIDGET:
 *   Each widget instance = one sheet. HF instance is scoped to
 *   that widget. Cross-sheet references between different Mendix
 *   DataView widgets are a future enhancement.
 *
 * LICENSE:
 *   HyperFormula is open source under GPL-v3.
 *   "gpl-v3" is the correct free license key for it.
 *   "non-commercial-and-evaluation" is HotTable's key — not HF's.
 *
 * BUILD NOTE:
 *   HyperFormula's dependency chevrotain uses eval() internally.
 *   Mendix's rollup pipeline blocks eval by default.
 *   This is solved by rollup.config.mjs in the widget root which
 *   suppresses the eval warning specifically for chevrotain.
 */

import { useRef, useState, useEffect } from "react";
import HyperFormula from "hyperformula";

/**
 * @returns {{
 *   hfRef:   React.MutableRefObject<HyperFormula|null>,
 *   hfReady: boolean
 * }}
 */
export function useHyperformula() {

    const hfRef = useRef(null);
    const [hfReady, setHfReady] = useState(false);

    useEffect(() => {
        try {
            const hf = HyperFormula.buildEmpty({
                licenseKey: "gpl-v3",
            });
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
            } catch (e) {
                // HotTable may have already torn it down — safe to ignore
            }
        };
    }, []);

    return { hfRef, hfReady };
}
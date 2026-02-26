/**
 * mendixBridge.js
 *
 * Single file responsible for ALL communication from the widget → Mendix.
 *
 * WHY THIS EXISTS:
 *   Mendix action objects (the things you wire in Studio Pro) have a specific
 *   contract: { canExecute: boolean, execute: () => void }.
 *   Rather than sprinkling `action?.canExecute && action.execute()` checks
 *   all over the codebase, every Mendix call goes through this file.
 *
 *   Benefits:
 *     - One place to add error handling / logging for all Mendix calls
 *     - Easy to mock in tests
 *     - If Mendix changes their action API, only this file needs updating
 *
 * FUNCTIONS IN THIS FILE:
 *   triggerSheetChange    — auto-save: write JSON back + fire commit microflow
 *   triggerSheetTabChange — notify Mendix when user switches sheet tabs
 *
 * WHAT THIS FILE DOES NOT DO:
 *   - Does not manage state
 *   - Does not debounce (useAutoSave owns debouncing)
 *   - Does not serialize data (dataService.js owns that)
 */

// ─────────────────────────────────────────────────────────────────────────────
//  triggerSheetChange
// ─────────────────────────────────────────────────────────────────────────────

/**
 * triggerSheetChange
 *
 * Called by useAutoSave after debounce fires.
 * Does two things in sequence:
 *   1. Writes the updated JSON string back into the Mendix sheetsJson attribute
 *      via setValue() — this marks the Mendix object as dirty in client memory
 *   2. Fires the onSheetChange microflow — Mendix commits the dirty object to DB
 *
 * @param {object} sheetsJsonAttr   - Mendix EditableValue for the sheetsJson attribute
 *                                    Must have { setValue, status } — only works with
 *                                    XPath datasources, not Microflow datasources
 * @param {string} serializedJson   - the JSON string produced by dataService.serializeSheets()
 * @param {object} onSheetChange    - Mendix action object { canExecute, execute }
 *                                    wired to the onSheetChange prop in Studio Pro
 *
 * @returns {boolean}               - true if both setValue and execute succeeded,
 *                                    false if anything was skipped/failed
 */
export function triggerSheetChange(sheetsJsonAttr, serializedJson, onSheetChange) {

    // ── Step 1: Write new JSON into Mendix attribute ───────────────────────
    if (!sheetsJsonAttr) {
        console.error(
            "[ExcelWidget] triggerSheetChange: sheetsJsonAttr is not provided. " +
            "Check that sheetsJson is wired to a Mendix attribute in Studio Pro."
        );
        return false;
    }

    if (sheetsJsonAttr.status !== "available") {
        console.warn(
            "[ExcelWidget] triggerSheetChange: sheetsJsonAttr status is " +
            `"${sheetsJsonAttr.status}" — skipping save. ` +
            "Attribute may still be loading."
        );
        return false;
    }

    if (typeof sheetsJsonAttr.setValue !== "function") {
        console.error(
            "[ExcelWidget] triggerSheetChange: sheetsJsonAttr.setValue is not a function. " +
            "The datasource MUST be XPath-based. Microflow datasources are read-only " +
            "and do not support setValue()."
        );
        return false;
    }

    try {
        sheetsJsonAttr.setValue(serializedJson);
    } catch (err) {
        console.error("[ExcelWidget] triggerSheetChange: setValue() threw an error:", err.message);
        return false;
    }

    // ── Step 2: Fire the commit microflow ─────────────────────────────────
    if (!onSheetChange) {
        // Not wired in Studio Pro — not necessarily an error, just a warning.
        // The setValue above still marks the object dirty in Mendix memory.
        // It just won't be committed to DB until something else triggers a commit.
        console.warn(
            "[ExcelWidget] triggerSheetChange: onSheetChange action is not wired. " +
            "Data was written to the Mendix attribute but will NOT be committed to the " +
            "database. Wire the onSheetChange prop to a commit microflow in Studio Pro."
        );
        return false;
    }

    if (!onSheetChange.canExecute) {
        console.warn(
            "[ExcelWidget] triggerSheetChange: onSheetChange.canExecute is false. " +
            "The microflow is blocked — check security roles or preconditions."
        );
        return false;
    }

    try {
        onSheetChange.execute();
        return true;
    } catch (err) {
        console.error("[ExcelWidget] triggerSheetChange: onSheetChange.execute() threw:", err.message);
        return false;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  triggerSheetTabChange
// ─────────────────────────────────────────────────────────────────────────────

/**
 * triggerSheetTabChange
 *
 * Fired when the user clicks a different sheet tab.
 * Optional — only fires if the onSheetTabChange prop is wired in Studio Pro.
 * Mendix can use this to track the active sheet, update URL params, etc.
 *
 * @param {object} onSheetTabChange  - Mendix action object { canExecute, execute }
 *
 * @returns {boolean}
 */
export function triggerSheetTabChange(onSheetTabChange) {
    if (!onSheetTabChange) return false;

    if (!onSheetTabChange.canExecute) {
        console.warn(
            "[ExcelWidget] triggerSheetTabChange: onSheetTabChange.canExecute is false."
        );
        return false;
    }

    try {
        onSheetTabChange.execute();
        return true;
    } catch (err) {
        console.error("[ExcelWidget] triggerSheetTabChange: execute() threw:", err.message);
        return false;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
//  isMendixActionReady  (utility)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * isMendixActionReady
 *
 * Utility to check whether a Mendix action is safe to call.
 * Use this anywhere you want to conditionally show/hide UI based on
 * whether an action is wired and executable.
 *
 * @param   {object}  action  - Mendix action object
 * @returns {boolean}
 */
export function isMendixActionReady(action) {
    return !!(action && action.canExecute);
}

// ─────────────────────────────────────────────────────────────────────────────
//  isMendixAttrWritable  (utility)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * isMendixAttrWritable
 *
 * Utility to check whether a Mendix EditableValue attribute
 * is available and writable (has setValue).
 *
 * @param   {object}  attr  - Mendix EditableValue
 * @returns {boolean}
 */
export function isMendixAttrWritable(attr) {
    return !!(
        attr &&
        attr.status === "available" &&
        typeof attr.setValue === "function"
    );
}
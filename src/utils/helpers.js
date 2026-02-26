/**
 * helpers.js
 *
 * Small, pure utility functions used across the widget.
 * No React, no Mendix, no side effects — pure JS only.
 * Easy to unit test in isolation.
 */

// ─── Identity ─────────────────────────────────────────────────────────────────

/**
 * generateId
 * Generates a lightweight unique ID for new sheets created client-side.
 * NOT a full UUID — good enough for temporary client IDs before Mendix assigns a GUID.
 *
 * @returns {string}  e.g. "sheet-1k3x9z-4f2"
 */
export function generateId(prefix = "id") {
    return `${prefix}-${Math.random().toString(36).slice(2, 8)}-${Date.now().toString(36)}`;
}

// ─── Object Utilities ─────────────────────────────────────────────────────────

/**
 * deepClone
 * Safe deep clone using JSON round-trip.
 * Fine for our data shapes (no functions, no Dates, no Sets).
 *
 * @param   {any} obj
 * @returns {any}
 */
export function deepClone(obj) {
    if (obj === null || obj === undefined) return obj;
    try {
        return JSON.parse(JSON.stringify(obj));
    } catch {
        return obj;
    }
}

/**
 * isEmpty
 * Returns true if a value is null, undefined, empty string, or empty array.
 *
 * @param   {any}     value
 * @returns {boolean}
 */
export function isEmpty(value) {
    if (value === null || value === undefined) return true;
    if (typeof value === "string" && value.trim() === "") return true;
    if (Array.isArray(value) && value.length === 0) return true;
    return false;
}

// ─── Cell Address Utilities ───────────────────────────────────────────────────

/**
 * cellKey
 * Produces the "row,col" string key used in cellMeta objects.
 * e.g. cellKey(0, 2) → "0,2"
 *
 * @param   {number} row
 * @param   {number} col
 * @returns {string}
 */
export function cellKey(row, col) {
    return `${row},${col}`;
}

/**
 * parseCellKey
 * Reverses cellKey — splits "row,col" back into { row, col }.
 *
 * @param   {string}          key  e.g. "3,5"
 * @returns {{ row, col }}
 */
export function parseCellKey(key) {
    const [row, col] = key.split(",").map(Number);
    return { row, col };
}

/**
 * colIndexToLetter
 * Converts a zero-based column index to Excel-style column letter.
 * e.g. 0 → "A", 25 → "Z", 26 → "AA"
 *
 * @param   {number} index  - zero-based column index
 * @returns {string}
 */
export function colIndexToLetter(index) {
    let result = "";
    let n = index;
    while (n >= 0) {
        result = String.fromCharCode((n % 26) + 65) + result;
        n = Math.floor(n / 26) - 1;
    }
    return result;
}

/**
 * letterToColIndex
 * Reverses colIndexToLetter.
 * e.g. "A" → 0, "Z" → 25, "AA" → 26
 *
 * @param   {string} letter  e.g. "AA"
 * @returns {number}         zero-based column index
 */
export function letterToColIndex(letter) {
    let result = 0;
    for (let i = 0; i < letter.length; i++) {
        result = result * 26 + (letter.charCodeAt(i) - 64);
    }
    return result - 1;
}

// ─── Array Utilities ──────────────────────────────────────────────────────────

/**
 * clampIndex
 * Clamps a number between min and max (inclusive).
 * Used to keep activeSheetIndex in bounds when sheets are removed.
 *
 * @param   {number} value
 * @param   {number} min
 * @param   {number} max
 * @returns {number}
 */
export function clampIndex(value, min, max) {
    return Math.min(Math.max(value, min), max);
}

/**
 * reorderArray
 * Moves an item from one index to another in an array.
 * Returns a NEW array — does not mutate input.
 * Used for sheet tab drag-to-reorder.
 *
 * @param   {any[]}  arr
 * @param   {number} fromIndex
 * @param   {number} toIndex
 * @returns {any[]}
 */
export function reorderArray(arr, fromIndex, toIndex) {
    if (fromIndex === toIndex) return arr;
    const result = [...arr];
    const [moved] = result.splice(fromIndex, 1);
    result.splice(toIndex, 0, moved);
    return result;
}

// ─── String Utilities ─────────────────────────────────────────────────────────

/**
 * truncate
 * Truncates a string to maxLength, appending "…" if cut.
 * Used for sheet tab name display.
 *
 * @param   {string} str
 * @param   {number} maxLength  default 20
 * @returns {string}
 */
export function truncate(str, maxLength = 20) {
    if (!str) return "";
    return str.length > maxLength ? `${str.slice(0, maxLength - 1)}…` : str;
}

// ─── Debounce ────────────────────────────────────────────────────────────────

/**
 * debounce
 * Classic debounce. Returns a function that delays invoking `fn`
 * until after `delay` ms have elapsed since the last call.
 *
 * Used in useAutoSave to avoid hammering Mendix on every keystroke.
 *
 * @param   {Function} fn
 * @param   {number}   delay  ms
 * @returns {Function}
 */
export function debounce(fn, delay) {
    let timer;
    return function (...args) {
        clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), delay);
    };
}
/**
 * constants.js
 *
 * Single source of truth for all magic values used across the widget.
 * Never hardcode these inline — always import from here.
 */

// ─── Handsontable License ─────────────────────────────────────────────────────

export const HOT_LICENSE_KEY = "non-commercial-and-evaluation";

// ─── Auto-Save ────────────────────────────────────────────────────────────────

/**
 * How long (ms) to wait after the last cell edit before firing
 * the onSheetChange Mendix action.
 * 800ms is a good balance — fast enough to feel instant, slow enough
 * to not spam the server on rapid typing.
 */
export const AUTOSAVE_DEBOUNCE_MS = 800;

// ─── Grid Defaults ────────────────────────────────────────────────────────────

export const DEFAULT_COL_WIDTH    = 120;   // pixels
export const DEFAULT_ROW_HEIGHT   = 23;    // pixels — matches Excel default
export const DEFAULT_GRID_HEIGHT  = 600;   // pixels — overridden by widget prop
export const MIN_ROWS             = 50;    // empty rows shown below data
export const MIN_COLS             = 26;    // A–Z visible by default

// ─── Sheet Defaults ───────────────────────────────────────────────────────────

/**
 * Default empty sheet data when a new sheet is created with no data.
 * 50 rows × 26 cols of null values.
 */
export const EMPTY_SHEET_DATA = () =>
    Array.from({ length: MIN_ROWS }, () => Array(MIN_COLS).fill(null));

/**
 * Default new sheet object shape.
 * Used as a template when Mendix sends a sheet with missing fields.
 */
export const DEFAULT_SHEET = {
    sheetId:     "",
    sheetName:   "Sheet",
    orderIndex:  0,
    isEditable:  false,     // default safe — must be explicitly granted
    data:        null,      // null means use EMPTY_SHEET_DATA()
    cellMeta:    {},
    colWidths:   [],
    rowHeights:  [],
    mergedCells: [],
};

// ─── Permission Levels ────────────────────────────────────────────────────────

export const PERMISSION = Object.freeze({
    VIEW: "view",
    EDIT: "edit",
});

// ─── Handsontable Context Menu Items ─────────────────────────────────────────
/**
 * Which context menu items to show.
 * 'true' means show all. Array means show only listed items.
 * Full list: https://handsontable.com/docs/react-data-grid/context-menu/
 */
export const CONTEXT_MENU_ITEMS = [
    "row_above",
    "row_below",
    "---------",
    "col_left",
    "col_right",
    "---------",
    "remove_row",
    "remove_col",
    "---------",
    "undo",
    "redo",
    "---------",
    "make_read_only",
    "alignment",
    "---------",
    "copy",
    "cut",
];

// ─── Supported Cell Types ─────────────────────────────────────────────────────

export const CELL_TYPES = Object.freeze({
    TEXT:     "text",
    NUMERIC:  "numeric",
    DATE:     "date",
    CHECKBOX: "checkbox",
    DROPDOWN: "dropdown",
    PASSWORD: "password",
});

// ─── Toolbar Actions (used as identifiers) ────────────────────────────────────

export const TOOLBAR_ACTION = Object.freeze({
    BOLD:            "bold",
    ITALIC:          "italic",
    UNDERLINE:       "underline",
    FONT_COLOR:      "fontColor",
    BG_COLOR:        "bgColor",
    ALIGN_LEFT:      "alignLeft",
    ALIGN_CENTER:    "alignCenter",
    ALIGN_RIGHT:     "alignRight",
    MERGE_CELLS:     "mergeCells",
    UNMERGE_CELLS:   "unmergeCells",
    BORDERS:         "borders",
    CLEAR_FORMAT:    "clearFormat",
});

// ─── CSS Class Names ──────────────────────────────────────────────────────────
// Centralised so renaming in CSS only requires changing here

export const CSS = Object.freeze({
    WORKBOOK_ROOT:     "eww-root",
    HEADER:            "eww-header",
    TOOLBAR:           "eww-toolbar",
    GRID_WRAPPER:      "eww-grid-wrapper",
    TAB_BAR:           "eww-tab-bar",
    TAB:               "eww-tab",
    TAB_ACTIVE:        "eww-tab--active",
    TAB_READONLY:      "eww-tab--readonly",
    READONLY_BADGE:    "eww-readonly-badge",
    SAVING_INDICATOR:  "eww-saving-indicator",
});
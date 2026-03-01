/**
 * constants.js
 * Single source of truth for all magic values used across the widget.
 */

export const HOT_LICENSE_KEY = "non-commercial-and-evaluation";

export const AUTOSAVE_DEBOUNCE_MS = 800;

export const DEFAULT_COL_WIDTH   = 120;
export const DEFAULT_ROW_HEIGHT  = 23;
export const DEFAULT_GRID_HEIGHT = 600;
export const MIN_ROWS            = 50;
export const MIN_COLS            = 26;

export const EMPTY_SHEET_DATA = () =>
    Array.from({ length: MIN_ROWS }, () => Array(MIN_COLS).fill(null));

export const DEFAULT_SHEET = {
    sheetId:     "",
    sheetName:   "Sheet",
    orderIndex:  0,
    isEditable:  false,
    data:        null,
    columns:     [],
    rowLabels:   [],
    cellMeta:    {},
    colWidths:   [],
    rowHeights:  [],
    mergedCells: [],
};

export const PERMISSION = Object.freeze({
    VIEW: "view",
    EDIT: "edit",
});

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

export const CELL_TYPES = Object.freeze({
    TEXT:     "text",
    NUMERIC:  "numeric",
    DATE:     "date",
    CHECKBOX: "checkbox",
    DROPDOWN: "dropdown",
});

/**
 * Column type metadata for ColumnSettingsPanel UI.
 * Time removed — not reliably supported without Moment.js.
 * Icons/emojis removed — plain labels only.
 */
export const COLUMN_TYPE_META = [
    { value: "text",     label: "Text",     hotType: "text",     hasSource: false, hasFormat: false },
    { value: "numeric",  label: "Number",   hotType: "numeric",  hasSource: false, hasFormat: true  },
    { value: "date",     label: "Date",     hotType: "date",     hasSource: false, hasFormat: true  },
    { value: "checkbox", label: "Checkbox", hotType: "checkbox", hasSource: false, hasFormat: false },
    { value: "dropdown", label: "Dropdown", hotType: "dropdown", hasSource: true,  hasFormat: false },
];

export const DEFAULT_NUMERIC_FORMAT = "0,0.00";
export const DEFAULT_DATE_FORMAT    = "DD/MM/YYYY";

export const DEFAULT_COLUMN = {
    key:      "",
    header:   "Column",
    type:     "text",
    width:    120,
    source:   [],
    format:   "",
    readOnly: false,
};

export const TOOLBAR_ACTION = Object.freeze({
    BOLD:          "bold",
    ITALIC:        "italic",
    UNDERLINE:     "underline",
    FONT_COLOR:    "fontColor",
    BG_COLOR:      "bgColor",
    ALIGN_LEFT:    "alignLeft",
    ALIGN_CENTER:  "alignCenter",
    ALIGN_RIGHT:   "alignRight",
    MERGE_CELLS:   "mergeCells",
    UNMERGE_CELLS: "unmergeCells",
    BORDERS:       "borders",
    CLEAR_FORMAT:  "clearFormat",
});

export const CSS = Object.freeze({
    WORKBOOK_ROOT:    "eww-root",
    HEADER:           "eww-header",
    TOOLBAR:          "eww-toolbar",
    GRID_WRAPPER:     "eww-grid-wrapper",
    TAB_BAR:          "eww-tab-bar",
    TAB:              "eww-tab",
    TAB_ACTIVE:       "eww-tab--active",
    TAB_READONLY:     "eww-tab--readonly",
    READONLY_BADGE:   "eww-readonly-badge",
    SAVING_INDICATOR: "eww-saving-indicator",
});
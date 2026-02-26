/**
 * usePermissions.js
 *
 * Resolves whether the current user can edit a given sheet.
 *
 * TWO LEVELS of permission:
 *
 *   Level 1 — Workbook level (isReadOnly prop from Mendix)
 *     If true, NO sheet in this workbook is editable.
 *     Mendix computes this server-side from SheetPermission records.
 *
 *   Level 2 — Sheet level (isEditable flag inside each sheet object)
 *     Per-sheet flag set by the admin.
 *     Only meaningful if workbook-level isReadOnly is false.
 *
 * LOGIC:
 *   canEdit = !workbookIsReadOnly && sheet.isEditable
 *
 * WHY A HOOK:
 *   Future phases can extend this (e.g. time-based access, row-level locks).
 *   Centralising here means WorkbookContainer never does permission math.
 *
 * @param   {boolean}  workbookIsReadOnly  - from Mendix isReadOnly prop
 * @returns {{ canEditSheet: (sheetIsEditable: boolean) => boolean }}
 */
import { useCallback } from "react";

export function usePermissions(workbookIsReadOnly) {

    /**
     * canEditSheet
     *
     * @param   {boolean} sheetIsEditable  - the isEditable flag from the sheet object
     * @returns {boolean}
     */
    const canEditSheet = useCallback((sheetIsEditable) => {
        // Workbook-level lock overrides everything
        if (workbookIsReadOnly) return false;
        // Sheet-level flag
        return sheetIsEditable === true;
    }, [workbookIsReadOnly]);

    return { canEditSheet };
}
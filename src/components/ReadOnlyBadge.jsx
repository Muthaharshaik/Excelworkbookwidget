/**
 * ReadOnlyBadge.jsx
 *
 * Small pill badge shown in the workbook header when the
 * currently active sheet is view-only for this user.
 *
 * Kept as its own file because:
 *  - It's reusable (could appear in toolbar, sheet grid header too)
 *  - WorkbookContainer stays clean
 */

import { createElement } from "react";
import { CSS }           from "../utils/constants";

export function ReadOnlyBadge() {
    return (
        <span className={CSS.READONLY_BADGE} aria-label="This sheet is read only">
            <span className={`${CSS.READONLY_BADGE}__icon`} aria-hidden="true">🔒</span>
            View Only
        </span>
    );
}
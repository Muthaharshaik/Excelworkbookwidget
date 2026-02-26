/**
 * ExcelWorkbookWidget.jsx
 *
 * Mendix entry point — this is the file Mendix calls directly.
 *
 * RULE: Keep this file completely thin.
 *   - No state
 *   - No logic
 *   - No hooks
 *   - Just import CSS, import WorkbookContainer, pass all props through.
 *
 * WHY: Mendix re-renders this component whenever any prop changes.
 * Keeping logic out of here means nothing accidentally re-runs on
 * unrelated prop updates.
 */

import { createElement } from "react";
import "./ui/ExcelWorkbookWidget.css";
import { WorkbookContainer } from "./components/WorkbookContainer";

export function ExcelWorkbookWidget(props) {
    return <WorkbookContainer {...props} />;
}
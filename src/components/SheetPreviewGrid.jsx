import { useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { extractThemeColors, getCellFillColor } from "../lib/cellFill.js";

// Fallback window size when the sheet is too large to render in full (see
// MAX_CELLS below) — centered on the target cell.
const VISIBLE_ROWS = 150;
const VISIBLE_COLS = 40;

// A sheet's !ref range is often much bigger than its real data (Excel keeps
// stray formatting on cells far outside actual content), so rendering
// "however big !ref claims to be" isn't safe as a blind default — a
// worksheet reporting e.g. A1:AZ50000 would mean 1.3M <td> elements and lock
// up the tab. Below this cell-count budget we render the sheet's whole
// range; above it we fall back to a centered window around the target.
const MAX_CELLS = 6000;

const InfoIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
  </svg>
);

const clamp = (v, min, max) => Math.max(min, Math.min(max, v));

const cellDisplay = (cell) => {
  if (!cell) return "";
  if (cell.w !== undefined) return cell.w;
  if (cell.v !== undefined) return String(cell.v);
  return "";
};

/**
 * Renders a scrollable window of a worksheet's grid (row/column headers,
 * cell values) with one cell highlighted — lets you see exactly where a
 * dependent/dependency actually sits on its sheet.
 *
 * When onCellClick is given, non-target cells become clickable — lets you
 * jump straight from the preview into Find for that cell.
 */
const SheetPreviewGrid = ({ workbook, sheet, cellAddress, onCellClick }) => {
  const targetCellRef = useRef(null);

  const target = cellAddress
    ? XLSX.utils.decode_cell(cellAddress.split(":")[0].replace(/\$/g, ""))
    : null;

  const worksheet = sheet ? workbook?.Sheets[sheet] : null;

  const themeColors = useMemo(() => extractThemeColors(workbook), [workbook]);

  const sheetRange = useMemo(() => {
    if (!worksheet) return null;
    return worksheet["!ref"] ? XLSX.utils.decode_range(worksheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  }, [worksheet]);

  useEffect(() => {
    targetCellRef.current?.scrollIntoView({ block: "center", inline: "center" });
  }, [sheet, cellAddress]);

  if (!workbook || !sheet || !target || !worksheet || !sheetRange) {
    return (
      <div className="sheet-preview-empty">
        <InfoIcon />
        Click a cell in the results below to preview where it sits on its sheet.
      </div>
    );
  }

  const minRow = sheetRange.s.r;
  const minCol = sheetRange.s.c;
  const maxRow = sheetRange.e.r;
  const maxCol = sheetRange.e.c;

  const fitsInFull = (maxRow - minRow + 1) * (maxCol - minCol + 1) <= MAX_CELLS;

  const startR = fitsInFull ? minRow : clamp(target.r - Math.floor(VISIBLE_ROWS / 3), minRow, Math.max(minRow, maxRow - VISIBLE_ROWS + 1));
  const startC = fitsInFull ? minCol : clamp(target.c - Math.floor(VISIBLE_COLS / 3), minCol, Math.max(minCol, maxCol - VISIBLE_COLS + 1));
  const endR = fitsInFull ? maxRow : Math.min(startR + VISIBLE_ROWS - 1, maxRow);
  const endC = fitsInFull ? maxCol : Math.min(startC + VISIBLE_COLS - 1, maxCol);

  const rows = [];
  for (let r = startR; r <= endR; r++) rows.push(r);
  const cols = [];
  for (let c = startC; c <= endC; c++) cols.push(c);

  return (
    <div className="sheet-preview-wrapper">
      <table className="sheet-preview-grid">
        <thead>
          <tr>
            <th className="sheet-preview-corner"></th>
            {cols.map((c) => (
              <th key={c} className={c === target.c ? "sheet-preview-target-col" : undefined}>
                {XLSX.utils.encode_col(c)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r}>
              <th className={r === target.r ? "sheet-preview-target-row" : undefined}>{r + 1}</th>
              {cols.map((c) => {
                const addr = XLSX.utils.encode_cell({ r, c });
                const cell = worksheet[addr];
                const isTarget = r === target.r && c === target.c;
                const display = cellDisplay(cell);
                const clickable = Boolean(onCellClick) && !isTarget;
                let title = cell?.f ? `${addr}: =${cell.f}` : (display ? `${addr}: ${display}` : addr);
                if (clickable) title += " — click to navigate";
                const className = [
                  isTarget && "sheet-preview-target-cell",
                  clickable && "sheet-preview-clickable-cell",
                ].filter(Boolean).join(" ") || undefined;
                const fillColor = getCellFillColor(cell, themeColors);
                return (
                  <td
                    key={c}
                    ref={isTarget ? targetCellRef : undefined}
                    className={className}
                    title={title}
                    style={fillColor ? { backgroundColor: fillColor } : undefined}
                    onClick={clickable ? () => onCellClick(addr) : undefined}
                  >
                    {display}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default SheetPreviewGrid;

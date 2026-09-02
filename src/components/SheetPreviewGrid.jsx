import { useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

const VISIBLE_ROWS = 26;
const VISIBLE_COLS = 14;

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
 */
const SheetPreviewGrid = ({ workbook, sheet, cellAddress }) => {
  const targetCellRef = useRef(null);

  const target = cellAddress
    ? XLSX.utils.decode_cell(cellAddress.split(":")[0].replace(/\$/g, ""))
    : null;

  const worksheet = sheet ? workbook?.Sheets[sheet] : null;

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

  const maxRow = sheetRange.e.r;
  const maxCol = sheetRange.e.c;

  const startR = clamp(target.r - Math.floor(VISIBLE_ROWS / 3), 0, Math.max(0, maxRow - VISIBLE_ROWS + 1));
  const startC = clamp(target.c - Math.floor(VISIBLE_COLS / 3), 0, Math.max(0, maxCol - VISIBLE_COLS + 1));

  const rows = [];
  for (let r = startR; r <= Math.min(startR + VISIBLE_ROWS - 1, maxRow); r++) rows.push(r);
  const cols = [];
  for (let c = startC; c <= Math.min(startC + VISIBLE_COLS - 1, maxCol); c++) cols.push(c);

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
                const title = cell?.f ? `${addr}: =${cell.f}` : (display ? `${addr}: ${display}` : addr);
                return (
                  <td
                    key={c}
                    ref={isTarget ? targetCellRef : undefined}
                    className={isTarget ? "sheet-preview-target-cell" : undefined}
                    title={title}
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

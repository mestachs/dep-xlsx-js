import { useState, useImperativeHandle, forwardRef } from "react";
import * as XLSX from "xlsx";

const InfoIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
  </svg>
);

/**
 * Card that lets the user pick a sheet + cell and find every formula cell
 * in the workbook that references it.
 *
 * Imperative handle (via ref):
 *   ref.current.reset()   ← clears form and results, called when a new file loads
 */
const CellDependentsPanel = forwardRef(function CellDependentsPanel(
  { sheetNames, workbook },
  ref
) {
  const [selectedSheet, setSelectedSheet]   = useState("");
  const [cellCoordinate, setCellCoordinate] = useState("");
  const [dependents, setDependents]         = useState([]);

  useImperativeHandle(ref, () => ({
    reset: () => {
      setSelectedSheet("");
      setCellCoordinate("");
      setDependents([]);
    },
  }));

  const handleFind = () => {
    if (!selectedSheet || !cellCoordinate || !workbook) {
      setDependents([{ error: "Please select a sheet and enter a cell coordinate." }]);
      return;
    }

    const coord = cellCoordinate.toUpperCase();
    let targetCell;
    try {
      targetCell = XLSX.utils.decode_cell(coord);
    } catch {
      setDependents([{ error: "Invalid cell coordinate. Please use A1 format (e.g., A1, B2)." }]);
      return;
    }

    const results = [];
    const a1RefRegex = /(?:(?:'([^']+)'|([a-zA-Z0-9_]+))!)?(\$?[A-Z]+)(\$?\d+)(?::(\$?[A-Z]+)(\$?\d+))?/g;

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];

      for (const cellAddress in worksheet) {
        if (cellAddress.startsWith("!")) continue;

        const cell = worksheet[cellAddress];
        let formula = null;

        if (cell.f) {
          formula = cell.f;
        } else if (cell.F) {
          const range  = XLSX.utils.decode_range(cell.F);
          const topLeft = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
          formula = worksheet[topLeft]?.f ?? null;
        }

        if (!formula) continue;

        a1RefRegex.lastIndex = 0;
        let match;
        while ((match = a1RefRegex.exec(formula)) !== null) {
          const refSheet = match[1] || match[2] || sheetName;
          if (refSheet !== selectedSheet) continue;

          const [, , , startCol, startRow, endCol, endRow] = match;
          let hit = false;

          if (endCol && endRow) {
            const s = XLSX.utils.decode_cell(`${startCol}${startRow}`);
            const e = XLSX.utils.decode_cell(`${endCol}${endRow}`);
            hit = targetCell.c >= s.c && targetCell.c <= e.c &&
                  targetCell.r >= s.r && targetCell.r <= e.r;
          } else {
            const ref = XLSX.utils.decode_cell(`${startCol}${startRow}`);
            hit = ref.c === targetCell.c && ref.r === targetCell.r;
          }

          if (hit) {
            results.push({ sheetName, coordinates: cellAddress, formula });
            break;
          }
        }
      }
    }

    setDependents(results.length > 0 ? results : [{ message: "No dependents found for this cell." }]);
  };

  return (
    <div className="card cell-analysis">
      <div className="card-header">
        <h2>Find Cell Dependents</h2>
      </div>
      <div className="card-body">
        <div className="cell-analysis-fields">
          <div className="field-group">
            <label htmlFor="sheet-select">Sheet</label>
            <select
              id="sheet-select"
              value={selectedSheet}
              onChange={(e) => setSelectedSheet(e.target.value)}
            >
              <option value="">Select a sheet…</option>
              {sheetNames?.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </div>

          <div className="field-group">
            <label htmlFor="cell-input">Cell</label>
            <input
              id="cell-input"
              type="text"
              value={cellCoordinate}
              onChange={(e) => setCellCoordinate(e.target.value.toUpperCase())}
              placeholder="e.g. A1"
            />
          </div>

          <button className="btn-primary" onClick={handleFind}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
            </svg>
            Find Dependents
          </button>
        </div>

        {dependents.length > 0 && (
          <div className="dependents-result">
            {dependents[0].error ? (
              <div className="result-message error">
                <InfoIcon />
                {dependents[0].error}
              </div>
            ) : dependents[0].message ? (
              <div className="result-message empty">
                <InfoIcon />
                {dependents[0].message}
              </div>
            ) : (
              <div className="dependents-table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th>Sheet</th>
                      <th>Cell</th>
                      <th>Formula</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dependents.map((dep, i) => (
                      <tr key={i}>
                        <td>{dep.sheetName}</td>
                        <td><code style={{ fontSize: "0.85em" }}>{dep.coordinates}</code></td>
                        <td><code style={{ fontSize: "0.82em", wordBreak: "break-all" }}>{dep.formula}</code></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
});

export default CellDependentsPanel;

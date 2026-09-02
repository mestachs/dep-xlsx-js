import { useState, useEffect, useRef, useImperativeHandle, forwardRef } from "react";
import * as XLSX from "xlsx";
import SheetPreviewGrid from "./SheetPreviewGrid.jsx";
import PanelControls from "./PanelControls.jsx";

const InfoIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
  </svg>
);

const A1_REF_REGEX = /(?:(?:'([^']+)'|([a-zA-Z0-9_]+))!)?(\$?[A-Z]+)(\$?\d+)(?::(\$?[A-Z]+)(\$?\d+))?/g;
const NAME_REF_REGEX = /^(?:'([^']+)'|([^!]+))!(\$?[A-Z]+)(\$?\d+)(?::(\$?[A-Z]+)(\$?\d+))?$/;

const escapeRegExp = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// A cell's own formula, resolving shared-formula cells (cell.F) back to the
// top-left cell that actually holds the formula text (cell.f).
const formulaOf = (worksheet, cellAddress) => {
  const cell = worksheet?.[cellAddress];
  if (!cell) return null;
  if (cell.f) return cell.f;
  if (cell.F) {
    const range   = XLSX.utils.decode_range(cell.F);
    const topLeft = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
    return worksheet[topLeft]?.f ?? null;
  }
  return null;
};

// Parses a defined name's Ref (e.g. "'11_Normes'!$B$6" or "Sheet1!$A$1:$B$2")
// into a {sheet, start, end} shape, or null if it's not a plain sheet range
// (multi-area names, formula-valued names, etc. are left unsupported).
const parseNameRef = (ref) => {
  if (!ref) return null;
  const match = ref.replace(/^=/, "").match(NAME_REF_REGEX);
  if (!match) return null;
  const [, quotedSheet, bareSheet, startCol, startRow, endCol, endRow] = match;
  return {
    sheet: quotedSheet || bareSheet,
    start: XLSX.utils.decode_cell(`${startCol}${startRow}`),
    end: endCol && endRow ? XLSX.utils.decode_cell(`${endCol}${endRow}`) : null,
  };
};

const cellWithinNameRef = (parsed, sheetName, cell) => {
  if (!parsed || parsed.sheet !== sheetName) return false;
  const end = parsed.end || parsed.start;
  return cell.c >= parsed.start.c && cell.c <= end.c &&
         cell.r >= parsed.start.r && cell.r <= end.r;
};

// Named ranges (workbook.Workbook.Names) whose range covers this cell —
// formulas can reference the cell by name instead of by A1 coordinates.
const namesCovering = (workbook, sheetName, cell) =>
  (workbook.Workbook?.Names || [])
    .filter((def) => cellWithinNameRef(parseNameRef(def.Ref), sheetName, cell))
    .map((def) => def.Name);

// Splits "11_Normes!B6" into {sheet: "11_Normes", coord: "B6"}; a bare "B6"
// falls back to whatever sheet is passed in (typically the dropdown value).
const parseTarget = (raw, fallbackSheet) => {
  const trimmed = raw.trim();
  const bangIndex = trimmed.lastIndexOf("!");
  if (bangIndex === -1) return { sheet: fallbackSheet, coord: trimmed };
  const sheet = trimmed.slice(0, bangIndex).trim().replace(/^'(.*)'$/, "$1");
  const coord = trimmed.slice(bangIndex + 1).trim();
  return { sheet, coord };
};

/**
 * Card that lets the user pick a sheet + cell and see, both ways:
 *   - Dependents:   every formula cell in the workbook that references it,
 *                   whether directly (A1 coordinates) or via a named range
 *                   defined over that cell (e.g. SETTINGS_CURRENT_YEAR)
 *   - Dependencies: every cell its own formula references, same two ways
 *
 * Imperative handle (via ref):
 *   ref.current.reset()   ← clears form and results, called when a new file loads
 */
const CellDependentsPanel = forwardRef(function CellDependentsPanel(
  { sheetNames, workbook, collapsed, maximized, onToggleCollapse, onToggleMaximize },
  ref
) {
  const [selectedSheet, setSelectedSheet]   = useState("");
  const [cellCoordinate, setCellCoordinate] = useState("");
  const [dependents, setDependents]         = useState([]);
  const [dependencies, setDependencies]     = useState([]);
  // Trail of {sheet, coord} previously visited by clicking through results,
  // so the dependency tree can be walked and backed out of.
  const [trail, setTrail]                   = useState([]);
  // Cell currently shown in the sheet preview grid — set from a result row
  // without disturbing the active Find (that's what the "Go" button is for).
  const [previewTarget, setPreviewTarget]   = useState(null);
  const rootRef                             = useRef(null);

  useImperativeHandle(ref, () => ({
    reset: () => {
      setSelectedSheet("");
      setCellCoordinate("");
      setDependents([]);
      setDependencies([]);
      setTrail([]);
      setPreviewTarget(null);
    },
    // Called from outside (e.g. clicking a cell reference in the markdown
    // summary) to run Find on a specific cell and bring this panel into view.
    findCell: (sheet, coord) => {
      setTrail([]);
      runFind(sheet, coord);
      rootRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    },
    // Called from outside (e.g. clicking a sheet node in the dependency
    // graph) to select that sheet and show its preview, without running a
    // Find (there's no specific cell yet — that's still what Find is for).
    selectSheet: (sheetName) => {
      const resolved = workbook?.SheetNames.find(
        (name) => name.toLowerCase() === sheetName.toLowerCase()
      );
      if (!resolved) return;
      setSelectedSheet(resolved);
      setCellCoordinate("");
      setDependents([]);
      setDependencies([]);
      setTrail([]);
      setPreviewTarget({ sheet: resolved, coord: "A1" });
    },
  }));

  const findDependents = (selectedSheet, targetCell, coord) => {
    const results = [];

    // Named ranges pointing at this cell — a formula can reference it by
    // name (e.g. SETTINGS_CURRENT_YEAR) instead of by A1 coordinates.
    const coveringNames = namesCovering(workbook, selectedSheet, targetCell);
    const nameRegexes = coveringNames.map((name) => new RegExp(`\\b${escapeRegExp(name)}\\b`));

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];

      for (const cellAddress in worksheet) {
        if (cellAddress.startsWith("!")) continue;

        const formula = formulaOf(worksheet, cellAddress);
        if (!formula) continue;

        A1_REF_REGEX.lastIndex = 0;
        let match;
        let hit = false;
        let via = null;

        while ((match = A1_REF_REGEX.exec(formula)) !== null) {
          const refSheet = match[1] || match[2] || sheetName;
          if (refSheet !== selectedSheet) continue;

          const [, , , startCol, startRow, endCol, endRow] = match;

          if (endCol && endRow) {
            const s = XLSX.utils.decode_cell(`${startCol}${startRow}`);
            const e = XLSX.utils.decode_cell(`${endCol}${endRow}`);
            hit = targetCell.c >= s.c && targetCell.c <= e.c &&
                  targetCell.r >= s.r && targetCell.r <= e.r;
          } else {
            const ref = XLSX.utils.decode_cell(`${startCol}${startRow}`);
            hit = ref.c === targetCell.c && ref.r === targetCell.r;
          }

          if (hit) break;
        }

        if (!hit) {
          const namedHitIndex = nameRegexes.findIndex((re) => re.test(formula));
          if (namedHitIndex !== -1) {
            hit = true;
            via = coveringNames[namedHitIndex];
          }
        }

        if (hit) results.push({ sheetName, coordinates: cellAddress, formula, via });
      }
    }

    return results.length > 0
      ? results
      : [{ message: `No dependents found for ${selectedSheet}!${coord}.` }];
  };

  const findDependencies = (selectedSheet, coord) => {
    const worksheet   = workbook.Sheets[selectedSheet];
    const ownFormula  = formulaOf(worksheet, coord);

    if (!ownFormula) {
      return [{ message: `${selectedSheet}!${coord} has no formula, so it has no dependencies.` }];
    }

    const results = [];
    A1_REF_REGEX.lastIndex = 0;
    let match;
    while ((match = A1_REF_REGEX.exec(ownFormula)) !== null) {
      const refSheet = match[1] || match[2] || selectedSheet;
      const [, , , startCol, startRow, endCol, endRow] = match;
      const isRange = Boolean(endCol && endRow);
      const coordinates = isRange
        ? `${startCol}${startRow}:${endCol}${endRow}`
        : `${startCol}${startRow}`;

      const formula = isRange ? null : formulaOf(workbook.Sheets[refSheet], coordinates.replace(/\$/g, ""));
      results.push({ sheetName: refSheet, coordinates, formula });
    }

    // Named ranges referenced by identifier (e.g. SETTINGS_CURRENT_YEAR)
    // rather than by A1 coordinates.
    for (const def of workbook.Workbook?.Names || []) {
      if (!new RegExp(`\\b${escapeRegExp(def.Name)}\\b`).test(ownFormula)) continue;
      const parsed = parseNameRef(def.Ref);
      if (!parsed) continue;

      const coordinates = parsed.end
        ? `${XLSX.utils.encode_cell(parsed.start)}:${XLSX.utils.encode_cell(parsed.end)}`
        : XLSX.utils.encode_cell(parsed.start);
      const formula = parsed.end ? null : formulaOf(workbook.Sheets[parsed.sheet], coordinates);
      results.push({ sheetName: parsed.sheet, coordinates, formula, via: def.Name });
    }

    return results.length > 0
      ? results
      : [{ message: `${selectedSheet}!${coord}'s formula doesn't reference any other cells.` }];
  };

  // Computes results for {sheet, rawCoord} and makes it the current selection —
  // used by the Find button, by clicking a result row, and by loading a
  // bookmarked ?cell_ref= URL. Always keeps the URL's cell_ref in sync.
  const runFind = (sheet, rawCoord) => {
    if (!workbook) return;

    if (!sheet || !rawCoord) {
      const errorResult = [{ error: "Please select a sheet and enter a cell coordinate (or type Sheet!Cell)." }];
      setDependents(errorResult);
      setDependencies(errorResult);
      return;
    }

    const resolvedSheet = workbook.SheetNames.find(
      (name) => name.toLowerCase() === sheet.toLowerCase()
    );
    if (!resolvedSheet) {
      const errorResult = [{ error: `Unknown sheet "${sheet}".` }];
      setDependents(errorResult);
      setDependencies(errorResult);
      return;
    }
    sheet = resolvedSheet;

    const coord = rawCoord.toUpperCase();
    let targetCell;
    try {
      targetCell = XLSX.utils.decode_cell(coord);
    } catch {
      const errorResult = [{ error: "Invalid cell coordinate. Please use A1 format (e.g., A1, B2)." }];
      setDependents(errorResult);
      setDependencies(errorResult);
      return;
    }

    setSelectedSheet(sheet);
    setCellCoordinate(coord);
    setDependents(findDependents(sheet, targetCell, coord));
    setDependencies(findDependencies(sheet, coord));
    setPreviewTarget({ sheet, coord });

    const params = new URLSearchParams(window.location.search);
    params.set("cell_ref", `${sheet}!${coord}`);
    window.history.replaceState(null, "", `${window.location.pathname}?${params.toString()}`);
  };

  // Shows a result row's cell in the preview grid without changing the
  // active Find (that's what clicking "Go" is for).
  const previewCell = (sheet, rawCoord) => {
    const coord = rawCoord.split(":")[0].replace(/\$/g, "");
    setPreviewTarget({ sheet, coord });
  };

  const handleFind = () => {
    const { sheet, coord } = parseTarget(cellCoordinate, selectedSheet);
    setTrail([]);
    runFind(sheet, coord);
  };

  // Jumps to a cell clicked in a dependents/dependencies row, remembering
  // where we came from so it can be backed out of.
  const navigateTo = (sheet, rawCoord) => {
    const coord = rawCoord.split(":")[0].replace(/\$/g, "");
    if (selectedSheet && cellCoordinate) {
      setTrail((t) => [...t, `${selectedSheet}!${cellCoordinate}`]);
    }
    runFind(sheet, coord);
  };

  const handleBack = () => {
    if (trail.length === 0) return;
    const prev = trail[trail.length - 1];
    setTrail((t) => t.slice(0, -1));
    const { sheet, coord } = parseTarget(prev, selectedSheet);
    runFind(sheet, coord);
  };

  const jumpToTrail = (index) => {
    const target = trail[index];
    setTrail(trail.slice(0, index));
    const { sheet, coord } = parseTarget(target, selectedSheet);
    runFind(sheet, coord);
  };

  // User-facing "Clear" — unlike the imperative reset() (used when a new
  // workbook loads), this also drops ?cell_ref= from the URL so the panel
  // goes back to its compact, empty state and stays that way on reload.
  const handleClear = () => {
    setSelectedSheet("");
    setCellCoordinate("");
    setDependents([]);
    setDependencies([]);
    setTrail([]);
    setPreviewTarget(null);

    const params = new URLSearchParams(window.location.search);
    params.delete("cell_ref");
    const query = params.toString();
    window.history.replaceState(null, "", `${window.location.pathname}${query ? `?${query}` : ""}`);
  };

  useEffect(() => {
    if (!workbook) return;
    const params = new URLSearchParams(window.location.search);
    const cellRef = params.get("cell_ref");
    if (!cellRef) return;
    const { sheet, coord } = parseTarget(cellRef, "");
    setTrail([]);
    runFind(sheet, coord);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [workbook]);

  const renderResults = (results) => {
    if (results.length === 0) return null;
    if (results[0].error) {
      return (
        <div className="result-message error">
          <InfoIcon />
          {results[0].error}
        </div>
      );
    }
    if (results[0].message) {
      return (
        <div className="result-message empty">
          <InfoIcon />
          {results[0].message}
        </div>
      );
    }
    return (
      <div className="dependents-table-wrapper">
        <table>
          <thead>
            <tr>
              <th>Sheet</th>
              <th>Cell</th>
              <th>Via</th>
              <th>Formula</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {results.map((dep, i) => {
              const navigable = workbook?.SheetNames.includes(dep.sheetName);
              return (
                <tr key={i}>
                  <td>{dep.sheetName}</td>
                  <td>
                    {navigable ? (
                      <button
                        type="button"
                        className="cell-coord-link"
                        title={`Preview ${dep.sheetName}!${dep.coordinates}`}
                        onClick={() => previewCell(dep.sheetName, dep.coordinates)}
                      >
                        {dep.coordinates}
                      </button>
                    ) : (
                      <code style={{ fontSize: "0.85em" }}>{dep.coordinates}</code>
                    )}
                  </td>
                  <td>
                    {dep.via ? (
                      <code className="named-range-badge" title={`Named range: ${dep.via}`}>{dep.via}</code>
                    ) : (
                      <span style={{ color: "var(--color-text-secondary)" }}>direct</span>
                    )}
                  </td>
                  <td>
                    {dep.formula ? (
                      <code style={{ fontSize: "0.82em", wordBreak: "break-all" }}>{dep.formula}</code>
                    ) : (
                      <span style={{ color: "var(--color-text-secondary)" }}>—</span>
                    )}
                  </td>
                  <td>
                    {navigable && (
                      <button
                        type="button"
                        className="cell-nav-btn"
                        title={`Find dependents & dependencies of ${dep.sheetName}!${dep.coordinates}`}
                        onClick={() => navigateTo(dep.sheetName, dep.coordinates)}
                      >
                        Go
                        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                          <line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>
                        </svg>
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  const hasSelection = selectedSheet || cellCoordinate || dependents.length > 0 || dependencies.length > 0;

  return (
    <div className="card cell-analysis" ref={rootRef}>
      <div className="card-header">
        <h2>Find Cell Dependents &amp; Dependencies</h2>
        <PanelControls
          collapsed={collapsed}
          maximized={maximized}
          onToggleCollapse={onToggleCollapse}
          onToggleMaximize={onToggleMaximize}
        />
      </div>
      <div className={`card-body${collapsed ? " panel-collapsed" : ""}`}>
        <div className="cell-analysis-fields">
          <div className="field-group">
            <label htmlFor="sheet-select">Sheet</label>
            <select
              id="sheet-select"
              value={selectedSheet}
              onChange={(e) => {
                const sheet = e.target.value;
                setSelectedSheet(sheet);
                if (sheet) setPreviewTarget({ sheet, coord: "A1" });
              }}
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
              onChange={(e) => setCellCoordinate(e.target.value)}
              placeholder="e.g. A1 or 11_Normes!B6"
            />
          </div>

          <button className="btn-primary" onClick={handleFind}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
            </svg>
            Find
          </button>

          {trail.length > 0 && (
            <button type="button" className="btn-secondary" onClick={handleBack}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/>
              </svg>
              Back
            </button>
          )}

          {hasSelection && (
            <button type="button" className="btn-secondary" onClick={handleClear} title="Clear selection and results">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
              </svg>
              Clear
            </button>
          )}
        </div>

        {(trail.length > 0 || (selectedSheet && cellCoordinate)) && (
          <div className="cell-trail">
            {trail.map((entry, i) => (
              <span key={i} className="cell-trail-item">
                <button type="button" onClick={() => jumpToTrail(i)}>{entry}</button>
                <span className="cell-trail-sep">→</span>
              </span>
            ))}
            {selectedSheet && cellCoordinate && (
              <span className="cell-trail-current">{selectedSheet}!{cellCoordinate}</span>
            )}
          </div>
        )}

        {previewTarget && (
          <div className="sheet-preview">
            <h3 className="dependents-result-heading">
              Sheet Preview <span>— {previewTarget.sheet}!{previewTarget.coord}</span>
            </h3>
            {(() => {
              const previewFormula = formulaOf(workbook?.Sheets?.[previewTarget.sheet], previewTarget.coord);
              return previewFormula ? (
                <code className="sheet-preview-formula" style={{ display: "block", fontSize: "0.82em", wordBreak: "break-all", marginBottom: "0.5em" }}>
                  ={previewFormula}
                </code>
              ) : null;
            })()}
            <SheetPreviewGrid
              workbook={workbook}
              sheet={previewTarget.sheet}
              cellAddress={previewTarget.coord}
              onCellClick={(coord) => navigateTo(previewTarget.sheet, coord)}
            />
          </div>
        )}

        {dependents.length > 0 && (
          <div className="dependents-result">
            <h3 className="dependents-result-heading">
              Dependents <span>— cells that reference this one</span>
            </h3>
            {renderResults(dependents)}
          </div>
        )}

        {dependencies.length > 0 && (
          <div className="dependents-result">
            <h3 className="dependents-result-heading">
              Dependencies <span>— cells this one's formula references</span>
            </h3>
            {renderResults(dependencies)}
          </div>
        )}
      </div>
    </div>
  );
});

export default CellDependentsPanel;

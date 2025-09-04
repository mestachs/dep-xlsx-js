import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import mermaid from "mermaid";
import { buildSheetDependencies } from "./dependencyBuilder.js";
import { marked } from "marked";
import { buildMarkdownSummary } from "./markdownBuilder.js";

mermaid.initialize({
  theme: "base",
  fontFamily: "Virgil, sans-serif",
  themeVariables: {
    primaryColor: "#f4f4f4",
    primaryTextColor: "#333",
    lineColor: "#666",
    tertiaryColor: "#f4f4f4",
  },
});

function App() {
  const [summary, setSummary] = useState("");
  const [mermaidGraph, setMermaidGraph] = useState("");
  const [error, setError] = useState("");
  const [copyFeedback, setCopyFeedback] = useState("");
  const [sheetVisibility, setSheetVisibility] = useState({});
  const [zoomLevel, setZoomLevel] = useState(1);
  const [graphDirection, setGraphDirection] = useState("TD");

  const [selectedSheetForCellAnalysis, setSelectedSheetForCellAnalysis] = useState("");
  const [selectedCellCoordinate, setSelectedCellCoordinate] = useState("");
  const [cellDependents, setCellDependents] = useState([]);

  const sheetNamesRef = useRef(null);
  const sheetDependenciesRef = useRef(null);
  const workbookRef = useRef(null); // New ref for workbook

  const generateGraph = (
    sheetNames,
    sheetDependencies,
    currentGraphDirection,
    currentSheetVisibility
  ) => {
    const graphLines = [
      `graph ${currentGraphDirection}`,
      "classDef visible fill:#afa,stroke:#333,stroke-width:2px;",
      "classDef hidden fill:#fca,stroke:#333,stroke-width:2px;",
      "classDef veryHidden fill:#fcc,stroke:#333,stroke-width:2px;",
    ];
    sheetNames.forEach((sheetName) => {
      const className = currentSheetVisibility[sheetName];
      graphLines.push(
        `    ${sheetName.replace(/ /g, "_")}[${sheetName}]:::${className}`
      );
    });
    sheetNames.forEach((sheetName) => {
      sheetDependencies[sheetName].forEach((dep) => {
        graphLines.push(
          `    ${dep.replace(/ /g, "_")} --> ${sheetName.replace(/ /g, "_")}`
        );
      });
    });
    return graphLines.join("\n");
  };

  useEffect(() => {
    if (mermaidGraph) {
      const mermaidElement = document.querySelector(".mermaid");
      if (mermaidElement) {
        mermaidElement.innerHTML = mermaidGraph;
        mermaidElement.removeAttribute("data-processed");
        mermaid.run({ nodes: [mermaidElement] });
      }
    }
  }, [mermaidGraph]);

  useEffect(() => {
    if (sheetNamesRef.current && sheetDependenciesRef.current) {
      const graph = generateGraph(
        sheetNamesRef.current,
        sheetDependenciesRef.current,
        graphDirection,
        sheetVisibility
      );
      setMermaidGraph(graph);
    }
  }, [graphDirection, sheetVisibility]);

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) {
      return;
    }

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellFormula: true });
      workbookRef.current = workbook; // Store workbook
      const sheetNames = workbook.SheetNames;
      const { sheetDependencies, sheetFormulaDetails } = buildSheetDependencies(workbook); // Destructure here

      sheetNamesRef.current = sheetNames;
      sheetDependenciesRef.current = sheetDependencies;

      const initialVisibility = {};
      workbook.SheetNames.forEach((name, index) => {
        const sheetProps = workbook.Workbook.Sheets[index];
        let visibility = "visible";
        if (sheetProps && sheetProps.Hidden === 1) {
          visibility = "hidden";
        } else if (sheetProps && sheetProps.Hidden === 2) {
          visibility = "veryHidden";
        }
        initialVisibility[name] = visibility;
      });
      setSheetVisibility(initialVisibility);

      const graph = generateGraph(
        sheetNames,
        sheetDependencies,
        graphDirection,
        initialVisibility
      );
      setMermaidGraph(graph);

      const summaryText = buildMarkdownSummary(
        sheetNames,
        sheetDependencies,
        workbook,
        initialVisibility,
        graph,
        sheetFormulaDetails // Pass sheetFormulaDetails
      );
      setSummary(summaryText);
      setError("");
    } catch (err) {
      console.error(err);
      setError("Error parsing XLSX file.");
      setSummary("");
      setMermaidGraph("");
    }
  };

  const copyMarkdownToClipboard = () => {
    navigator.clipboard
      .writeText(summary)
      .then(() => {
        setCopyFeedback("Copied to clipboard!");
        setTimeout(() => setCopyFeedback(""), 2000);
      })
      .catch((err) => {
        setCopyFeedback("Failed to copy!");
        console.error("Failed to copy markdown: ", err);
        setTimeout(() => setCopyFeedback(""), 2000);
      });
  };

  const handleZoomIn = () => {
    setZoomLevel((prevZoom) => prevZoom * 1.1);
  };

  const handleZoomOut = () => {
    setZoomLevel((prevZoom) => prevZoom * 0.9);
  };

  const handleToggleDirection = () => {
    setGraphDirection((prevDirection) =>
      prevDirection === "TD" ? "LR" : "TD"
    );
  };

  const handleFindDependents = () => {
    const targetSheetName = selectedSheetForCellAnalysis;
    const targetCellCoordinate = selectedCellCoordinate.toUpperCase(); // Ensure uppercase
    const dependents = [];

    if (!targetSheetName || !targetCellCoordinate || !workbookRef.current) {
      setCellDependents([{ error: 'Please select a sheet and enter a cell coordinate.' }]);
      return;
    }

    // Parse the target cell coordinate
    let targetCell = null;
    try {
      targetCell = XLSX.utils.decode_cell(targetCellCoordinate);
    } catch (e) {
      setCellDependents([{ error: 'Invalid cell coordinate. Please use A1 format (e.g., A1, $B$2).' }]);
      return;
    }

    const workbook = workbookRef.current;
    const sheetNames = workbook.SheetNames;

    // Regex to find all A1-style references (cells or ranges)
    // Captures: (sheet_name_quoted)? (sheet_name_unquoted)? (col_ref) (row_ref) (end_col_ref)? (end_row_ref)?
    const a1RefRegex = /(?:(?:'([^']+)'|([a-zA-Z0-9_]+))!)?(\$?[A-Z]+)(\$?\d+)(?::(\$?[A-Z]+)(\$?\d+))?/g;

    sheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      for (const cellAddress in worksheet) {
        const cell = worksheet[cellAddress];

        // Skip metadata cells (like !ref, !margins, etc.)
        if (cellAddress.startsWith('!')) continue;

        let formulaToProcess = null;

        if (cell.f) { // Regular formula
            formulaToProcess = cell.f;
        } else if (cell.F) { // Array formula
            const range = XLSX.utils.decode_range(cell.F);
            const topLeftCellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
            const topLeftCell = worksheet[topLeftCellAddress];
            if (topLeftCell && topLeftCell.f) {
                formulaToProcess = topLeftCell.f;
            }
        }

        if (formulaToProcess) {
          let match;
          // Reset regex lastIndex for each new formula
          a1RefRegex.lastIndex = 0;
          while ((match = a1RefRegex.exec(formulaToProcess)) !== null) {
            const refSheet = match[1] || match[2] || sheetName; // Referenced sheet or current sheet if not specified

            // If the referenced sheet is not the target sheet, skip
            if (refSheet !== targetSheetName) continue;

            const startColRef = match[3];
            const startRowRef = match[4];
            const endColRef = match[5]; // For ranges
            const endRowRef = match[6]; // For ranges

            let isDependent = false;

            if (endColRef && endRowRef) { // It's a range reference (e.g., A1:B5)
              const rangeStart = XLSX.utils.decode_cell(`${startColRef}${startRowRef}`);
              const rangeEnd = XLSX.utils.decode_cell(`${endColRef}${endRowRef}`);
              const decodedRange = { s: rangeStart, e: rangeEnd };

              // Check if target cell is within this range
              if (targetCell.c >= decodedRange.s.c && targetCell.c <= decodedRange.e.c &&
                  targetCell.r >= decodedRange.s.r && targetCell.r <= decodedRange.e.r) {
                isDependent = true;
              }
            } else { // It's a single cell reference (e.g., A1)
              const referencedCell = XLSX.utils.decode_cell(`${startColRef}${startRowRef}`);
              if (referencedCell.c === targetCell.c && referencedCell.r === targetCell.r) {
                isDependent = true;
              }
            }

            if (isDependent) {
              dependents.push({
                sheetName: sheetName,
                coordinates: cellAddress,
                formula: formulaToProcess
              });
              break; // Found a dependency, move to next cell in worksheet
            }
          }
        }
      }
    });
    setCellDependents(dependents.length > 0 ? dependents : [{ message: 'No dependents found.' }]);
  };

  return (
    <div>
      <header>
        <h1>XLSX Dependency Visualizer</h1>
        <p>Upload an XLSX file to see the sheet dependencies.</p>
        <input type="file" onChange={handleFile} accept=".xlsx" />
        {error && <p style={{ color: "red" }}>{error}</p>}
      </header>
      <main>
        <div className="markdown-container">
          {summary && (
            <>
              <button onClick={copyMarkdownToClipboard}>
                Copy Markdown to Clipboard
              </button>
              {copyFeedback && (
                <span style={{ marginLeft: "10px", color: "green" }}>
                  {copyFeedback}
                </span>
              )}
              <div
                dangerouslySetInnerHTML={{ __html: marked.parse(summary) }}
              />
            </>
          )}
        </div>
        <div className="graph-container">
          {mermaidGraph && (
            <>
              <div>
                <button onClick={handleZoomIn}>Zoom In</button>
                <button onClick={handleZoomOut}>Zoom Out</button>
                <button onClick={handleToggleDirection}>
                  Toggle Direction ({graphDirection})
                </button>
              </div>
              <h2>Rendered Graph</h2>
              <div
                className="mermaid"
                style={{
                  transform: `scale(${zoomLevel})`,
                  transformOrigin: "top left",
                }}
              ></div>
            </>
          )}

          {/* New section for cell dependency analysis */}
          <h3>Find Cell Dependents</h3>
          <div>
            <label htmlFor="sheet-select">Sheet:</label>
            <select
              id="sheet-select"
              value={selectedSheetForCellAnalysis}
              onChange={(e) => setSelectedSheetForCellAnalysis(e.target.value)}
            >
              <option value="">Select a sheet</option>
              {sheetNamesRef.current && sheetNamesRef.current.map(name => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </div>
          <div>
            <label htmlFor="cell-input">Cell Coordinate:</label>
            <input
              id="cell-input"
              type="text"
              value={selectedCellCoordinate}
              onChange={(e) => setSelectedCellCoordinate(e.target.value.toUpperCase())}
              placeholder="e.g., A1 or $B$2"
            />
          </div>
          <button onClick={handleFindDependents}>Find Dependents</button>

          {cellDependents.length > 0 && (
            <div>
              <h4>Dependents:</h4>
              {cellDependents[0].error ? ( // Check if it's an error message
                <p style={{ color: "red" }}>{cellDependents[0].error}</p>
              ) : cellDependents[0].message ? ( // Check if it's a "No dependents found" message
                <p>{cellDependents[0].message}</p>
              ) : (
                <table>
                  <thead>
                    <tr>
                      <th>Sheet Name</th>
                      <th>Coordinates</th>
                      <th>Formula</th>
                    </tr>
                  </thead>
                  <tbody>
                    {cellDependents.map((dep, index) => (
                      <tr key={index}>
                        <td>{dep.sheetName}</td>
                        <td>{dep.coordinates}</td>
                        <td>{dep.formula}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}
        </div>
      </main>
    </div>
  );
}

export default App;

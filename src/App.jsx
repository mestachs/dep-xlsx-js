import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import mermaid from "mermaid";
import { buildSheetDependencies } from "./dependencyBuilder.js";
import { buildMermaidGraph } from "./lib/graphBuilder.js";
import { marked } from "marked";
import { buildMarkdownSummary } from "./markdownBuilder.js";
import PanZoom from "./components/PanZoom.jsx";
import CellDependentsPanel from "./components/CellDependentsPanel.jsx";

mermaid.initialize({
  theme: "base",
  themeVariables: {
    primaryColor: "#eff6ff",
    primaryTextColor: "#1e293b",
    lineColor: "#94a3b8",
    tertiaryColor: "#f8fafc",
    edgeLabelBackground: "#ffffff",
  },
});

function App() {
  const [summary, setSummary]           = useState("");
  const [mermaidGraph, setMermaidGraph] = useState("");
  const [error, setError]               = useState("");
  const [copyFeedback, setCopyFeedback] = useState("");
  const [sheetVisibility, setSheetVisibility] = useState({});
  const [graphDirection, setGraphDirection]   = useState("TD");
  const [fileName, setFileName]         = useState("");
  const [highlightedSheet, setHighlightedSheet] = useState(null);

  const sheetNamesRef        = useRef(null);
  const sheetDependenciesRef = useRef(null);
  const workbookRef          = useRef(null);
  const panZoomRef           = useRef(null);
  const cellAnalysisRef      = useRef(null);

  const generateGraph = (sheetNames, sheetDependencies, direction, visibility) =>
    buildMermaidGraph(sheetNames, sheetDependencies, direction, visibility);

  useEffect(() => {
    if (!mermaidGraph) return;
    const el = document.querySelector(".mermaid");
    if (!el) return;
    el.innerHTML = mermaidGraph;
    el.removeAttribute("data-processed");
    mermaid.run({ nodes: [el] }).then(() => {
      el.querySelectorAll(".node").forEach((node) => {
        const label = node.querySelector(".nodeLabel")?.textContent?.trim();
        if (!label) return;
        node.style.cursor = "pointer";
        node.addEventListener("click", () => setHighlightedSheet(label));
      });
    }).catch((err) => {
      console.error("Mermaid rendering failed:", err);
      setError("Failed to render dependency graph — try reloading the page.");
    });
  }, [mermaidGraph]);

  useEffect(() => {
    if (!highlightedSheet) return;
    const body = document.querySelector(".markdown-body");
    if (!body) return;
    body.querySelectorAll("h3.sheet-highlight").forEach((h) => h.classList.remove("sheet-highlight"));
    const heading = Array.from(body.querySelectorAll("h3")).find(
      (h) => h.textContent.trim() === highlightedSheet
    );
    if (heading) {
      heading.classList.add("sheet-highlight");
      heading.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }, [highlightedSheet]);

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
    if (!file) return;

    setFileName(file.name);

    try {
      const data = await file.arrayBuffer();
      // sheetStubs:true preserves cells whose cached value is empty (<v/>),
      // which xlsx.js otherwise drops — those cells often carry the only formula text.
      const workbook = XLSX.read(data, { type: "array", cellFormula: true, sheetStubs: true });
      workbookRef.current = workbook;
      const sheetNames = workbook.SheetNames;
      const { sheetDependencies, sheetFormulaDetails } = buildSheetDependencies(workbook);

      sheetNamesRef.current       = sheetNames;
      sheetDependenciesRef.current = sheetDependencies;

      const initialVisibility = {};
      workbook.SheetNames.forEach((name, index) => {
        const sheetProps = workbook.Workbook.Sheets[index];
        let visibility = "visible";
        if (sheetProps && sheetProps.Hidden === 1) visibility = "hidden";
        else if (sheetProps && sheetProps.Hidden === 2) visibility = "veryHidden";
        initialVisibility[name] = visibility;
      });
      setSheetVisibility(initialVisibility);

      const graph = generateGraph(sheetNames, sheetDependencies, graphDirection, initialVisibility);
      setMermaidGraph(graph);

      const summaryText = buildMarkdownSummary(
        sheetNames, sheetDependencies, workbook, initialVisibility, graph, sheetFormulaDetails
      );
      setSummary(summaryText);
      setError("");
      panZoomRef.current?.reset();
      cellAnalysisRef.current?.reset();
      setHighlightedSheet(null);
    } catch (err) {
      console.error(err);
      setError("Error parsing XLSX file. Please ensure it is a valid .xlsx file.");
      setSummary("");
      setMermaidGraph("");
    }
  };

  const copyMarkdownToClipboard = () => {
    navigator.clipboard
      .writeText(summary)
      .then(() => {
        setCopyFeedback("Copied!");
        setTimeout(() => setCopyFeedback(""), 2000);
      })
      .catch(() => {
        setCopyFeedback("Failed to copy");
        setTimeout(() => setCopyFeedback(""), 2000);
      });
  };

  const handleToggleDirection = () =>
    setGraphDirection((d) => (d === "TD" ? "LR" : "TD"));

  const hasContent = summary || mermaidGraph;

  return (
    <div>
      <header>
        <div className="header-inner">
          <div className="header-brand">
            <div className="header-icon">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="3" width="7" height="7"/>
                <rect x="14" y="3" width="7" height="7"/>
                <rect x="14" y="14" width="7" height="7"/>
                <rect x="3" y="14" width="7" height="7"/>
              </svg>
            </div>
            <div className="header-title-group">
              <h1>XLSX Dependency Visualizer</h1>
              <p>Analyze and visualize sheet dependencies in Excel workbooks</p>
            </div>
          </div>

          <div className="header-upload">
            {error && (
              <span className="error-badge">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                {error}
              </span>
            )}
            {fileName && !error && (
              <span className="upload-filename" title={fileName}>
                {fileName}
              </span>
            )}
            <label className="file-upload-label">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                <polyline points="17 8 12 3 7 8"/>
                <line x1="12" y1="3" x2="12" y2="15"/>
              </svg>
              Upload XLSX
              <input type="file" onChange={handleFile} accept=".xlsx" />
            </label>
          </div>
        </div>
      </header>

      <main>
        {!hasContent ? (
          <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <div className="card" style={{ maxWidth: 420, width: "100%" }}>
              <div className="empty-state">
                <svg className="empty-state-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                  <polyline points="14 2 14 8 20 8"/>
                  <line x1="16" y1="13" x2="8" y2="13"/>
                  <line x1="16" y1="17" x2="8" y2="17"/>
                  <polyline points="10 9 9 9 8 9"/>
                </svg>
                <div>
                  <p style={{ fontWeight: 600, color: "var(--color-text-primary)", marginBottom: 4 }}>No file loaded</p>
                  <p>Upload an Excel (.xlsx) file to visualize its sheet dependencies and analyze cell references.</p>
                </div>
                <label className="file-upload-label" style={{ marginTop: 4 }}>
                  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="17 8 12 3 7 8"/>
                    <line x1="12" y1="3" x2="12" y2="15"/>
                  </svg>
                  Choose XLSX File
                  <input type="file" onChange={handleFile} accept=".xlsx" />
                </label>
              </div>
            </div>
          </div>
        ) : (
          <>
            <div className="markdown-container">
              {summary && (
                <div className="card">
                  <div className="card-header">
                    <h2>Summary</h2>
                    <div className="markdown-actions">
                      {copyFeedback && <span className="copy-feedback">{copyFeedback}</span>}
                      <button onClick={copyMarkdownToClipboard}>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                          <rect x="9" y="9" width="13" height="13" rx="2" ry="2"/>
                          <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
                        </svg>
                        Copy Markdown
                      </button>
                    </div>
                  </div>
                  <div
                    className="markdown-body"
                    dangerouslySetInnerHTML={{ __html: marked.parse(summary) }}
                  />
                </div>
              )}
            </div>

            <div className="graph-container">
              {mermaidGraph && (
                <div className="card graph-card">
                  <div className="card-header">
                    <h2>Dependency Graph</h2>
                    <div className="graph-controls">
                      <button onClick={() => panZoomRef.current?.zoomOut()} title="Zoom out">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="8" y1="11" x2="14" y2="11"/>
                        </svg>
                      </button>
                      <button onClick={() => panZoomRef.current?.reset()} title="Reset view">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                          <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/>
                          <path d="M3 3v5h5"/>
                        </svg>
                      </button>
                      <button onClick={() => panZoomRef.current?.zoomIn()} title="Zoom in">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/>
                        </svg>
                      </button>
                      <button onClick={handleToggleDirection} title="Toggle layout direction">
                        {graphDirection === "TD" ? (
                          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>
                          </svg>
                        ) : (
                          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/>
                          </svg>
                        )}
                        {graphDirection === "TD" ? "Left→Right" : "Top→Down"}
                      </button>
                    </div>
                  </div>
                  <PanZoom ref={panZoomRef}>
                    <div className="mermaid" />
                  </PanZoom>
                </div>
              )}

              <CellDependentsPanel
                ref={cellAnalysisRef}
                sheetNames={sheetNamesRef.current}
                workbook={workbookRef.current}
              />
            </div>
          </>
        )}
      </main>
    </div>
  );
}

export default App;

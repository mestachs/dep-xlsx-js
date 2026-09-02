import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import mermaid from "mermaid";
import { buildSheetDependencies } from "./dependencyBuilder.js";
import { buildMermaidGraph } from "./lib/graphBuilder.js";
import { marked } from "marked";
import { buildMarkdownSummary } from "./markdownBuilder.js";
import PanZoom from "./components/PanZoom.jsx";
import CellDependentsPanel from "./components/CellDependentsPanel.jsx";
import PanelControls from "./components/PanelControls.jsx";

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

// Accepts any Google Sheets URL shape (edit, view, gid fragment, ...) and
// pulls out the spreadsheet id, e.g. .../spreadsheets/d/<id>/edit?gid=123
const extractGoogleSheetId = (url) => {
  const match = String(url).match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
};

// Content-Disposition: attachment; filename="foo.xlsx"; filename*=UTF-8''foo.xlsx
const extractFilenameFromContentDisposition = (header) => {
  if (!header) return null;
  const utf8Match = header.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match) return decodeURIComponent(utf8Match[1]);
  const plainMatch = header.match(/filename="?([^";]+)"?/i);
  return plainMatch ? plainMatch[1] : null;
};

function App() {
  const [summary, setSummary]           = useState("");
  const [mermaidGraph, setMermaidGraph] = useState("");
  const [error, setError]               = useState("");
  const [copyFeedback, setCopyFeedback] = useState("");
  const [sheetVisibility, setSheetVisibility] = useState({});
  const [graphDirection, setGraphDirection]   = useState("LR");
  const [fileName, setFileName]         = useState("");
  const [highlightedSheet, setHighlightedSheet] = useState(null);
  const [sheetUrlInput, setSheetUrlInput] = useState("");
  const [urlLoading, setUrlLoading]     = useState(false);
  // Per-panel collapse (hide body, keep header) and the single panel — if
  // any — currently maximized to fill the whole main area.
  const [collapsedPanels, setCollapsedPanels] = useState({ summary: false, graph: false, find: false });
  const [maximizedPanel, setMaximizedPanel]   = useState(null);

  const sheetNamesRef        = useRef(null);
  const sheetDependenciesRef = useRef(null);
  const workbookRef          = useRef(null);
  const panZoomRef           = useRef(null);
  const cellAnalysisRef      = useRef(null);

  const generateGraph = (sheetNames, sheetDependencies, direction, visibility) =>
    buildMermaidGraph(sheetNames, sheetDependencies, direction, visibility);

  const toggleCollapse = (key) =>
    setCollapsedPanels((prev) => ({ ...prev, [key]: !prev[key] }));

  const toggleMaximize = (key) =>
    setMaximizedPanel((prev) => (prev === key ? null : key));

  // Escape restores the maximized panel, as a fallback to its own button.
  useEffect(() => {
    if (!maximizedPanel) return;
    const onKeyDown = (e) => {
      if (e.key === "Escape") setMaximizedPanel(null);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [maximizedPanel]);

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
        node.addEventListener("click", () => {
          setHighlightedSheet(label);
          cellAnalysisRef.current?.selectSheet(label);
        });
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

  const loadWorkbookFromArrayBuffer = (data, name) => {
    setFileName(name);

    // sheetStubs:true preserves cells whose cached value is empty (<v/>),
    // which xlsx.js otherwise drops — those cells often carry the only formula text.
    const workbook = XLSX.read(data, { type: "array", cellFormula: true, sheetStubs: true });
    workbookRef.current = workbook;
    const sheetNames = workbook.SheetNames;
    const { sheetDependencies, sheetFormulaDetails } = buildSheetDependencies(workbook);

    sheetNamesRef.current       = sheetNames;
    sheetDependenciesRef.current = sheetDependencies;

    const initialVisibility = {};
    workbook.SheetNames.forEach((sheetName, index) => {
      const sheetProps = workbook.Workbook.Sheets[index];
      let visibility = "visible";
      if (sheetProps && sheetProps.Hidden === 1) visibility = "hidden";
      else if (sheetProps && sheetProps.Hidden === 2) visibility = "veryHidden";
      initialVisibility[sheetName] = visibility;
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
  };

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      loadWorkbookFromArrayBuffer(data, file.name);
    } catch (err) {
      console.error(err);
      setError("Error parsing XLSX file. Please ensure it is a valid .xlsx file.");
      setSummary("");
      setMermaidGraph("");
    }
  };

  const loadFromSheetUrl = async (rawUrl) => {
    const url = rawUrl.trim();
    if (!url) return;

    const sheetId = extractGoogleSheetId(url);
    if (!sheetId) {
      setError("That doesn't look like a Google Sheets URL (expected .../spreadsheets/d/<id>/...).");
      return;
    }

    setUrlLoading(true);
    setError("");
    try {
      const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
      const response = await fetch(exportUrl);
      if (!response.ok) {
        throw new Error(`Google Sheets responded with ${response.status}`);
      }
      const data = await response.arrayBuffer();
      const name =
        extractFilenameFromContentDisposition(response.headers.get("content-disposition")) ||
        `${sheetId}.xlsx`;
      loadWorkbookFromArrayBuffer(data, name);

      const params = new URLSearchParams(window.location.search);
      params.set("sheet_url", url);
      window.history.replaceState(null, "", `${window.location.pathname}?${params.toString()}`);
    } catch (err) {
      console.error(err);
      setError(
        "Couldn't load that Google Sheet. Make sure it's shared as \"Anyone with the link can view\", then try again."
      );
      setSummary("");
      setMermaidGraph("");
    } finally {
      setUrlLoading(false);
    }
  };

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const initialUrl = params.get("sheet_url");
    if (initialUrl) {
      setSheetUrlInput(initialUrl);
      loadFromSheetUrl(initialUrl);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleSheetUrlSubmit = (e) => {
    e.preventDefault();
    loadFromSheetUrl(sheetUrlInput);
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

  // Cell references rendered in the markdown summary (e.g. the "Formulas
  // referencing other sheets" tables) carry data-sheet-ref/data-cell-ref —
  // clicking one runs Find on that cell in the panel below. Sheet names
  // (e.g. the "Used by this sheet" / "Uses this sheet" lists) carry only
  // data-sheet-ref — clicking one just selects that sheet and previews it.
  const handleMarkdownClick = (e) => {
    const cellTarget = e.target.closest("[data-cell-ref]");
    if (cellTarget) {
      const sheet = cellTarget.getAttribute("data-sheet-ref");
      const coord = cellTarget.getAttribute("data-cell-ref");
      if (sheet && coord) cellAnalysisRef.current?.findCell(sheet, coord);
      return;
    }

    const sheetTarget = e.target.closest("[data-sheet-ref]");
    if (sheetTarget) {
      const sheet = sheetTarget.getAttribute("data-sheet-ref");
      if (sheet) cellAnalysisRef.current?.selectSheet(sheet);
    }
  };

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

            <form className="sheet-url-form" onSubmit={handleSheetUrlSubmit}>
              <input
                type="url"
                className="sheet-url-input"
                placeholder="Paste a Google Sheets link…"
                value={sheetUrlInput}
                onChange={(e) => setSheetUrlInput(e.target.value)}
                disabled={urlLoading}
              />
              <button type="submit" className="sheet-url-submit" disabled={urlLoading || !sheetUrlInput.trim()}>
                {urlLoading ? (<><span className="spinner spinner-sm" /> Loading…</>) : "Load"}
              </button>
            </form>
          </div>
        </div>
      </header>

      <main data-maximized={maximizedPanel || undefined}>
        {!hasContent ? (
          <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <div className="card" style={{ maxWidth: 420, width: "100%" }}>
              {urlLoading ? (
                <div className="empty-state">
                  <span className="spinner" />
                  <div>
                    <p style={{ fontWeight: 600, color: "var(--color-text-primary)", marginBottom: 4 }}>Loading spreadsheet…</p>
                    <p>Fetching the Google Sheet and building the dependency graph.</p>
                  </div>
                </div>
              ) : (
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
                    <p>Upload an Excel (.xlsx) file, or paste a Google Sheets link above, to visualize its sheet dependencies and analyze cell references.</p>
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
              )}
            </div>
          </div>
        ) : (
          <>
            <div className="markdown-container">
              {summary && (
                <div className="card">
                  <div className="card-header">
                    <h2>Summary</h2>
                    <div className="card-header-actions">
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
                      <PanelControls
                        collapsed={collapsedPanels.summary}
                        maximized={maximizedPanel === "summary"}
                        onToggleCollapse={() => toggleCollapse("summary")}
                        onToggleMaximize={() => toggleMaximize("summary")}
                      />
                    </div>
                  </div>
                  <div
                    className={`markdown-body${collapsedPanels.summary ? " panel-collapsed" : ""}`}
                    onClick={handleMarkdownClick}
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
                    <div className="card-header-actions">
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
                      <PanelControls
                        collapsed={collapsedPanels.graph}
                        maximized={maximizedPanel === "graph"}
                        onToggleCollapse={() => toggleCollapse("graph")}
                        onToggleMaximize={() => toggleMaximize("graph")}
                      />
                    </div>
                  </div>
                  <div className={`graph-body${collapsedPanels.graph ? " panel-collapsed" : ""}`}>
                    <PanZoom ref={panZoomRef}>
                      <div className="mermaid" />
                    </PanZoom>
                  </div>
                </div>
              )}

              <CellDependentsPanel
                ref={cellAnalysisRef}
                sheetNames={sheetNamesRef.current}
                workbook={workbookRef.current}
                collapsed={collapsedPanels.find}
                maximized={maximizedPanel === "find"}
                onToggleCollapse={() => toggleCollapse("find")}
                onToggleMaximize={() => toggleMaximize("find")}
              />
            </div>
          </>
        )}
      </main>
    </div>
  );
}

export default App;

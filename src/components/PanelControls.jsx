const ChevronIcon = ({ up }) => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    {up ? <polyline points="18 15 12 9 6 15" /> : <polyline points="6 9 12 15 18 9" />}
  </svg>
);

const MaximizeIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="15 3 21 3 21 9" /><polyline points="9 21 3 21 3 15" />
    <line x1="21" y1="3" x2="14" y2="10" /><line x1="3" y1="21" x2="10" y2="14" />
  </svg>
);

const RestoreIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="4 14 10 14 10 20" /><polyline points="20 10 14 10 14 4" />
    <line x1="14" y1="10" x2="21" y2="3" /><line x1="3" y1="21" x2="10" y2="14" />
  </svg>
);

/**
 * A pair of header buttons — collapse/expand (hide the body, keep the
 * header) and maximize/restore (this panel fills the whole main area,
 * the other two panels hidden) — shared by the Summary, Dependency Graph
 * and Find Cell & Dependencies cards. All state lives in the parent
 * (App's collapsedPanels/maximizedPanel); this component is just the UI.
 */
const PanelControls = ({ collapsed, maximized, onToggleCollapse, onToggleMaximize }) => (
  <div className="panel-controls">
    {!maximized && (
      <button
        type="button"
        className="panel-control-btn"
        onClick={onToggleCollapse}
        title={collapsed ? "Expand panel" : "Collapse panel"}
      >
        <ChevronIcon up={!collapsed} />
      </button>
    )}
    <button
      type="button"
      className="panel-control-btn"
      onClick={onToggleMaximize}
      title={maximized ? "Restore panel" : "Maximize panel"}
    >
      {maximized ? <RestoreIcon /> : <MaximizeIcon />}
    </button>
  </div>
);

export default PanelControls;

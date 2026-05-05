/**
 * Converts a sheet name into a valid Mermaid node identifier.
 * Mermaid node IDs must be alphanumeric + underscore only, and cannot start
 * with a digit. Everything else is replaced with `_`.
 *
 * @param {string} name
 * @returns {string}
 */
export function sanitizeId(name) {
  const id = name.replace(/[^a-zA-Z0-9]/g, "_");
  if (/^[0-9]/.test(id)) return "n_" + id;
  return id || "n_empty";
}

/**
 * Builds a Mermaid flowchart string from workbook sheet dependency data.
 *
 * @param {string[]} sheetNames
 * @param {Record<string, Set<string>>} sheetDependencies  dep → Set of sheets it depends on
 * @param {"TD"|"LR"} direction
 * @param {Record<string, "visible"|"hidden"|"veryHidden">} sheetVisibility
 * @returns {string}
 */
export function buildMermaidGraph(sheetNames, sheetDependencies, direction, sheetVisibility) {
  const lines = [
    `graph ${direction}`,
    "classDef visible fill:#dbeafe,stroke:#3b82f6,stroke-width:1.5px,color:#1e40af;",
    "classDef hidden fill:#fef3c7,stroke:#f59e0b,stroke-width:1.5px,color:#92400e;",
    "classDef veryHidden fill:#fee2e2,stroke:#ef4444,stroke-width:1.5px,color:#991b1b;",
  ];

  for (const name of sheetNames) {
    lines.push(`    ${sanitizeId(name)}["${name}"]:::${sheetVisibility[name]}`);
  }

  for (const name of sheetNames) {
    for (const dep of sheetDependencies[name]) {
      lines.push(`    ${sanitizeId(dep)} --> ${sanitizeId(name)}`);
    }
  }

  return lines.join("\n");
}

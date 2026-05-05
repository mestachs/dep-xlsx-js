/**
 * Escapes a string so it is safe to embed in a RegExp pattern.
 * @param {string} str
 * @returns {string}
 */
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Extracts all sheet names referenced in a formula string.
 *
 * Handles three reference types:
 *   - Quoted sheet:   'Sheet Name'!A1
 *   - Unquoted sheet: SheetName!A1
 *   - Named range:    a name whose Ref resolves to a known sheet
 *
 * @param {string} formula        - formula text without leading `=`
 * @param {string[]} sheetNames   - all sheet names in the workbook
 * @param {Map<string,string>} [namedRanges] - lowercase-name → sheet-name; omit to skip named-range lookup
 * @returns {Set<string>}
 */
export function extractSheetRefs(formula, sheetNames, namedRanges) {
  const sheets = new Set();

  // 1. 'Quoted Sheet Name'!ref  (handles spaces / special chars in sheet names)
  const quotedSheetRe = /'([^']+)'!/g;
  let m;
  while ((m = quotedSheetRe.exec(formula)) !== null) {
    if (sheetNames.includes(m[1])) sheets.add(m[1]);
  }

  // 2. UnquotedSheetName!ref
  const unquotedSheetRe = /([A-Za-z0-9_]+)!/g;
  while ((m = unquotedSheetRe.exec(formula)) !== null) {
    if (sheetNames.includes(m[1])) sheets.add(m[1]);
  }

  // 3. Named range identifiers — use word-boundary around escaped name
  if (namedRanges) {
    for (const [name, sheet] of namedRanges.entries()) {
      const re = new RegExp(`\\b${escapeRegex(name)}\\b`, 'gi');
      if (re.test(formula)) sheets.add(sheet);
    }
  }

  return sheets;
}

/**
 * Builds a Map<lowercaseName, sheetName> for all workbook-level named ranges
 * whose Ref can be resolved to at least one sheet.
 *
 * @param {object}   workbook
 * @param {string[]} sheetNames
 * @returns {Map<string, string>}
 */
export function buildNamedRangesMap(workbook, sheetNames) {
  const namedRanges = new Map();
  const names = workbook?.Workbook?.Names;
  if (!names) return namedRanges;

  for (const entry of names) {
    const key = entry.Name.toLowerCase();
    // Resolve Ref without named-range context to avoid recursion
    const sheets = extractSheetRefs(entry.Ref, sheetNames);
    if (sheets.size > 0) {
      namedRanges.set(key, sheets.values().next().value);
    }
  }

  return namedRanges;
}

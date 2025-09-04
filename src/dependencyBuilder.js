import * as XLSX from 'xlsx';

function findSheetNameInFormula(formula, sheetNames, namedRanges) {
    const sheets = new Set();

    // 1. Direct sheet references
    const quotedSheetRegex = /'([^']+)'!/g;
    let match;
    while ((match = quotedSheetRegex.exec(formula)) !== null) {
        if (sheetNames.includes(match[1])) {
            sheets.add(match[1]);
        }
    }

    const unquotedSheetRegex = /([a-zA-Z0-9_]+)!/g;
    while ((match = unquotedSheetRegex.exec(formula)) !== null) {
        if (sheetNames.includes(match[1])) {
            sheets.add(match[1]);
        }
    }

    // 2. Named range references
    if (namedRanges) {
        for (const [name, sheet] of namedRanges.entries()) {
            // Use a regex to find the named range as a whole word, case-insensitive
            const namedRangeRegex = new RegExp(`\\b${name}\\b`, 'gi'); // 'gi' for global and case-insensitive
            if (namedRangeRegex.test(formula)) {
                sheets.add(sheet);
            }
        }
    }

    return sheets;
}

export function buildSheetDependencies(workbook) {
    const sheetDependencies = {};
    const sheetFormulaDetails = {}; // New object to store formula details
    const sheetNames = workbook.SheetNames;

    // Create a map of named ranges to their sheets
    const namedRanges = new Map();
    if (workbook.Workbook && workbook.Workbook.Names) {
        workbook.Workbook.Names.forEach(namedRange => {
            // Convert named range name to lowercase for consistent lookup
            const nameKey = namedRange.Name.toLowerCase();
            // Find the sheet name from the named range's reference
            const referencedSheets = findSheetNameInFormula(namedRange.Ref, sheetNames); // Pass without namedRanges to avoid recursion
            if (referencedSheets.size > 0) {
                // For simplicity, we take the first sheet found in the reference.
                // A named range could in theory span multiple sheets, but that's rare.
                const sheetName = referencedSheets.values().next().value;
                namedRanges.set(nameKey, sheetName); // Store with lowercase key
            }
        });
    }

    sheetNames.forEach(sheetName => {
        sheetDependencies[sheetName] = new Set();
        sheetFormulaDetails[sheetName] = []; // Initialize array for each sheet
    });

    sheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const processedArrayFormulaRanges = new Set(); // To avoid processing the same array formula multiple times

        for (const cellAddress in worksheet) {
            const cell = worksheet[cellAddress];

            // Skip metadata cells (like !ref, !margins, etc.)
            if (cellAddress.startsWith('!')) continue;

            let formulaToProcess = null;

            if (cell.f) { // Regular formula
                formulaToProcess = cell.f;
            } else if (cell.F) { // Array formula
                // cell.F is the range of the array formula, e.g., "A1:B2"
                // The formula itself is only in the top-left cell of this range.
                // We need to ensure we process this array formula only once.
                if (!processedArrayFormulaRanges.has(cell.F)) {
                    processedArrayFormulaRanges.add(cell.F);

                    // Get the top-left cell of the array formula range
                    const range = XLSX.utils.decode_range(cell.F);
                    const topLeftCellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
                    const topLeftCell = worksheet[topLeftCellAddress];

                    if (topLeftCell && topLeftCell.f) {
                        formulaToProcess = topLeftCell.f;
                    }
                }
            }

            if (formulaToProcess) {
                const referencedSheets = findSheetNameInFormula(formulaToProcess, sheetNames, namedRanges);
                referencedSheets.forEach(referencedSheet => {
                    if (referencedSheet !== sheetName) {
                        sheetDependencies[sheetName].add(referencedSheet);
                    }
                });
                // Store formula details if it references other sheets
                if (referencedSheets.size > 0) {
                    sheetFormulaDetails[sheetName].push({
                        cellAddress: cellAddress,
                        formula: formulaToProcess,
                        referencedSheets: Array.from(referencedSheets) // Convert Set to Array for easier use
                    });
                }
            }
        }
    });

    return { sheetDependencies, sheetFormulaDetails }; // Return both
}

import * as XLSX from 'xlsx';
import { extractSheetRefs, buildNamedRangesMap } from './lib/formulaParser.js';

export function buildSheetDependencies(workbook) {
    const sheetNames = workbook.SheetNames;
    const namedRanges = buildNamedRangesMap(workbook, sheetNames);

    const sheetDependencies = {};
    const sheetFormulaDetails = {};

    sheetNames.forEach(sheetName => {
        sheetDependencies[sheetName] = new Set();
        sheetFormulaDetails[sheetName] = [];
    });

    sheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const processedArrayRanges = new Set();

        for (const cellAddress in worksheet) {
            if (cellAddress.startsWith('!')) continue;

            const cell = worksheet[cellAddress];
            let formula = null;

            if (cell.f) {
                // Regular formula, or top-left cell of an array formula
                formula = cell.f;
                // Mark the array range as processed so non-top-left cells are skipped
                if (cell.F) processedArrayRanges.add(cell.F);
            } else if (cell.F && !processedArrayRanges.has(cell.F)) {
                // Non-top-left cell of an array formula — find the formula on the top-left cell
                processedArrayRanges.add(cell.F);
                const range = XLSX.utils.decode_range(cell.F);
                const topLeft = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
                formula = worksheet[topLeft]?.f ?? null;
            }

            if (!formula) continue;

            const refs = extractSheetRefs(formula, sheetNames, namedRanges);
            refs.forEach(ref => {
                if (ref !== sheetName) sheetDependencies[sheetName].add(ref);
            });

            if (refs.size > 0) {
                sheetFormulaDetails[sheetName].push({
                    cellAddress,
                    formula,
                    referencedSheets: Array.from(refs),
                });
            }
        }
    });

    return { sheetDependencies, sheetFormulaDetails };
}

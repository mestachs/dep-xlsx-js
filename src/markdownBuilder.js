import * as XLSX from 'xlsx';

export function buildMarkdownSummary(sheetNames, sheetDependencies, workbook, initialVisibility, graph, sheetFormulaDetails) {
    const summaryLines = [];
    summaryLines.push('# XLSX Dependency Analysis');
    summaryLines.push('');
    summaryLines.push('## Sheets');
    summaryLines.push('');

    sheetNames.forEach(sheetName => {
        summaryLines.push(`### ${sheetName}`);
        const worksheet = workbook.Sheets[sheetName];
        if (worksheet && worksheet['!ref']) {
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const numRows = range.e.r + 1;
            const maxCol = XLSX.utils.encode_col(range.e.c);
            summaryLines.push(`- Rows: ${numRows}`);
            summaryLines.push(`- Max Column: ${maxCol}`);
            summaryLines.push(`- Visibility: ${initialVisibility[sheetName]}`);
        } else {
            summaryLines.push('- No data in this sheet.');
        }

        // "Used by this sheet" (dependees)
        const usedByThisSheet = Array.from(sheetDependencies[sheetName] || []);
        if (usedByThisSheet.length > 0) {
            summaryLines.push(`- Used by this sheet:`);
            usedByThisSheet.forEach(sheet => summaryLines.push(`  - ${sheet}`)); // Indented list item
        } else {
            summaryLines.push('- Used by this sheet: None');
        }

        // "Uses this sheet" (dependers)
        const usesThisSheet = sheetNames.filter(otherSheetName => {
            return sheetDependencies[otherSheetName] && sheetDependencies[otherSheetName].has(sheetName);
        });
        if (usesThisSheet.length > 0) {
            summaryLines.push(`- Uses this sheet:`);
            usesThisSheet.forEach(sheet => summaryLines.push(`  - ${sheet}`)); // Indented list item
        } else {
            summaryLines.push('- Uses this sheet: None');
        }

        // Collapsed section for formula details
        const formulasInSheet = sheetFormulaDetails[sheetName] || [];
        if (formulasInSheet.length > 0) {
            summaryLines.push('<details>');
            summaryLines.push(`<summary>Formulas referencing other sheets (${formulasInSheet.length} found)</summary>`);
            summaryLines.push(''); // Add a blank line for better rendering

            // Group formulas by referenced sheet
            const formulasByReferencedSheet = new Map();
            formulasInSheet.forEach(detail => {
                detail.referencedSheets.forEach(refSheet => {
                    if (!formulasByReferencedSheet.has(refSheet)) {
                        formulasByReferencedSheet.set(refSheet, []);
                    }
                    formulasByReferencedSheet.get(refSheet).push(detail);
                });
            });

            // Sort referenced sheets for consistent output
            const sortedReferencedSheets = Array.from(formulasByReferencedSheet.keys()).sort();

            sortedReferencedSheets.forEach(refSheet => {
                summaryLines.push(`#### Referencing: ${refSheet}`);
                summaryLines.push('| Cell | Formula |');
                summaryLines.push('|---|---|');
                formulasByReferencedSheet.get(refSheet).slice(0, 10).forEach(detail => {
                    summaryLines.push(`| \`${detail.cellAddress}\` | \`${detail.formula}\` |`);
                });
                // Add a message if there are more than 10 formulas for this specific referenced sheet
                if (formulasByReferencedSheet.get(refSheet).length > 10) {
                    summaryLines.push(`- ... ${formulasByReferencedSheet.get(refSheet).length - 10} more formulas not shown for ${refSheet}.`);
                }
                summaryLines.push(''); // Blank line after each sub-section
            });

            summaryLines.push('</details>');
        } else {
            summaryLines.push('No formulas in this sheet reference other sheets.');
        }

        summaryLines.push('');
    });

    summaryLines.push('## Dependency Graph');
    summaryLines.push('');
    summaryLines.push('```mermaid');
    summaryLines.push(graph);
    summaryLines.push('```');
    return summaryLines.join('\n');
}
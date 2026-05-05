import { describe, test, expect } from 'bun:test';
import { extractSheetRefs, buildNamedRangesMap } from './formulaParser.js';
import { buildSheetDependencies } from '../dependencyBuilder.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Minimal mock workbook used by buildSheetDependencies tests */
function makeWorkbook({ sheet1Cells = {}, sheet2Cells = {}, names = [] } = {}) {
  return {
    SheetNames: ['Sheet1', 'Sheet2'],
    Sheets: {
      Sheet1: { '!ref': 'A1:Z100', ...sheet1Cells },
      Sheet2: { '!ref': 'A1:Z100', ...sheet2Cells },
    },
    Workbook: {
      Sheets: [{ Hidden: 0 }, { Hidden: 0 }],
      Names: names,
    },
  };
}

// ---------------------------------------------------------------------------
// extractSheetRefs
// ---------------------------------------------------------------------------

describe('extractSheetRefs', () => {
  const sheets = ['Sheet1', 'Sheet2', 'Sales Data'];

  test('quoted sheet reference', () => {
    const refs = extractSheetRefs("'Sheet2'!A1", sheets);
    expect(refs).toEqual(new Set(['Sheet2']));
  });

  test('unquoted sheet reference', () => {
    const refs = extractSheetRefs('Sheet2!A1+Sheet2!B1', sheets);
    expect(refs).toEqual(new Set(['Sheet2']));
  });

  test('sheet name with spaces (quoted)', () => {
    const refs = extractSheetRefs("'Sales Data'!A1:B10", sheets);
    expect(refs).toEqual(new Set(['Sales Data']));
  });

  test('named range maps to correct sheet', () => {
    const namedRanges = new Map([['salesdata', 'Sheet2']]);
    const refs = extractSheetRefs('SUM(SalesData)', sheets, namedRanges);
    expect(refs).toEqual(new Set(['Sheet2']));
  });

  test('named range lookup is case-insensitive', () => {
    const namedRanges = new Map([['salesdata', 'Sheet2']]);
    const refs = extractSheetRefs('SUM(SALESDATA)', sheets, namedRanges);
    expect(refs).toEqual(new Set(['Sheet2']));
  });

  test('named range with dots in name is matched exactly', () => {
    // Without escaping, "Sales.Data" regex would also match "SalesXData"
    const namedRanges = new Map([['sales.data', 'Sheet2']]);
    const refsMatch = extractSheetRefs('SUM(Sales.Data)', sheets, namedRanges);
    expect(refsMatch).toEqual(new Set(['Sheet2']));

    // "SalesXData" must NOT be matched by the "sales.data" pattern
    const refsMiss = extractSheetRefs('SUM(SalesXData)', sheets, namedRanges);
    expect(refsMiss).toEqual(new Set());
  });

  test('named range is not matched as a substring of a longer identifier', () => {
    const namedRanges = new Map([['data', 'Sheet2']]);
    const refs = extractSheetRefs('SalesData + OtherData', sheets, namedRanges);
    // "data" appears only inside longer words — \b must prevent a match
    expect(refs).toEqual(new Set());
  });

  test('returns empty set when formula has no recognisable reference', () => {
    const refs = extractSheetRefs('A1+B2*C3', sheets);
    expect(refs).toEqual(new Set());
  });

  test('multiple different sheet references in one formula', () => {
    const refs = extractSheetRefs("'Sales Data'!A1 + Sheet2!B2", sheets);
    expect(refs).toEqual(new Set(['Sales Data', 'Sheet2']));
  });

  test('skips tokens that match the regex pattern but are not in sheetNames', () => {
    const refs = extractSheetRefs('UNKNOWN!A1 + Sheet2!B1', sheets);
    expect(refs).toEqual(new Set(['Sheet2']));
  });
});

// ---------------------------------------------------------------------------
// buildNamedRangesMap
// ---------------------------------------------------------------------------

describe('buildNamedRangesMap', () => {
  test('returns empty map when workbook has no Names', () => {
    const wb = { Workbook: { Names: [] }, SheetNames: ['Sheet1'] };
    expect(buildNamedRangesMap(wb, ['Sheet1'])).toEqual(new Map());
  });

  test('returns empty map when Workbook property is missing', () => {
    const wb = { SheetNames: ['Sheet1'] };
    expect(buildNamedRangesMap(wb, ['Sheet1'])).toEqual(new Map());
  });

  test('maps named range to sheet resolved from Ref', () => {
    const wb = {
      SheetNames: ['Sheet1', 'Sheet2'],
      Workbook: {
        Names: [{ Name: 'SalesData', Ref: 'Sheet2!$A$1:$A$100' }],
      },
    };
    const map = buildNamedRangesMap(wb, ['Sheet1', 'Sheet2']);
    expect(map.get('salesdata')).toBe('Sheet2');
  });

  test('stores key in lower-case', () => {
    const wb = {
      SheetNames: ['Sheet1', 'Sheet2'],
      Workbook: {
        Names: [{ Name: 'MyRange', Ref: 'Sheet2!A1' }],
      },
    };
    const map = buildNamedRangesMap(wb, ['Sheet1', 'Sheet2']);
    expect(map.has('myrange')).toBe(true);
    expect(map.has('MyRange')).toBe(false);
  });

  test('ignores named range whose Ref cannot be resolved to a sheet', () => {
    const wb = {
      SheetNames: ['Sheet1'],
      Workbook: {
        Names: [{ Name: 'LocalRef', Ref: 'A1:B10' }], // no sheet qualifier
      },
    };
    const map = buildNamedRangesMap(wb, ['Sheet1']);
    expect(map.size).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// buildSheetDependencies — regular formula cases
// ---------------------------------------------------------------------------

describe('buildSheetDependencies – regular formulas', () => {
  test('detects direct cross-sheet reference', () => {
    const wb = makeWorkbook({
      sheet1Cells: { A1: { f: 'Sheet2!B1*2', v: 4, t: 'n' } },
    });
    const { sheetDependencies } = buildSheetDependencies(wb);
    expect([...sheetDependencies['Sheet1']]).toContain('Sheet2');
    expect([...sheetDependencies['Sheet2']]).toHaveLength(0);
  });

  test('detects named range reference in regular formula', () => {
    const wb = makeWorkbook({
      sheet1Cells: { A1: { f: 'SUM(SalesData)', v: 0, t: 'n' } },
      names: [{ Name: 'SalesData', Ref: 'Sheet2!$A$1:$A$10' }],
    });
    const { sheetDependencies } = buildSheetDependencies(wb);
    expect([...sheetDependencies['Sheet1']]).toContain('Sheet2');
  });

  test('no dependency when formula only references own sheet', () => {
    const wb = makeWorkbook({
      sheet1Cells: { A2: { f: 'A1*2', v: 2, t: 'n' } },
    });
    const { sheetDependencies } = buildSheetDependencies(wb);
    expect([...sheetDependencies['Sheet1']]).toHaveLength(0);
  });

  // xlsx.js drops cells whose cached value is empty (<v/>) unless sheetStubs:true is set.
  // Those stub cells (t:"z") often carry the only formula text in the workbook.
  test('detects dependency in stub cell (t:"z") — the sheetStubs:true scenario', () => {
    // Mimics what xlsx produces for a formula cell with no precomputed value
    const wb = makeWorkbook({
      sheet1Cells: {
        A1: { t: 'z', f: "TRANSPOSE('Sheet2'!A1:A5)", F: 'A1:A5', v: 0 },
        A2: { F: 'A1:A5', v: 0, t: 'n' },
        A3: { F: 'A1:A5', v: 0, t: 'n' },
      },
    });
    const { sheetDependencies } = buildSheetDependencies(wb);
    expect([...sheetDependencies['Sheet1']]).toContain('Sheet2');
  });

  test('stub cell with named range reference is resolved correctly', () => {
    const wb = makeWorkbook({
      sheet1Cells: {
        // top-left cell is a stub (no cached value) carrying the array formula
        A1: { t: 'z', f: 'SUM(SalesData)', F: 'A1:A3', v: 0 },
        A2: { F: 'A1:A3', v: 0, t: 'n' },
        A3: { F: 'A1:A3', v: 0, t: 'n' },
      },
      names: [{ Name: 'SalesData', Ref: 'Sheet2!$A$1:$A$5' }],
    });
    const { sheetDependencies } = buildSheetDependencies(wb);
    expect([...sheetDependencies['Sheet1']]).toContain('Sheet2');
  });
});

// ---------------------------------------------------------------------------
// buildSheetDependencies — array formula cases
// ---------------------------------------------------------------------------

describe('buildSheetDependencies – array formulas', () => {
  test('detects named range dependency on top-left cell of a single-cell array formula', () => {
    // {=SUM(SalesData)} in a single cell — cell has both f and F
    const wb = makeWorkbook({
      sheet1Cells: {
        A1: { f: 'SUM(SalesData)', F: 'A1', v: 15, t: 'n' },
      },
      names: [{ Name: 'SalesData', Ref: 'Sheet2!$A$1:$A$5' }],
    });
    const { sheetDependencies } = buildSheetDependencies(wb);
    expect([...sheetDependencies['Sheet1']]).toContain('Sheet2');
  });

  test('detects named range dependency via non-top-left cells of a multi-cell array formula', () => {
    // {=SalesData*2} spilled into A1:A5; only A1 carries the formula (f+F), A2-A5 carry only F
    const wb = makeWorkbook({
      sheet1Cells: {
        A1: { f: 'SalesData*2', F: 'A1:A5', v: 2, t: 'n' },
        A2: { F: 'A1:A5', v: 4, t: 'n' },
        A3: { F: 'A1:A5', v: 6, t: 'n' },
        A4: { F: 'A1:A5', v: 8, t: 'n' },
        A5: { F: 'A1:A5', v: 10, t: 'n' },
      },
      names: [{ Name: 'SalesData', Ref: 'Sheet2!$A$1:$A$5' }],
    });
    const { sheetDependencies } = buildSheetDependencies(wb);
    expect([...sheetDependencies['Sheet1']]).toContain('Sheet2');
  });

  test('array formula dependency is recorded only once even with multiple spill cells', () => {
    const wb = makeWorkbook({
      sheet1Cells: {
        A1: { f: 'SalesData*2', F: 'A1:A5', v: 2, t: 'n' },
        A2: { F: 'A1:A5', v: 4, t: 'n' },
        A3: { F: 'A1:A5', v: 6, t: 'n' },
        A4: { F: 'A1:A5', v: 8, t: 'n' },
        A5: { F: 'A1:A5', v: 10, t: 'n' },
      },
      names: [{ Name: 'SalesData', Ref: 'Sheet2!$A$1:$A$5' }],
    });
    const { sheetFormulaDetails } = buildSheetDependencies(wb);
    // The formula should appear exactly once in the details
    expect(sheetFormulaDetails['Sheet1']).toHaveLength(1);
  });

  test('detects direct sheet ref in multi-cell array formula via non-top-left cell', () => {
    // If iteration happens to encounter a non-top-left cell before the top-left cell
    // the dependency should still be detected correctly
    const wb = makeWorkbook({
      sheet1Cells: {
        B1: { f: 'SUM(Sheet2!A1:A5)', F: 'B1:B3', v: 15, t: 'n' },
        B2: { F: 'B1:B3', v: 15, t: 'n' },
        B3: { F: 'B1:B3', v: 15, t: 'n' },
      },
    });
    const { sheetDependencies } = buildSheetDependencies(wb);
    expect([...sheetDependencies['Sheet1']]).toContain('Sheet2');
  });

  test('two independent array formulas on the same sheet are both resolved', () => {
    const wb = makeWorkbook({
      sheet1Cells: {
        A1: { f: 'SUM(RangeA)', F: 'A1:A3', v: 0, t: 'n' },
        A2: { F: 'A1:A3', v: 0, t: 'n' },
        A3: { F: 'A1:A3', v: 0, t: 'n' },
        B1: { f: 'SUM(RangeB)', F: 'B1:B3', v: 0, t: 'n' },
        B2: { F: 'B1:B3', v: 0, t: 'n' },
        B3: { F: 'B1:B3', v: 0, t: 'n' },
      },
      names: [
        { Name: 'RangeA', Ref: 'Sheet2!$A$1:$A$3' },
        { Name: 'RangeB', Ref: 'Sheet2!$B$1:$B$3' },
      ],
    });
    const { sheetDependencies, sheetFormulaDetails } = buildSheetDependencies(wb);
    expect([...sheetDependencies['Sheet1']]).toContain('Sheet2');
    expect(sheetFormulaDetails['Sheet1']).toHaveLength(2);
  });
});

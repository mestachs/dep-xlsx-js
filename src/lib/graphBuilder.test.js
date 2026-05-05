import { describe, test, expect } from 'bun:test';
import { sanitizeId, buildMermaidGraph } from './graphBuilder.js';

// ---------------------------------------------------------------------------
// sanitizeId
// ---------------------------------------------------------------------------

describe('sanitizeId', () => {
  test('plain alphanumeric name is unchanged', () => {
    expect(sanitizeId('Sheet1')).toBe('Sheet1');
  });

  test('spaces become underscores', () => {
    expect(sanitizeId('Sheet Name')).toBe('Sheet_Name');
  });

  test('parentheses are replaced — Reference records(soft deleted)', () => {
    expect(sanitizeId('Reference records(soft deleted)')).toBe(
      'Reference_records_soft_deleted_'
    );
  });

  test('ampersand is replaced — Print - Goals & Impact', () => {
    expect(sanitizeId('Print - Goals & Impact')).toBe(
      'Print___Goals___Impact'
    );
  });

  test('leading/trailing spaces become underscores — " Disaggregation - Section E "', () => {
    const id = sanitizeId(' Disaggregation - Section E ');
    expect(id).toBe('_Disaggregation___Section_E_');
  });

  test('name starting with a digit gets n_ prefix', () => {
    expect(sanitizeId('2024 Report')).toBe('n_2024_Report');
  });

  test('empty string returns n_empty', () => {
    expect(sanitizeId('')).toBe('n_empty');
  });

  test('hyphens and special chars in section names', () => {
    expect(sanitizeId('Overview - Section A')).toBe('Overview___Section_A');
    expect(sanitizeId('WPTM - Section F')).toBe('WPTM___Section_F');
  });

  test('hash and dot are replaced', () => {
    expect(sanitizeId('Sheet#1.v2')).toBe('Sheet_1_v2');
  });
});

// ---------------------------------------------------------------------------
// buildMermaidGraph
// ---------------------------------------------------------------------------

describe('buildMermaidGraph', () => {
  const visibility = {
    Sheet1: 'visible',
    Sheet2: 'hidden',
  };

  test('produces correct graph header', () => {
    const graph = buildMermaidGraph(['Sheet1'], { Sheet1: new Set() }, 'TD', visibility);
    expect(graph).toMatch(/^graph TD/);
  });

  test('includes all three classDef lines', () => {
    const graph = buildMermaidGraph(['Sheet1'], { Sheet1: new Set() }, 'TD', visibility);
    expect(graph).toContain('classDef visible');
    expect(graph).toContain('classDef hidden');
    expect(graph).toContain('classDef veryHidden');
  });

  test('uses sanitized ID and quoted label for a plain name', () => {
    const graph = buildMermaidGraph(['Sheet1'], { Sheet1: new Set() }, 'TD', visibility);
    expect(graph).toContain('Sheet1["Sheet1"]:::visible');
  });

  test('uses sanitized ID but preserves original name in label — parens', () => {
    const vis = { 'Reference records(soft deleted)': 'veryHidden' };
    const graph = buildMermaidGraph(
      ['Reference records(soft deleted)'],
      { 'Reference records(soft deleted)': new Set() },
      'TD',
      vis
    );
    expect(graph).toContain(
      'Reference_records_soft_deleted_["Reference records(soft deleted)"]:::veryHidden'
    );
  });

  test('uses sanitized ID but preserves original name in label — ampersand', () => {
    const vis = { 'Print - Goals & Impact': 'visible' };
    const graph = buildMermaidGraph(
      ['Print - Goals & Impact'],
      { 'Print - Goals & Impact': new Set() },
      'TD',
      vis
    );
    expect(graph).toContain(
      'Print___Goals___Impact["Print - Goals & Impact"]:::visible'
    );
  });

  test('LR direction is reflected in header', () => {
    const graph = buildMermaidGraph(['Sheet1'], { Sheet1: new Set() }, 'LR', visibility);
    expect(graph).toMatch(/^graph LR/);
  });

  test('edge uses sanitized IDs for both ends', () => {
    const vis = {
      'Reference records(soft deleted)': 'veryHidden',
      'Overview - Section A': 'visible',
    };
    const deps = {
      'Reference records(soft deleted)': new Set(),
      'Overview - Section A': new Set(['Reference records(soft deleted)']),
    };
    const graph = buildMermaidGraph(
      ['Reference records(soft deleted)', 'Overview - Section A'],
      deps,
      'TD',
      vis
    );
    expect(graph).toContain(
      'Reference_records_soft_deleted_ --> Overview___Section_A'
    );
  });

  test('sheet with no dependencies produces no edge lines for it', () => {
    const vis = { Sheet1: 'visible', Sheet2: 'hidden' };
    const deps = { Sheet1: new Set(), Sheet2: new Set() };
    const graph = buildMermaidGraph(['Sheet1', 'Sheet2'], deps, 'TD', vis);
    expect(graph).not.toContain('-->');
  });

  test('full real-world node set from the reported failing graph', () => {
    const names = [
      'Overview - Section A',
      'Reference Records',
      'Reference records(soft deleted)',
      'Print - Goals & Impact',
      'Print - Objectives & Outcome',
      ' Disaggregation - Section E ',
    ];
    const vis = Object.fromEntries(names.map((n) => [n, 'visible']));
    const deps = Object.fromEntries(names.map((n) => [n, new Set()]));
    // Add one edge that exercises special chars on both sides
    deps['Overview - Section A'].add('Reference records(soft deleted)');

    const graph = buildMermaidGraph(names, deps, 'TD', vis);

    // Every node line must use a quoted label
    for (const name of names) {
      expect(graph).toContain(`["${name}"]`);
    }
    // The edge must use sanitized IDs only
    expect(graph).toContain(
      'Reference_records_soft_deleted_ --> Overview___Section_A'
    );
    // No raw parens or ampersands should appear in ID positions
    const lines = graph.split('\n');
    const edgeAndNodeLines = lines.filter((l) => l.includes(':::') || l.includes('-->'));
    for (const line of edgeAndNodeLines) {
      // Strip out the quoted label portion before checking for illegal chars
      const withoutLabel = line.replace(/"[^"]*"/, '');
      expect(withoutLabel).not.toMatch(/[()&]/);
    }
  });
});

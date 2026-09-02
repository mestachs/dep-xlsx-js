// Resolves a worksheet cell's fill (background) color from the style object
// SheetJS attaches to `cell.s` when the workbook is read with `cellStyles: true`.
//
// SheetJS's community build (the `xlsx` npm package) only ever populates the
// *fill* portion of a cell's style — font weight/italic/color, alignment and
// borders are parsed internally but never exposed per-cell — so this module
// deliberately only resolves background color, not full cell styling.

// Default Office theme colors, index 0–11 (dk1, lt1, dk2, lt2, accent1–6, hlink, folHlink).
// Used when a workbook doesn't expose its own theme.
export const OFFICE_THEME_DEFAULTS = [
  "000000", "FFFFFF", "44546A", "E7E6E6",
  "4472C4", "ED7D31", "A5A5A5", "FFC000",
  "5B9BD5", "70AD47", "0563C1", "954F72",
];

// Standard Excel indexed color palette (ECMA-376 §18.8.27), 64 entries.
const EXCEL_INDEXED_COLORS = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", // 0-7
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", // 8-15
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", // 16-23
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", // 24-31
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", // 32-39
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", // 40-47
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", // 48-55
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333", // 56-63
];

// Applies an OOXML luminance tint to a 6-char hex color (HLS space).
// tint > 0 lightens toward white; tint < 0 darkens toward black.
const applyTint = (hex, tint) => {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  const newL = tint >= 0 ? l + (1 - l) * tint : l + l * tint;
  const toHex = (x) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, "0").toUpperCase();
  if (s === 0) return toHex(newL) + toHex(newL) + toHex(newL);
  const q = newL < 0.5 ? newL * (1 + s) : newL + s - newL * s;
  const p = 2 * newL - q;
  const h2r = (pp, qq, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return pp + (qq - pp) * 6 * t;
    if (t < 1 / 2) return qq;
    if (t < 2 / 3) return pp + (qq - pp) * (2 / 3 - t) * 6;
    return pp;
  };
  return toHex(h2r(p, q, h + 1 / 3)) + toHex(h2r(p, q, h)) + toHex(h2r(p, q, h - 1 / 3));
};

// Resolves a SheetJS color object ({rgb} | {theme, tint} | {indexed}) to a
// 6-char uppercase hex string.
const resolveColor = (c, themeColors) => {
  if (!c) return undefined;
  if (c.rgb) return c.rgb.length === 8 ? c.rgb.slice(2) : c.rgb;
  if (c.theme !== undefined) {
    const base = themeColors[c.theme];
    if (!base) return undefined;
    return c.tint ? applyTint(base, c.tint) : base;
  }
  if (c.indexed !== undefined) {
    // 64 = system foreground (black), 65 = system background (white)
    if (c.indexed === 64) return "000000";
    if (c.indexed === 65) return "FFFFFF";
    return EXCEL_INDEXED_COLORS[c.indexed];
  }
  return undefined;
};

// Reads a workbook's own theme color palette; falls back to Office defaults
// when the workbook doesn't expose one (SheetJS attaches it at `wb.Themes`,
// an undocumented property).
export const extractThemeColors = (workbook) => {
  const scheme = workbook?.Themes?.themeElements?.clrScheme;
  if (!scheme) return OFFICE_THEME_DEFAULTS;
  const names = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  return names.map((name, i) => {
    const entry = scheme[name];
    const raw = entry?.srgbClr ?? entry?.rgb ?? entry?.lastClr;
    if (!raw) return OFFICE_THEME_DEFAULTS[i] ?? "000000";
    return raw.length === 8 ? raw.slice(2) : raw;
  });
};

// Resolves a cell's background fill color as a "#rrggbb" CSS color, or
// undefined when the cell has no meaningful fill.
//
// `cell.s` (when present) *is* the fill object directly — e.g.
// `{ patternType: "solid", bgColor: { rgb: "FF0000" } }` — not nested under
// a `.fill` property. For a solid pattern, Excel stores the effective color
// in `bgColor`, not `fgColor` (verified against real workbooks); `fgColor`
// is only meaningful for a two-color pattern fill (stripes, etc.).
export const getCellFillColor = (cell, themeColors = OFFICE_THEME_DEFAULTS) => {
  const s = cell?.s;
  if (!s || typeof s !== "object" || !s.patternType || s.patternType === "none") return undefined;

  const hex = s.patternType === "solid"
    ? resolveColor(s.bgColor, themeColors) ?? resolveColor(s.fgColor, themeColors)
    : resolveColor(s.fgColor, themeColors) ?? resolveColor(s.bgColor, themeColors);

  if (!hex) return undefined;
  // Pure white/black fills are almost always the workbook's default — skip
  // them so the preview isn't blanketed in "highlighted" cells.
  if (hex.toUpperCase() === "FFFFFF" || hex.toUpperCase() === "000000") return undefined;
  return `#${hex}`;
};

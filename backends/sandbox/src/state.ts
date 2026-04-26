// skills/yome-skill-xl/backends/sandbox/src/state.ts
//
// Virtual world state for the xl (Excel) domain. Drives sandbox replays /
// benchmarks in the Hub, and feeds the viewer (viewer/index.html) via
// postMessage.
//
// Per spec 3.6, this state lives in BenchmarkFixtures.skillStates.xl for a
// case's t=0 snapshot, and the sandbox handler mutates a working copy as
// the trace plays out.

export interface XlCell {
  value?: string | number;
  formula?: string;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  color?: string;
  bg?: string;
  align?: 'left' | 'center' | 'right';
  numFmt?: string;
  border?: 'all' | 'top' | 'bottom' | 'left' | 'right' | 'outline';
  /** True if this cell is the top-left of a merged range. */
  mergeAnchor?: string; // e.g. 'A1:B3'
  /** True if this cell is part of a merged range but not the anchor. */
  mergedInto?: string;  // anchor address
}

export interface XlSheet {
  /** 1-based sheet index */
  index: number;
  name: string;
  active?: boolean;
  /** Map keyed by A1 address, e.g. 'B3'. */
  cells: Record<string, XlCell>;
  /** Column widths keyed by column letter. */
  colWidths?: Record<string, number>;
  /** Row heights keyed by row number (string). */
  rowHeights?: Record<string, number>;
  /** Used range row/col counts (cached from last `used`/handler ops). */
  rows: number;
  cols: number;
}

export interface XlFile {
  path: string;        // e.g. '~/Desktop/Q3.xlsx'
  sheets: XlSheet[];
  active?: boolean;    // True if this file is the currently active workbook.
  dirty?: boolean;
}

export interface XlWorldState {
  openFiles: XlFile[];
}

export function createInitialXlState(seed?: Partial<XlWorldState>): XlWorldState {
  return {
    openFiles: seed?.openFiles ?? [],
  };
}

export function getActiveFile(state: XlWorldState): XlFile | undefined {
  return state.openFiles.find(f => f.active) ?? state.openFiles[0];
}

export function setActiveFile(state: XlWorldState, path: string): void {
  for (const f of state.openFiles) f.active = f.path === path;
}

export function getActiveSheet(file: XlFile): XlSheet | undefined {
  return file.sheets.find(s => s.active) ?? file.sheets[0];
}

export function findSheet(file: XlFile, name?: string): XlSheet | undefined {
  if (!name) return getActiveSheet(file);
  return file.sheets.find(s => s.name === name);
}

/** Default new-workbook seed: one empty Sheet1. */
export function newBlankFile(path: string): XlFile {
  return {
    path,
    dirty: true,
    active: true,
    sheets: [
      {
        index: 1,
        name: 'Sheet1',
        active: true,
        cells: {},
        rows: 0,
        cols: 0,
      },
    ],
  };
}

// ──────────────────────────────────────────────────────────────────────
// Address helpers (A1 ↔ {row, col}) — colocated with state because the
// sandbox handlers need them and we want zero runtime deps.
// ──────────────────────────────────────────────────────────────────────

export function colLetter(idx: number): string {
  // 0-based: 0 → A, 25 → Z, 26 → AA
  let n = idx;
  let out = '';
  while (n >= 0) {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  }
  return out;
}

export function colIndex(letter: string): number {
  // 1-based: A → 1, Z → 26, AA → 27. Returns 0 for invalid input.
  const s = letter.toUpperCase();
  let result = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 65 || c > 90) return 0;
    result = result * 26 + (c - 64);
  }
  return result;
}

export interface CellAddr { col: number; row: number; }

/** 'B3' → { col: 2, row: 3 }; returns null if malformed. */
export function parseAddr(addr: string): CellAddr | null {
  const m = addr.match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return null;
  const col = colIndex(m[1]);
  const row = parseInt(m[2], 10);
  if (col === 0 || !Number.isFinite(row) || row < 1) return null;
  return { col, row };
}

export function formatAddr(c: CellAddr): string {
  return `${colLetter(c.col - 1)}${c.row}`;
}

export interface RangeBox { start: CellAddr; end: CellAddr; }

/** 'A1:C10' → { start, end }. Single-cell 'B3' → degenerate range. */
export function parseRange(spec: string): RangeBox | null {
  if (!spec.includes(':')) {
    const a = parseAddr(spec);
    return a ? { start: a, end: a } : null;
  }
  const [s, e] = spec.split(':');
  const a = parseAddr(s);
  const b = parseAddr(e);
  if (!a || !b) return null;
  return {
    start: { col: Math.min(a.col, b.col), row: Math.min(a.row, b.row) },
    end:   { col: Math.max(a.col, b.col), row: Math.max(a.row, b.row) },
  };
}

export interface ParsedRef {
  sheet?: string;
  /** raw address part after the optional Sheet@ prefix */
  address: string;
}

/** 'Sheet2@A1:B3' → { sheet: 'Sheet2', address: 'A1:B3' }. */
export function parseRef(raw: string): ParsedRef {
  const at = raw.indexOf('@');
  if (at >= 0) return { sheet: raw.slice(0, at), address: raw.slice(at + 1) };
  return { address: raw };
}

/** Recompute sheet.rows / sheet.cols from the cells map. */
export function recomputeUsed(sheet: XlSheet): void {
  let maxRow = 0;
  let maxCol = 0;
  for (const addr of Object.keys(sheet.cells)) {
    const a = parseAddr(addr);
    if (!a) continue;
    if (a.row > maxRow) maxRow = a.row;
    if (a.col > maxCol) maxCol = a.col;
  }
  sheet.rows = maxRow;
  sheet.cols = maxCol;
}

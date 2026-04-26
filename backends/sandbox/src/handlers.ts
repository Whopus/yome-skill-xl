// skills/yome-skill-xl/backends/sandbox/src/handlers.ts
//
// Sandbox backend for the xl (Excel) domain. Implements the xl signature
// against an in-memory XlWorldState. Real macOS backend lives in
// ../../macos/ as a declarative manifest + AppleScript templates;
// real Server runtime lives in Server/agent/commands/xlCommands.ts.
//
// Contract (matches existing handlers in web/lib/sandbox/handlers/*):
//   handleXl(world, cmd) -> { stdout, stderr, exitCode }

import {
  colIndex,
  colLetter,
  createInitialXlState,
  findSheet,
  formatAddr,
  getActiveFile,
  getActiveSheet,
  newBlankFile,
  parseAddr,
  parseRange,
  parseRef,
  recomputeUsed,
  setActiveFile,
  type XlCell,
  type XlFile,
  type XlSheet,
  type XlWorldState,
} from './state';

export interface ParsedCommand {
  domain: string;
  action: string;
  positionals: string[];
  flags: Record<string, string>;
}

export interface HandlerResult {
  stdout: string;
  stderr: string;
  exitCode: number;
}

const ok  = (stdout: string): HandlerResult => ({ stdout, stderr: '', exitCode: 0 });
const err = (stderr: string, code = 1): HandlerResult => ({ stdout: '', stderr, exitCode: code });

interface XlCarrier {
  scratch: { skillStates?: { xl?: XlWorldState } & Record<string, unknown> } & Record<string, unknown>;
}

function getState(world: XlCarrier): XlWorldState {
  if (!world.scratch.skillStates) world.scratch.skillStates = {};
  if (!world.scratch.skillStates.xl) {
    world.scratch.skillStates.xl = createInitialXlState();
  }
  return world.scratch.skillStates.xl!;
}

function findFile(state: XlWorldState, path: string): XlFile | undefined {
  return state.openFiles.find(f => f.path === path);
}

function parseIntOrNull(s: string | undefined): number | null {
  if (s === undefined || s === '') return null;
  const n = Number.parseInt(s, 10);
  return Number.isFinite(n) ? n : null;
}

function ensureCell(sheet: XlSheet, addr: string): XlCell {
  let c = sheet.cells[addr];
  if (!c) {
    c = {};
    sheet.cells[addr] = c;
  }
  return c;
}

function resolveSheet(file: XlFile, refSheet: string | undefined): XlSheet | undefined {
  if (!refSheet) return getActiveSheet(file);
  return findSheet(file, refSheet);
}

export function handleXl(world: XlCarrier, cmd: ParsedCommand): HandlerResult {
  const state = getState(world);

  switch (cmd.action) {
    // ─── Workbook ──────────────────────────────────────────────────
    case 'open': {
      const path = cmd.positionals[0] ?? cmd.flags.path;
      if (!path) return err('xl open: missing <path>');
      let file = findFile(state, path);
      if (!file) {
        file = { path, sheets: [{ index: 1, name: 'Sheet1', active: true, cells: {}, rows: 0, cols: 0 }], active: true };
        state.openFiles.push(file);
      }
      setActiveFile(state, path);
      return ok(JSON.stringify({ ok: true, opened: path, sheets: file.sheets.length }));
    }

    case 'new': {
      const path = cmd.positionals[0] ?? cmd.flags.path ?? `~/Desktop/Untitled-${state.openFiles.length + 1}.xlsx`;
      if (findFile(state, path) && cmd.flags.force !== 'true') {
        return err(`xl new: ${path} already open (use --force)`);
      }
      const file = newBlankFile(path);
      state.openFiles = state.openFiles.filter(f => f.path !== path);
      state.openFiles.push(file);
      setActiveFile(state, path);
      return ok(JSON.stringify({ ok: true, created: path, sheets: 1 }));
    }

    case 'save': {
      const file = getActiveFile(state);
      if (!file) return err('xl save: no active workbook');
      const newPath = cmd.flags.path;
      if (newPath) {
        if (state.openFiles.some(f => f.path === newPath) && cmd.flags.force !== 'true') {
          return err(`xl save: ${newPath} already exists (use --force)`);
        }
        file.path = newPath;
      }
      file.dirty = false;
      return ok(JSON.stringify({ ok: true, path: file.path }));
    }

    case 'close': {
      const file = getActiveFile(state);
      if (!file) return err('xl close: no active workbook');
      const shouldSave = cmd.flags.save !== 'false';
      if (shouldSave) file.dirty = false;
      state.openFiles = state.openFiles.filter(f => f.path !== file.path);
      const next = state.openFiles[0];
      if (next) next.active = true;
      return ok(JSON.stringify({ ok: true, closed: file.path, saved: shouldSave }));
    }

    case 'books': {
      if (state.openFiles.length === 0) return ok('name\tsheets');
      const lines = ['name\tsheets'];
      for (const f of state.openFiles) lines.push(`${f.path}\t${f.sheets.length}`);
      return ok(lines.join('\n'));
    }

    // ─── Sheets ────────────────────────────────────────────────────
    case 'sheets': {
      const file = getActiveFile(state);
      if (!file) return err('xl sheets: no active workbook');
      const lines = ['name\trows\tcols\tactive'];
      for (const s of file.sheets) {
        const isActive = s.active ? '*' : '';
        lines.push([s.name, String(s.rows), String(s.cols), isActive].join('\t'));
      }
      return ok(lines.join('\n'));
    }

    case 'sheet': {
      const file = getActiveFile(state);
      if (!file) return err('xl sheet: no active workbook');
      const name = cmd.positionals[0];
      if (!name) return err('xl sheet: missing <name>');
      const target = file.sheets.find(s => s.name === name);
      if (!target) return err(`xl sheet: not found "${name}"`);
      for (const s of file.sheets) s.active = s === target;
      return ok(JSON.stringify({ ok: true, active: name }));
    }

    case 'sheet.add': {
      const file = getActiveFile(state);
      if (!file) return err('xl sheet.add: no active workbook');
      const name = cmd.flags.name;
      if (!name) return err('xl sheet.add: missing --name');
      if (file.sheets.some(s => s.name === name)) return err(`xl sheet.add: "${name}" already exists`);
      const sheet: XlSheet = { index: file.sheets.length + 1, name, cells: {}, rows: 0, cols: 0 };
      file.sheets.push(sheet);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, added: name, totalSheets: file.sheets.length }));
    }

    case 'sheet.rename': {
      const file = getActiveFile(state);
      if (!file) return err('xl sheet.rename: no active workbook');
      const oldName = cmd.positionals[0];
      const newName = cmd.flags.name;
      if (!oldName || !newName) return err('xl sheet.rename: need <oldName> --name');
      const target = file.sheets.find(s => s.name === oldName);
      if (!target) return err(`xl sheet.rename: not found "${oldName}"`);
      if (file.sheets.some(s => s.name === newName)) return err(`xl sheet.rename: "${newName}" already exists`);
      target.name = newName;
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, from: oldName, to: newName }));
    }

    case 'sheet.delete': {
      const file = getActiveFile(state);
      if (!file) return err('xl sheet.delete: no active workbook');
      const name = cmd.positionals[0];
      if (!name) return err('xl sheet.delete: missing <name>');
      const idx = file.sheets.findIndex(s => s.name === name);
      if (idx < 0) return err(`xl sheet.delete: not found "${name}"`);
      if (file.sheets.length === 1) return err('xl sheet.delete: cannot delete the only sheet');
      const wasActive = file.sheets[idx].active;
      file.sheets.splice(idx, 1);
      file.sheets.forEach((s, i) => { s.index = i + 1; });
      if (wasActive) file.sheets[0].active = true;
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, deleted: name, totalSheets: file.sheets.length }));
    }

    // ─── Read ──────────────────────────────────────────────────────
    case 'get': {
      const file = getActiveFile(state);
      if (!file) return err('xl get: no active workbook');
      const raw = cmd.positionals[0];
      if (!raw) return err('xl get: missing <cell>');
      const ref = parseRef(raw);
      const sheet = resolveSheet(file, ref.sheet);
      if (!sheet) return err(`xl get: sheet "${ref.sheet}" not found`);
      const cell = sheet.cells[ref.address] ?? {};
      const value = cell.value !== undefined ? String(cell.value) : '';
      const formula = cell.formula ?? '';
      const fmt = (cmd.flags.format ?? 'json').toLowerCase();
      if (fmt === 'raw') return ok(value);
      if (fmt === 'tsv') return ok(`${value}\t${formula}`);
      return ok(JSON.stringify({ cell: raw, value, formula }, null, 2));
    }

    case 'range': {
      const file = getActiveFile(state);
      if (!file) return err('xl range: no active workbook');
      const raw = cmd.positionals[0];
      if (!raw) return err('xl range: missing <range>');
      const ref = parseRef(raw);
      const sheet = resolveSheet(file, ref.sheet);
      if (!sheet) return err(`xl range: sheet "${ref.sheet}" not found`);
      const box = parseRange(ref.address);
      if (!box) return err(`xl range: invalid range "${ref.address}"`);
      const lines: string[] = [];
      for (let r = box.start.row; r <= box.end.row; r++) {
        const cols: string[] = [];
        for (let c = box.start.col; c <= box.end.col; c++) {
          const addr = formatAddr({ col: c, row: r });
          const cell = sheet.cells[addr];
          cols.push(cell?.value !== undefined ? String(cell.value) : '');
        }
        lines.push(cols.join('\t'));
      }
      return ok(lines.join('\n'));
    }

    case 'used': {
      const file = getActiveFile(state);
      if (!file) return err('xl used: no active workbook');
      const sheet = getActiveSheet(file);
      if (!sheet) return err('xl used: no active sheet');
      recomputeUsed(sheet);
      const range = sheet.rows > 0 && sheet.cols > 0
        ? `A1:${colLetter(sheet.cols - 1)}${sheet.rows}`
        : 'A1';
      return ok(JSON.stringify({ sheet: sheet.name, range, rows: sheet.rows, cols: sheet.cols }, null, 2));
    }

    case 'find': {
      const file = getActiveFile(state);
      if (!file) return err('xl find: no active workbook');
      const what = cmd.flags.what ?? cmd.positionals[0];
      if (!what) return err('xl find: missing keyword');
      const sheet = getActiveSheet(file);
      if (!sheet) return err('xl find: no active sheet');
      const inExpr = cmd.flags.in;
      let inBox = inExpr ? parseRange(parseRef(inExpr).address) : null;
      if (!inBox) {
        recomputeUsed(sheet);
        if (sheet.rows === 0 || sheet.cols === 0) return ok('address\tvalue');
        inBox = { start: { col: 1, row: 1 }, end: { col: sheet.cols, row: sheet.rows } };
      }
      const lines = ['address\tvalue'];
      for (let r = inBox.start.row; r <= inBox.end.row; r++) {
        for (let c = inBox.start.col; c <= inBox.end.col; c++) {
          const addr = formatAddr({ col: c, row: r });
          const v = sheet.cells[addr]?.value;
          if (v !== undefined && String(v).includes(what)) {
            lines.push(`${addr}\t${String(v)}`);
          }
        }
      }
      return ok(lines.join('\n'));
    }

    // ─── Write ─────────────────────────────────────────────────────
    case 'set': {
      const file = getActiveFile(state);
      if (!file) return err('xl set: no active workbook');
      const raw = cmd.positionals[0];
      if (!raw) return err('xl set: missing <cell>');
      const ref = parseRef(raw);
      const sheet = resolveSheet(file, ref.sheet);
      if (!sheet) return err(`xl set: sheet "${ref.sheet}" not found`);
      if (!parseAddr(ref.address)) return err(`xl set: invalid cell "${ref.address}"`);
      const cell = ensureCell(sheet, ref.address);
      if (cmd.flags.formula !== undefined) {
        cell.formula = cmd.flags.formula;
        cell.value = cmd.flags.formula; // sandbox doesn't evaluate formulas
      } else if (cmd.flags.value !== undefined) {
        const v = cmd.flags.value;
        const t = (cmd.flags.type ?? 'auto').toLowerCase();
        if (t === 'text') {
          cell.value = v;
          cell.numFmt = '@';
        } else if (t === 'number') {
          const n = Number(v);
          if (!Number.isFinite(n)) return err(`xl set: --type=number but "${v}" is not numeric`);
          cell.value = n;
        } else {
          const n = Number(v);
          cell.value = (v !== '' && Number.isFinite(n)) ? n : v;
        }
        cell.formula = undefined;
      } else {
        return err('xl set: missing --value or --formula');
      }
      recomputeUsed(sheet);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, cell: ref.address }));
    }

    case 'fill': {
      const file = getActiveFile(state);
      if (!file) return err('xl fill: no active workbook');
      const raw = cmd.positionals[0];
      if (!raw) return err('xl fill: missing <range>');
      const values = cmd.flags.values;
      if (values === undefined) return err('xl fill: missing --values (TSV)');
      const ref = parseRef(raw);
      const sheet = resolveSheet(file, ref.sheet);
      if (!sheet) return err(`xl fill: sheet "${ref.sheet}" not found`);
      const box = parseRange(ref.address);
      if (!box) return err(`xl fill: invalid range "${ref.address}"`);
      const rows = values.replace(/\\n/g, '\n').split('\n');
      let written = 0;
      for (let i = 0; i < rows.length; i++) {
        const cols = rows[i].split('\t');
        for (let j = 0; j < cols.length; j++) {
          const addr = formatAddr({ col: box.start.col + j, row: box.start.row + i });
          const cell = ensureCell(sheet, addr);
          const v = cols[j];
          const n = Number(v);
          cell.value = (v !== '' && Number.isFinite(n)) ? n : v;
          cell.formula = undefined;
          written++;
        }
      }
      recomputeUsed(sheet);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, written }));
    }

    case 'clear': {
      const file = getActiveFile(state);
      if (!file) return err('xl clear: no active workbook');
      const raw = cmd.positionals[0];
      if (!raw) return err('xl clear: missing <range>');
      const ref = parseRef(raw);
      const sheet = resolveSheet(file, ref.sheet);
      if (!sheet) return err(`xl clear: sheet "${ref.sheet}" not found`);
      const box = parseRange(ref.address);
      if (!box) return err(`xl clear: invalid range "${ref.address}"`);
      let cleared = 0;
      for (let r = box.start.row; r <= box.end.row; r++) {
        for (let c = box.start.col; c <= box.end.col; c++) {
          const addr = formatAddr({ col: c, row: r });
          if (sheet.cells[addr]) {
            sheet.cells[addr].value = undefined;
            sheet.cells[addr].formula = undefined;
            cleared++;
          }
        }
      }
      recomputeUsed(sheet);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, cleared }));
    }

    // ─── Format ────────────────────────────────────────────────────
    case 'fmt': {
      const file = getActiveFile(state);
      if (!file) return err('xl fmt: no active workbook');
      const raw = cmd.positionals[0];
      if (!raw) return err('xl fmt: missing <range>');
      const ref = parseRef(raw);
      const sheet = resolveSheet(file, ref.sheet);
      if (!sheet) return err(`xl fmt: sheet "${ref.sheet}" not found`);
      const box = parseRange(ref.address);
      if (!box) return err(`xl fmt: invalid range "${ref.address}"`);
      const applied: Record<string, unknown> = {};
      const apply = (cell: XlCell) => {
        if (cmd.flags.bold   !== undefined) { cell.bold   = cmd.flags.bold   !== 'false'; applied.bold   = cell.bold; }
        if (cmd.flags.italic !== undefined) { cell.italic = cmd.flags.italic !== 'false'; applied.italic = cell.italic; }
        const sz = parseIntOrNull(cmd.flags.size);
        if (sz !== null) { cell.fontSize = sz; applied.size = sz; }
        if (cmd.flags.color !== undefined)  { cell.color  = cmd.flags.color;  applied.color  = cell.color; }
        if (cmd.flags.bg    !== undefined)  { cell.bg     = cmd.flags.bg;     applied.bg     = cell.bg; }
        if (cmd.flags.align !== undefined)  {
          const a = cmd.flags.align as 'left' | 'center' | 'right';
          cell.align = a;
          applied.align = a;
        }
        if (cmd.flags.numfmt !== undefined) { cell.numFmt = cmd.flags.numfmt; applied.numfmt = cell.numFmt; }
        if (cmd.flags.border !== undefined) {
          const b = cmd.flags.border as XlCell['border'];
          cell.border = b;
          applied.border = b;
        }
      };
      for (let r = box.start.row; r <= box.end.row; r++) {
        for (let c = box.start.col; c <= box.end.col; c++) {
          const addr = formatAddr({ col: c, row: r });
          apply(ensureCell(sheet, addr));
        }
      }
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, range: ref.address, applied }));
    }

    case 'width': {
      const file = getActiveFile(state);
      if (!file) return err('xl width: no active workbook');
      const col = cmd.positionals[0];
      const size = cmd.flags.size ?? cmd.flags.width;
      if (!col || size === undefined) return err('xl width: need <column> --size');
      if (colIndex(col) === 0) return err(`xl width: invalid column "${col}"`);
      const n = Number(size);
      if (!Number.isFinite(n)) return err(`xl width: --size "${size}" not numeric`);
      const sheet = getActiveSheet(file);
      if (!sheet) return err('xl width: no active sheet');
      sheet.colWidths = sheet.colWidths ?? {};
      sheet.colWidths[col.toUpperCase()] = n;
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, column: col.toUpperCase(), size: n }));
    }

    // ─── Rows / Columns ────────────────────────────────────────────
    case 'row.add': {
      const file = getActiveFile(state);
      if (!file) return err('xl row.add: no active workbook');
      const at = parseIntOrNull(cmd.positionals[0]);
      if (at === null || at < 1) return err('xl row.add: missing or invalid <row>');
      const count = parseIntOrNull(cmd.flags.count) ?? 1;
      const sheet = getActiveSheet(file);
      if (!sheet) return err('xl row.add: no active sheet');
      shiftCells(sheet, { byRow: count, fromRow: at });
      recomputeUsed(sheet);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, inserted: count, at }));
    }

    case 'row.delete': {
      const file = getActiveFile(state);
      if (!file) return err('xl row.delete: no active workbook');
      const at = parseIntOrNull(cmd.positionals[0]);
      if (at === null || at < 1) return err('xl row.delete: missing or invalid <row>');
      const count = parseIntOrNull(cmd.flags.count) ?? 1;
      const sheet = getActiveSheet(file);
      if (!sheet) return err('xl row.delete: no active sheet');
      // Delete the rows in [at, at+count)
      for (const addr of Object.keys(sheet.cells)) {
        const a = parseAddr(addr);
        if (!a) continue;
        if (a.row >= at && a.row < at + count) delete sheet.cells[addr];
      }
      shiftCells(sheet, { byRow: -count, fromRow: at + count });
      recomputeUsed(sheet);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, deleted: count, at }));
    }

    case 'col.add': {
      const file = getActiveFile(state);
      if (!file) return err('xl col.add: no active workbook');
      const colArg = cmd.positionals[0];
      const at = colArg ? colIndex(colArg) : 0;
      if (at < 1) return err('xl col.add: missing or invalid <column>');
      const count = parseIntOrNull(cmd.flags.count) ?? 1;
      const sheet = getActiveSheet(file);
      if (!sheet) return err('xl col.add: no active sheet');
      shiftCells(sheet, { byCol: count, fromCol: at });
      recomputeUsed(sheet);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, inserted: count, at: colArg.toUpperCase() }));
    }

    case 'col.delete': {
      const file = getActiveFile(state);
      if (!file) return err('xl col.delete: no active workbook');
      const colArg = cmd.positionals[0];
      const at = colArg ? colIndex(colArg) : 0;
      if (at < 1) return err('xl col.delete: missing or invalid <column>');
      const count = parseIntOrNull(cmd.flags.count) ?? 1;
      const sheet = getActiveSheet(file);
      if (!sheet) return err('xl col.delete: no active sheet');
      for (const addr of Object.keys(sheet.cells)) {
        const a = parseAddr(addr);
        if (!a) continue;
        if (a.col >= at && a.col < at + count) delete sheet.cells[addr];
      }
      shiftCells(sheet, { byCol: -count, fromCol: at + count });
      recomputeUsed(sheet);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, deleted: count, at: colArg.toUpperCase() }));
    }

    // ─── Merge ─────────────────────────────────────────────────────
    case 'merge': {
      const file = getActiveFile(state);
      if (!file) return err('xl merge: no active workbook');
      const raw = cmd.positionals[0];
      if (!raw) return err('xl merge: missing <range>');
      const ref = parseRef(raw);
      const sheet = resolveSheet(file, ref.sheet);
      if (!sheet) return err(`xl merge: sheet "${ref.sheet}" not found`);
      const box = parseRange(ref.address);
      if (!box) return err(`xl merge: invalid range "${ref.address}"`);
      const anchor = formatAddr(box.start);
      const anchorRange = `${anchor}:${formatAddr(box.end)}`;
      ensureCell(sheet, anchor).mergeAnchor = anchorRange;
      for (let r = box.start.row; r <= box.end.row; r++) {
        for (let c = box.start.col; c <= box.end.col; c++) {
          const addr = formatAddr({ col: c, row: r });
          if (addr !== anchor) ensureCell(sheet, addr).mergedInto = anchor;
        }
      }
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, merged: anchorRange }));
    }

    case 'unmerge': {
      const file = getActiveFile(state);
      if (!file) return err('xl unmerge: no active workbook');
      const raw = cmd.positionals[0];
      if (!raw) return err('xl unmerge: missing <range>');
      const ref = parseRef(raw);
      const sheet = resolveSheet(file, ref.sheet);
      if (!sheet) return err(`xl unmerge: sheet "${ref.sheet}" not found`);
      const box = parseRange(ref.address);
      if (!box) return err(`xl unmerge: invalid range "${ref.address}"`);
      for (let r = box.start.row; r <= box.end.row; r++) {
        for (let c = box.start.col; c <= box.end.col; c++) {
          const addr = formatAddr({ col: c, row: r });
          const cell = sheet.cells[addr];
          if (cell) {
            cell.mergeAnchor = undefined;
            cell.mergedInto = undefined;
          }
        }
      }
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, sheet: sheet.name, unmerged: ref.address }));
    }

    // ─── Export ────────────────────────────────────────────────────
    case 'export': {
      const file = getActiveFile(state);
      if (!file) return err('xl export: no active workbook');
      const format = cmd.flags.format;
      const path = cmd.flags.path;
      if (!format || !path) return err('xl export: need --format and --path');
      if (format !== 'pdf' && format !== 'csv') return err(`xl export: unsupported format "${format}" (pdf|csv)`);
      return ok(JSON.stringify({ ok: true, exported: path, format, sheets: file.sheets.length }));
    }

    default:
      return err(`xl ${cmd.action}: not supported in sandbox`);
  }
}

interface ShiftSpec {
  byRow?: number;
  fromRow?: number;
  byCol?: number;
  fromCol?: number;
}

/**
 * Shift all cells whose row >= fromRow by byRow (and/or col >= fromCol by
 * byCol). Negative shifts (deletions) are caller-safe — they assume the
 * caller already removed the deleted cells. Cells that would land at
 * row<1 / col<1 are dropped.
 */
function shiftCells(sheet: XlSheet, spec: ShiftSpec): void {
  const byRow = spec.byRow ?? 0;
  const byCol = spec.byCol ?? 0;
  const fromRow = spec.fromRow ?? 1;
  const fromCol = spec.fromCol ?? 1;
  if (byRow === 0 && byCol === 0) return;

  const entries = Object.entries(sheet.cells);
  // Process in an order that doesn't overwrite cells we still need:
  //  - shift up/left (negative): ascending rows/cols
  //  - shift down/right (positive): descending rows/cols
  const order = (byRow > 0 || byCol > 0) ? -1 : 1;
  entries.sort(([a], [b]) => {
    const aa = parseAddr(a)!;
    const bb = parseAddr(b)!;
    if (aa.row !== bb.row) return order * (aa.row - bb.row);
    return order * (aa.col - bb.col);
  });

  const newCells: Record<string, XlCell> = {};
  for (const [addr, cell] of entries) {
    const a = parseAddr(addr)!;
    let row = a.row;
    let col = a.col;
    if (a.row >= fromRow) row = a.row + byRow;
    if (a.col >= fromCol) col = a.col + byCol;
    if (row < 1 || col < 1) continue;
    newCells[formatAddr({ col, row })] = cell;
  }
  sheet.cells = newCells;
}

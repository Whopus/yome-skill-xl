// skills/yome-skill-xl/backends/sandbox/src/viewerProps.ts
//
// Project an XlWorldState snapshot into the props the viewer
// (viewer/index.html) expects. Stays a pure function so the same
// projection can be applied at every trace step to drive seek().

import type { XlCell, XlSheet, XlWorldState } from './state';

export interface ViewerCell {
  address: string;
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
  mergeAnchor?: string;
  mergedInto?: string;
}

export interface ViewerSheet {
  index: number;
  name: string;
  active: boolean;
  rows: number;
  cols: number;
  cells: ViewerCell[];
  colWidths?: Record<string, number>;
}

export interface ViewerFile {
  path: string;
  active: boolean;
  dirty: boolean;
  sheets: ViewerSheet[];
}

export interface ViewerProps {
  files: ViewerFile[];
  /** Path of the active file, or undefined if no workbook open. */
  activePath?: string;
}

function projectCell(addr: string, c: XlCell): ViewerCell {
  return {
    address: addr,
    value: c.value,
    formula: c.formula,
    bold: c.bold,
    italic: c.italic,
    fontSize: c.fontSize,
    color: c.color,
    bg: c.bg,
    align: c.align,
    numFmt: c.numFmt,
    border: c.border,
    mergeAnchor: c.mergeAnchor,
    mergedInto: c.mergedInto,
  };
}

function projectSheet(s: XlSheet): ViewerSheet {
  return {
    index: s.index,
    name: s.name,
    active: !!s.active,
    rows: s.rows,
    cols: s.cols,
    colWidths: s.colWidths,
    cells: Object.entries(s.cells)
      .filter(([, c]) => c.value !== undefined || c.formula !== undefined || c.bold || c.italic
        || c.fontSize || c.color || c.bg || c.align || c.numFmt || c.border
        || c.mergeAnchor || c.mergedInto)
      .map(([addr, c]) => projectCell(addr, c)),
  };
}

export function toViewerProps(state: XlWorldState): ViewerProps {
  const files: ViewerFile[] = state.openFiles.map(f => ({
    path: f.path,
    active: !!f.active,
    dirty: !!f.dirty,
    sheets: f.sheets.map(projectSheet),
  }));
  const active = files.find(f => f.active);
  return { files, activePath: active?.path };
}

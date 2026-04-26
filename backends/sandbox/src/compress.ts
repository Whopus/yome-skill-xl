// skills/yome-skill-xl/backends/sandbox/src/compress.ts
//
// Tool-result compression functions for the xl domain. Spec 4.4 says the
// signature.json `compress` field's `module` should point at this file.
// During the v0.1 monorepo phase, the runtime registry in
// Server/agent/commands/xlCommands.ts holds the same logic inline.

function trunc(s: string, n = 30): string {
  return s.length > n ? s.slice(0, n) + '...[truncated]' : s;
}

export function compressBooks(content: string): string {
  const lines = content.split('\n').filter(l => l.trim() && /\t/.test(l));
  const isHeader = (lines[0] || '').startsWith('name');
  const rows = isHeader ? lines.slice(1) : lines;
  if (rows.length === 0) return '[compressed] 无打开工作簿';
  return `[compressed] ${rows.length}个工作簿:\n` + rows.map(r => {
    const c = r.split('\t');
    return `${trunc(c[0] || '')} ${c[1] || ''}张表`;
  }).join('\n');
}

export function compressSheets(content: string): string {
  // TSV: name\trows\tcols\tactive
  const lines = content.split('\n').filter(l => l.trim() && /\t/.test(l));
  const isHeader = (lines[0] || '').startsWith('name');
  const rows = isHeader ? lines.slice(1) : lines;
  if (rows.length === 0) return '[compressed] 无工作表';
  return `[compressed] ${rows.length}张表:\n` + rows.map(r => {
    const c = r.split('\t');
    const active = c[3] === '*' ? ' *' : '';
    return `${trunc(c[0] || '')}${active} ${c[1] || ''}行×${c[2] || ''}列`;
  }).join('\n');
}

export function compressRange(content: string): string {
  const lines = content.split('\n').filter(l => l.trim());
  if (lines.length === 0) return '[compressed] 空范围';
  const preview = lines.slice(0, 5).map(l => trunc(l)).join('\n');
  return `[compressed] ${lines.length}行:\n${preview}${lines.length > 5 ? '\n...[truncated]' : ''}`;
}

export function compressUsed(content: string): string {
  try {
    const o = JSON.parse(content) as Record<string, unknown>;
    return `[compressed] ${o.sheet} ${o.range} ${o.rows}行×${o.cols}列`;
  } catch {
    return content.length > 30 ? `[compressed] ${content.slice(0, 30)}...[truncated]` : `[compressed] ${content}`;
  }
}

export function compressFind(content: string): string {
  // TSV: address\tvalue
  const lines = content.split('\n').filter(l => l.trim() && /\t/.test(l));
  const isHeader = (lines[0] || '').startsWith('address');
  const rows = isHeader ? lines.slice(1) : lines;
  if (rows.length === 0) return '[compressed] 无匹配';
  const preview = rows.slice(0, 10).map(r => {
    const c = r.split('\t');
    return `${c[0]} ${trunc(c[1] || '')}`;
  }).join('\n');
  return `[compressed] ${rows.length}个匹配:\n${preview}${rows.length > 10 ? '\n...[truncated]' : ''}`;
}

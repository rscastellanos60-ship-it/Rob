/*
 * analysis.js — Convierte las hojas en estructuras útiles para el tablero:
 * bloques de tablas, indicadores (etiqueta + número), resolución de rangos
 * usados por los gráficos de Excel y sugerencias de gráfico automático.
 */
(function (global) {
  'use strict';

  const MAX_ROWS = 6000;
  const MAX_COLS = 160;

  const MONTHS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic', 'jan', 'apr', 'aug', 'dec', 'set'];
  const isMonth = (s) => {
    const t = String(s).toLowerCase().trim().replace(/\.$/, '');
    return MONTHS.some((m) => t === m || (t.startsWith(m) && t.length <= 10 && /^[a-záéíóú]+(\s*[-/ ]?\s*\d{2,4})?$/.test(t)));
  };
  const isYear = (v) => /^(19|20)\d{2}(\s*[pe*]|\s*\(.*\))?$/i.test(String(v).trim());
  const isQuarter = (s) => /^(t|q|trim\.?|trimestre|sem\.?|s)\s*[1-4]\b/i.test(String(s).trim()) || /^[1-4]\s*(t|q)\b/i.test(String(s).trim());
  const isTimeLike = (v) => v instanceof Date || isYear(v) || isMonth(v) || isQuarter(v) || /^\d{4}[-/]\d{1,2}([-/]\d{1,2})?$/.test(String(v).trim());
  const isTotal = (v) => /^\s*(total|totales|gran total|suma|subtotal)\b/i.test(String(v || ''));

  // Excel guarda los formatos con separadores en inglés; se muestran como en Colombia (1.234,5)
  function localize(text, fmt) {
    if (typeof text !== 'string') return text;
    if (fmt && /[dmyhs]/i.test(String(fmt).replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''))) return text;
    return text.replace(/(\d),(?=\d{3})/g, '$1\u0001').replace(/(\d)\.(\d)/g, '$1,$2').replace(/\u0001/g, '.');
  }

  function cellText(c) {
    if (!c) return '';
    if (c.w != null) return c.t === 'n' ? localize(String(c.w), c.z) : String(c.w);
    if (c.v instanceof Date) return c.v.toLocaleDateString('es-CO');
    return c.v == null ? '' : String(c.v);
  }
  function isNum(c) {
    return !!c && c.t === 'n' && typeof c.v === 'number' && isFinite(c.v);
  }
  function isBlank(c) {
    return !c || c.v == null || (typeof c.v === 'string' && c.v.trim() === '');
  }

  /* ---------- rejilla de celdas ---------- */
  function buildGrid(ws) {
    if (!ws || !ws['!ref']) return null;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const r1 = range.s.r, c1 = range.s.c;
    const r2 = Math.min(range.e.r, r1 + MAX_ROWS - 1);
    const c2 = Math.min(range.e.c, c1 + MAX_COLS - 1);
    const rows = [];
    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let c = c1; c <= c2; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        row.push(cell && !isBlank(cell) ? cell : null);
      }
      rows.push(row);
    }
    const hiddenRows = new Set();
    (ws['!rows'] || []).forEach((p, i) => p && p.hidden && hiddenRows.add(i));
    const hiddenCols = new Set();
    (ws['!cols'] || []).forEach((p, i) => p && p.hidden && hiddenCols.add(i));
    // las celdas combinadas ocupan su rango (sirve para detectar bloques)
    const merged = new Map();
    for (const m of ws['!merges'] || []) {
      for (let r = m.s.r; r <= m.e.r; r++)
        for (let c = m.s.c; c <= m.e.c; c++) if (r !== m.s.r || c !== m.s.c) merged.set(r + ':' + c, { r: m.s.r, c: m.s.c });
    }
    return { r1, c1, r2, c2, rows, hiddenRows, hiddenCols, merged, truncated: range.e.r > r2 || range.e.c > c2 };
  }

  const at = (g, r, c) => (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2 ? null : g.rows[r - g.r1][c - g.c1]);
  const occupied = (g, r, c) => !!at(g, r, c) || g.merged.has(r + ':' + c);

  /* ---------- bloques contiguos ---------- */
  function findBlocks(g) {
    const H = g.r2 - g.r1 + 1, W = g.c2 - g.c1 + 1;
    const seen = new Uint8Array(H * W);
    const blocks = [];
    for (let r = g.r1; r <= g.r2; r++) {
      for (let c = g.c1; c <= g.c2; c++) {
        const k = (r - g.r1) * W + (c - g.c1);
        if (seen[k] || !occupied(g, r, c)) continue;
        const b = { r1: r, r2: r, c1: c, c2: c, count: 0 };
        const stack = [[r, c]];
        seen[k] = 1;
        while (stack.length) {
          const [rr, cc] = stack.pop();
          b.count++;
          if (rr < b.r1) b.r1 = rr;
          if (rr > b.r2) b.r2 = rr;
          if (cc < b.c1) b.c1 = cc;
          if (cc > b.c2) b.c2 = cc;
          for (let dr = -1; dr <= 1; dr++)
            for (let dc = -1; dc <= 1; dc++) {
              const nr = rr + dr, nc = cc + dc;
              if (nr < g.r1 || nr > g.r2 || nc < g.c1 || nc > g.c2) continue;
              const nk = (nr - g.r1) * W + (nc - g.c1);
              if (seen[nk] || !occupied(g, nr, nc)) continue;
              seen[nk] = 1;
              stack.push([nr, nc]);
            }
        }
        blocks.push(b);
      }
    }
    // unir bloques cuyos rectángulos se solapan
    let changed = true;
    while (changed) {
      changed = false;
      outer: for (let i = 0; i < blocks.length; i++)
        for (let j = i + 1; j < blocks.length; j++) {
          const a = blocks[i], b = blocks[j];
          if (a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2) {
            a.r1 = Math.min(a.r1, b.r1); a.r2 = Math.max(a.r2, b.r2);
            a.c1 = Math.min(a.c1, b.c1); a.c2 = Math.max(a.c2, b.c2);
            a.count += b.count;
            blocks.splice(j, 1);
            changed = true;
            break outer;
          }
        }
    }
    return blocks.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
  }

  /* ---------- tablas ---------- */
  function rowStats(g, r, c1, c2) {
    let str = 0, num = 0, filled = 0;
    for (let c = c1; c <= c2; c++) {
      const cell = at(g, r, c);
      if (!cell) continue;
      filled++;
      if (isNum(cell) || cell.t === 'd') num++;
      else str++;
    }
    return { str, num, filled };
  }

  function detectTable(g, b, sheetName) {
    const width = b.c2 - b.c1 + 1;
    if (width < 2 || b.r2 - b.r1 < 2) return null;
    // fila de encabezados: mayoría de texto y debajo hay números
    let header = -1;
    for (let r = b.r1; r <= Math.min(b.r2 - 1, b.r1 + 8); r++) {
      const s = rowStats(g, r, b.c1, b.c2);
      // la fila siguiente con datos (puede haber una fila vacía o combinada en medio)
      let below = rowStats(g, r + 1, b.c1, b.c2);
      for (let k = 2; k <= 3 && !below.filled && r + k <= b.r2; k++) below = rowStats(g, r + k, b.c1, b.c2);
      const yearHeader = (() => {
        let y = 0;
        for (let c = b.c1; c <= b.c2; c++) if (at(g, r, c) && isYear(cellText(at(g, r, c)))) y++;
        return y >= 2 && y >= s.filled - 1;
      })();
      // las celdas combinadas cuentan como una sola columna
      let cols = 0;
      for (let c = b.c1; c <= b.c2; c++) if (!g.merged.has(r + ':' + c)) cols++;
      if (s.filled >= Math.max(2, Math.ceil(cols * 0.5)) && (s.str >= s.num || yearHeader) && below.num >= 1) {
        header = r;
        break;
      }
    }
    if (header < 0) return null;
    const titleParts = [];
    for (let r = b.r1; r < header; r++) {
      const s = rowStats(g, r, b.c1, b.c2);
      if (s.filled && s.filled <= 2)
        for (let c = b.c1; c <= b.c2; c++) if (at(g, r, c) && !isNum(at(g, r, c))) titleParts.push(cellText(at(g, r, c)));
    }
    // columnas: nombre del encabezado (con el de la fila superior si es un grupo combinado)
    const columns = [];
    for (let c = b.c1; c <= b.c2; c++) {
      let name = cellText(at(g, header, c)).trim();
      const above = header > b.r1 ? g.merged.get(header - 1 + ':' + c) || { r: header - 1, c } : null;
      const parent = above ? cellText(at(g, above.r, above.c)).trim() : '';
      if (parent && !/^[\d.,\s%]+$/.test(parent) && parent !== name && rowStats(g, header - 1, b.c1, b.c2).filled < width && !titleParts.includes(parent)) name = name ? parent + ' · ' + name : parent;
      columns.push({ index: c, name: name || XLSX.utils.encode_col(c) });
    }
    const rows = [];
    for (let r = header + 1; r <= b.r2; r++) {
      if (g.hiddenRows.has(r)) continue;
      const cells = columns.map((col) => at(g, r, col.index));
      if (cells.every((x) => !x)) continue;
      rows.push({ r, cells });
    }
    if (rows.length < 1) return null;
    columns.forEach((col, i) => {
      let n = 0, d = 0, filled = 0;
      for (const row of rows) {
        const cell = row.cells[i];
        if (!cell) continue;
        filled++;
        if (cell.t === 'd' || (isNum(cell) && /[dmy]{2}/i.test(cell.z || '') && !/0\.0|#/.test(cell.z || ''))) d++;
        else if (isNum(cell)) n++;
      }
      col.type = filled === 0 ? 'empty' : d / filled >= 0.7 ? 'date' : n / filled >= 0.6 ? 'number' : 'text';
      col.format = (rows.find((row) => isNum(row.cells[i])) || { cells: [] }).cells[i]?.z || null;
      col.hidden = g.hiddenCols.has(col.index);
    });
    const numeric = columns.filter((c) => c.type === 'number');
    if (!numeric.length) return null;
    const range = XLSX.utils.encode_range({ s: { r: b.r1, c: b.c1 }, e: { r: b.r2, c: b.c2 } });
    return {
      id: sheetName + '!' + range,
      sheet: sheetName,
      range,
      bounds: { r1: b.r1, r2: b.r2, c1: b.c1, c2: b.c2 },
      title: titleParts.join(' — '),
      headerRow: header,
      columns: columns.filter((c) => c.type !== 'empty' && !c.hidden),
      rows: rows.map((row) => ({ r: row.r, cells: columns.map((c, i) => ({ c, cell: row.cells[i] })).filter((x) => x.c.type !== 'empty' && !x.c.hidden).map((x) => x.cell) })),
    };
  }

  /* ---------- indicadores (etiqueta + valor) ---------- */
  function detectKpis(g, b) {
    const out = [];
    for (let r = b.r1; r <= b.r2; r++) {
      for (let c = b.c1; c <= b.c2; c++) {
        const cell = at(g, r, c);
        if (!isNum(cell)) continue;
        let label = null;
        for (let k = 1; k <= 3 && !label; k++) {
          const left = at(g, r, c - k);
          if (left && !isNum(left)) label = cellText(left);
          else if (left) break;
        }
        for (let k = 1; k <= 2 && !label; k++) {
          const up = at(g, r - k, c);
          if (up && !isNum(up)) label = cellText(up);
          else if (up) break;
        }
        if (!label) continue;
        const right = at(g, r, c + 1);
        const unit = right && !isNum(right) && cellText(right).length <= 12 ? cellText(right) : '';
        out.push({ label: label.trim(), value: cell.v, text: cellText(cell), unit, format: cell.z || null, address: XLSX.utils.encode_cell({ r, c }), r, c, hidden: g.hiddenRows.has(r) || g.hiddenCols.has(c) });
      }
    }
    return out;
  }

  function analyzeSheet(ws, name) {
    const g = buildGrid(ws);
    if (!g) return { name, tables: [], kpis: [], texts: [], empty: true };
    const blocks = findBlocks(g);
    const tables = [], kpis = [], texts = [];
    for (const b of blocks) {
      const t = detectTable(g, b, name);
      if (t && t.rows.length >= 2) {
        tables.push(t);
        continue;
      }
      const k = detectKpis(g, b);
      if (k.length) kpis.push(...k);
      else if (b.count <= 3) {
        for (let r = b.r1; r <= b.r2; r++)
          for (let c = b.c1; c <= b.c2; c++) {
            const cell = at(g, r, c);
            if (cell && typeof cell.v === 'string' && cell.v.trim().length > 2) texts.push({ text: cell.v.trim(), r, c });
          }
      }
    }
    // título de la tabla: texto suelto en las filas inmediatamente superiores
    for (const t of tables) {
      if (t.title) continue;
      const cand = texts
        .filter((x) => x.r < t.bounds.r1 && x.r >= t.bounds.r1 - 3 && x.c >= t.bounds.c1 - 1 && x.c <= t.bounds.c2)
        .sort((a, b) => b.r - a.r)[0];
      if (cand) {
        t.title = cand.text;
        cand.usedAsTitle = true;
      }
    }
    return { name, tables, kpis, texts: texts.filter((x) => !x.usedAsTitle), allTexts: texts, truncated: g.truncated, grid: g };
  }

  /* ---------- rangos de fórmulas (gráficos de Excel) ---------- */
  function splitTopLevel(s) {
    const parts = [];
    let depth = 0, inQ = false, cur = '';
    for (const ch of s) {
      if (ch === "'") inQ = !inQ;
      if (!inQ && ch === '(') depth++;
      if (!inQ && ch === ')') depth--;
      if (!inQ && depth === 0 && ch === ',') {
        parts.push(cur);
        cur = '';
      } else cur += ch;
    }
    if (cur) parts.push(cur);
    return parts;
  }

  function resolveRef(f, wb, depth = 0) {
    if (!f || depth > 4) return null;
    let s = String(f).trim().replace(/^=/, '');
    while (s.startsWith('(') && s.endsWith(')')) s = s.slice(1, -1);
    const parts = splitTopLevel(s);
    if (parts.length > 1) {
      const all = parts.map((p) => resolveRef(p, wb, depth + 1)).filter(Boolean);
      return all.length ? all.flat() : null;
    }
    const m = s.match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
    if (!m) {
      const nm = s.replace(/^\[\d+\]!?/, '').replace(/^.*!/, '');
      const def = ((wb.Workbook && wb.Workbook.Names) || []).find((n) => n.Name.toLowerCase() === nm.toLowerCase());
      return def ? resolveRef(def.Ref, wb, depth + 1) : null;
    }
    const sheet = (m[1] ? m[1].replace(/''/g, "'") : m[2]).replace(/^\[[^\]]*\]/, '');
    const ws = wb.Sheets[sheet];
    if (!ws) return null;
    const addr = m[3].replace(/\$/g, '');
    if (!/^[A-Z]+\d+(:[A-Z]+\d+)?$/i.test(addr)) {
      const def = ((wb.Workbook && wb.Workbook.Names) || []).find((n) => n.Name.toLowerCase() === addr.toLowerCase());
      return def ? resolveRef(def.Ref, wb, depth + 1) : null;
    }
    const range = XLSX.utils.decode_range(addr);
    const nR = range.e.r - range.s.r + 1, nC = range.e.c - range.s.c + 1;
    const out = [];
    const alongRows = nR >= nC;
    const len = alongRows ? nR : nC, depthN = alongRows ? nC : nR;
    for (let i = 0; i < len; i++) {
      const pieces = [];
      let last = null;
      for (let j = 0; j < depthN; j++) {
        const r = range.s.r + (alongRows ? i : j), c = range.s.c + (alongRows ? j : i);
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell && !isBlank(cell)) {
          pieces.push(cellText(cell));
          last = cell;
        }
      }
      if (depthN > 1) out.push({ v: pieces.join(' · '), w: pieces.join(' · '), t: 's' });
      else out.push(last ? { v: last.v, w: last.w != null ? String(last.w) : cellText(last), t: last.t, z: last.z } : { v: null, w: '', t: 'z' });
    }
    return out;
  }

  function refSheets(f) {
    if (!f) return [];
    const out = new Set();
    const re = /(?:'((?:[^']|'')+)'|([A-Za-z0-9_À-ſ.]+))!/g;
    let m;
    while ((m = re.exec(f))) out.add((m[1] ? m[1].replace(/''/g, "'") : m[2]).replace(/^\[[^\]]*\]/, ''));
    return Array.from(out);
  }

  function formatNumber(v, fmt) {
    if (v == null || v === '' || isNaN(v)) return '';
    if (fmt && fmt !== 'General') {
      try {
        return localize(XLSX.SSF.format(fmt, +v), fmt);
      } catch (e) {
        /* formato no soportado */
      }
    }
    const abs = Math.abs(v);
    return new Intl.NumberFormat('es-CO', { maximumFractionDigits: abs >= 100 ? 0 : abs >= 1 ? 2 : 4 }).format(v);
  }

  // Valores de una serie: primero la caché que Excel guardó, si no, las celdas
  function seriesValues(ref, wb, numeric) {
    if (!ref) return { values: null, fmt: null };
    if (ref.cache && ref.cache.some((x) => x != null && x !== '')) {
      const fmt = ref.formatCode && ref.formatCode !== 'General' ? ref.formatCode : null;
      if (numeric) return { values: ref.cache.map((x) => (x == null || x === '' || isNaN(+x) ? null : +x)), fmt };
      return {
        values: ref.cache.map((x) => {
          if (x == null) return '';
          if (fmt && ref.numeric && !isNaN(+x)) return formatNumber(+x, fmt);
          return String(x);
        }),
        fmt,
      };
    }
    const cells = resolveRef(ref.f, wb);
    if (!cells) return { values: null, fmt: null };
    if (numeric) {
      const fmtCell = cells.find((c) => c.t === 'n' && c.z);
      return { values: cells.map((c) => (typeof c.v === 'number' ? c.v : c.v == null || isNaN(+c.v) || c.v === '' ? null : +c.v)), fmt: fmtCell ? fmtCell.z : null };
    }
    return { values: cells.map((c) => cellText(c)), fmt: null };
  }

  function seriesName(name, wb, fallback) {
    if (!name) return fallback;
    if (name.text) return name.text;
    if (name.f) {
      const cells = resolveRef(name.f, wb);
      const t = cells && cells.map((c) => c.w || c.v).filter(Boolean).join(' ');
      if (t) return String(t);
    }
    return fallback;
  }

  /* ---------- sugerencia automática de gráfico para una tabla ---------- */
  function suggestChart(table) {
    const cols = table.columns;
    let cat = cols.find((c) => c.type === 'text' || c.type === 'date');
    const first = cols[0];
    if ((!cat || cols.indexOf(cat) > 0) && first.type === 'number') {
      const vals = table.rows.map((r) => r.cells[0]).filter(Boolean).map((c) => c.v);
      if (vals.length && vals.every((v) => Number.isInteger(v) && v > 1900 && v < 2200)) cat = first;
    }
    if (!cat) cat = null;
    const numeric = cols.filter((c) => c.type === 'number' && c !== cat && !/^(id|c[oó]d(igo)?|n[°º.]?|no\.?|item|#)$/i.test(c.name.trim()));
    if (!numeric.length) return null;
    const catValues = cat ? table.rows.map((r) => r.cells[cols.indexOf(cat)]) : [];
    const catTime = cat && (cat.type === 'date' || catValues.filter(Boolean).filter((c) => isTimeLike(c.v instanceof Date ? c.v : cellText(c))).length >= catValues.filter(Boolean).length * 0.7);
    const headerTime = numeric.filter((c) => isTimeLike(c.name.replace(/^.*·\s*/, ''))).length >= Math.max(3, numeric.length * 0.6);
    const transpose = !catTime && headerTime && table.rows.length <= 25;
    const share = numeric.length === 1 && /%|particip|particip|share|distribuci/i.test(numeric[0].name) && table.rows.length <= 8;
    let type = 'bar';
    if (transpose || catTime) type = 'line';
    if (share) type = 'doughnut';
    if (!transpose && !catTime && table.rows.length > 30) type = 'line';
    const horizontal = type === 'bar' && !catTime && cat && catValues.some((c) => c && cellText(c).length > 14);
    return { cat, numeric, transpose, type, horizontal, timeSeries: !!(catTime || transpose) };
  }

  function chartDataFromTable(table, spec, opts = {}) {
    const cols = table.columns;
    const ci = spec.cat ? cols.indexOf(spec.cat) : -1;
    const rows = table.rows.filter((r) => !(ci >= 0 && isTotal(cellText(r.cells[ci]))) && !isTotal(cellText(r.cells[0])));
    const active = opts.columns || spec.numeric.slice(0, 6);
    if (opts.transpose ?? spec.transpose) {
      const labels = active.map((c) => c.name);
      const series = rows.slice(0, 12).map((r, i) => ({
        name: ci >= 0 ? cellText(r.cells[ci]) : 'Fila ' + (i + 1),
        values: active.map((c) => {
          const cell = r.cells[cols.indexOf(c)];
          return isNum(cell) ? cell.v : null;
        }),
        fmt: active[0] && active[0].format,
      }));
      return { labels, series };
    }
    const labels = rows.map((r, i) => (ci >= 0 ? cellText(r.cells[ci]) : String(i + 1)));
    const series = active.map((c) => ({
      name: c.name,
      values: rows.map((r) => {
        const cell = r.cells[cols.indexOf(c)];
        return isNum(cell) ? cell.v : null;
      }),
      fmt: c.format,
    }));
    return { labels, series };
  }

  /* ---------- indicadores de series de tiempo ---------- */
  function timeKpis(table, spec) {
    if (!spec || !spec.timeSeries || spec.transpose) return [];
    const data = chartDataFromTable(table, spec);
    return data.series.slice(0, 4).map((s) => {
      const pts = s.values.map((v, i) => ({ v, label: data.labels[i] })).filter((p) => p.v != null);
      if (pts.length < 2) return null;
      const last = pts[pts.length - 1], prev = pts[pts.length - 2];
      return {
        label: s.name,
        value: last.v,
        text: formatNumber(last.v, s.fmt),
        period: last.label,
        prevPeriod: prev.label,
        delta: prev.v ? (last.v - prev.v) / Math.abs(prev.v) : null,
        spark: pts.slice(-24).map((p) => p.v),
      };
    }).filter(Boolean);
  }

  // Textos y fecha de la cabecera de una hoja (primeras filas)
  function sheetHeader(a) {
    const g = a && a.grid;
    if (!g) return { title: null, period: null };
    let title = null, period = null;
    for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 8); r++)
      for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 40); c++) {
        const cell = at(g, r, c);
        if (!cell) continue;
        if (!title && typeof cell.v === 'string' && cell.v.trim().length > 6 && !/^\d/.test(cell.v)) title = cell.v.trim();
        if (!period && (cell.t === 'd' || (isNum(cell) && /[my]/i.test(String(cell.z || '').replace(/"[^"]*"/g, ''))))) period = cellText(cell);
      }
    return { title, period };
  }

  global.Analysis = {
    sheetHeader,
    analyzeSheet,
    resolveRef,
    refSheets,
    seriesValues,
    seriesName,
    suggestChart,
    chartDataFromTable,
    timeKpis,
    formatNumber,
    cellText,
    isNum,
    isTimeLike,
    isTotal,
    localize,
  };
})(window);

/*
 * vba.js — Extrae y analiza las macros (VBA) del archivo.
 * Lee el contenedor OLE (vbaProject.bin o el propio .xls), descomprime cada
 * módulo según MS-OVBA y resume qué hace cada procedimiento: a qué hojas
 * navega, si actualiza tablas dinámicas, filtra, copia datos, etc.
 */
(function (global) {
  'use strict';

  /* ---------- MS-OVBA 2.4.1: descompresión ---------- */
  function decompress(buf, offset) {
    if (buf[offset] !== 1) throw new Error('Firma de contenedor comprimido inválida');
    const out = [];
    let pos = offset + 1;
    while (pos + 1 < buf.length) {
      const header = buf[pos] | (buf[pos + 1] << 8);
      const chunkStart = pos;
      pos += 2;
      const size = (header & 0x0fff) + 3;
      const compressed = (header >> 15) & 1;
      const chunkEnd = Math.min(buf.length, chunkStart + size);
      const decStart = out.length;
      if (!compressed) {
        for (let i = 0; i < 4096 && pos < buf.length; i++) out.push(buf[pos++]);
        continue;
      }
      while (pos < chunkEnd) {
        const flags = buf[pos++];
        for (let bit = 0; bit < 8 && pos < chunkEnd; bit++) {
          if (((flags >> bit) & 1) === 0) {
            out.push(buf[pos++]);
          } else {
            const token = buf[pos] | (buf[pos + 1] << 8);
            pos += 2;
            const d = out.length - decStart;
            let bitCount = 4;
            while ((1 << bitCount) < d && bitCount < 12) bitCount++;
            const lengthMask = 0xffff >> bitCount;
            const length = (token & lengthMask) + 3;
            const off = (token >> (16 - bitCount)) + 1;
            const src = out.length - off;
            if (src < 0) throw new Error('Token de copia fuera de rango');
            for (let k = 0; k < length; k++) out.push(out[src + k]);
          }
        }
      }
    }
    return Uint8Array.from(out);
  }

  function decoder(codepage) {
    const names = { 1252: 'windows-1252', 65001: 'utf-8', 1200: 'utf-16le', 28591: 'iso-8859-1' };
    try {
      return new TextDecoder(names[codepage] || 'windows-' + codepage);
    } catch (e) {
      return new TextDecoder('windows-1252');
    }
  }

  const toBytes = (content) => (content instanceof Uint8Array ? content : Uint8Array.from(content || []));

  /* ---------- MS-OVBA 2.3.4.2: flujo "dir" ---------- */
  function parseDir(bytes) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let pos = 0;
    let codepage = 1252;
    const modules = [];
    let cur = null;
    while (pos + 6 <= bytes.length) {
      const id = dv.getUint16(pos, true);
      let size = dv.getUint32(pos + 2, true);
      pos += 6;
      if (id === 0x0009) size = 6; // PROJECTVERSION: el tamaño declarado no incluye MinorVersion
      const data = bytes.subarray(pos, pos + size);
      switch (id) {
        case 0x0003:
          codepage = dv.getUint16(pos, true);
          break;
        case 0x0019:
          cur = { name: decoder(codepage).decode(data), stream: null, offset: 0, type: 'standard' };
          modules.push(cur);
          break;
        case 0x0047:
          if (cur) cur.name = new TextDecoder('utf-16le').decode(data) || cur.name;
          break;
        case 0x001a:
          if (cur) cur.stream = decoder(codepage).decode(data);
          break;
        case 0x0032:
          if (cur) cur.stream = new TextDecoder('utf-16le').decode(data) || cur.stream;
          break;
        case 0x0031:
          if (cur) cur.offset = dv.getUint32(pos, true);
          break;
        case 0x0022:
          if (cur) cur.type = 'document'; // módulo de hoja/libro o de clase
          break;
        default:
          break;
      }
      pos += size;
    }
    return { codepage, modules };
  }

  // Plan B: buscar contenedores comprimidos que contengan "Attribute VB_Name"
  function scanStream(bytes, codepage) {
    for (let i = 0; i < bytes.length - 3; i++) {
      if (bytes[i] !== 1) continue;
      const header = bytes[i + 1] | (bytes[i + 2] << 8);
      if (((header >> 12) & 0x7) !== 0b011) continue;
      try {
        const text = decoder(codepage).decode(decompress(bytes, i));
        if (/Attribute VB_Name/.test(text)) return text;
      } catch (e) {
        /* seguir buscando */
      }
    }
    return null;
  }

  function findVbaRoot(cfb) {
    const path = cfb.FullPaths.find((p) => /(^|\/)VBA\/dir$/i.test(p));
    return path ? path.slice(0, -3) : null;
  }

  function extractModules(cfbData) {
    const CFB = global.XLSX && global.XLSX.CFB;
    if (!CFB || !cfbData) return [];
    const cfb = CFB.read(toBytes(cfbData), { type: 'array' });
    const root = findVbaRoot(cfb);
    if (!root) return [];
    const entry = (p) => {
      const idx = cfb.FullPaths.findIndex((x) => x.toLowerCase() === p.toLowerCase());
      return idx >= 0 ? cfb.FileIndex[idx] : null;
    };
    const dirEntry = entry(root + 'dir');
    const modules = [];
    let codepage = 1252;
    try {
      const dir = parseDir(decompress(toBytes(dirEntry.content), 0));
      codepage = dir.codepage;
      for (const m of dir.modules) {
        const s = entry(root + (m.stream || m.name));
        if (!s) continue;
        const bytes = toBytes(s.content);
        let code;
        try {
          code = decoder(codepage).decode(decompress(bytes, m.offset));
        } catch (e) {
          code = scanStream(bytes, codepage);
        }
        if (code != null) modules.push({ name: m.name, kind: m.type, code });
      }
    } catch (e) {
      // dir dañado: recorrer todos los flujos de la carpeta VBA
      cfb.FullPaths.forEach((p, i) => {
        if (!p.startsWith(root) || /\/(dir|_VBA_PROJECT|__SRP_\d+|PROJECT(wm)?)$/i.test(p)) return;
        const code = scanStream(toBytes(cfb.FileIndex[i].content), codepage);
        if (code) modules.push({ name: p.slice(root.length), kind: 'standard', code });
      });
    }
    return modules;
  }

  /* ---------- análisis del código ---------- */
  const ACTIONS = [
    { key: 'refresh', label: 'Actualiza datos / tablas dinámicas', re: /\bRefreshAll\b|\.RefreshTable\b|PivotCache[\s\S]{0,40}\.Refresh|\.Refresh\b/i },
    { key: 'pivotfilter', label: 'Filtra tablas dinámicas', re: /\bPivotFields\b|\bPivotItems\b|\.CurrentPage\b|\bSlicerCache/i },
    { key: 'filter', label: 'Aplica filtros', re: /\bAutoFilter\b|\.AdvancedFilter\b|ShowAllData/i },
    { key: 'nav', label: 'Navega entre hojas', re: /\.(Select|Activate)\b|Application\.Goto/i },
    { key: 'chart', label: 'Modifica gráficos', re: /\bChartObjects\b|\bSeriesCollection\b|\bSetSourceData\b|\.ChartType\b/i },
    { key: 'copy', label: 'Copia y pega datos', re: /\.Copy\b|\bPasteSpecial\b|\.Value\s*=\s*[\w.()"]+\.Value/i },
    { key: 'import', label: 'Importa datos externos', re: /Workbooks\.Open|GetOpenFilename|QueryTables|\.Connections\b|ADODB|Power\s*Query|ListObjects\([^)]*\)\.QueryTable/i },
    { key: 'export', label: 'Exporta o imprime', re: /ExportAsFixedFormat|\.PrintOut\b|SaveAs\b|SaveCopyAs/i },
    { key: 'sort', label: 'Ordena datos', re: /\.Sort\b|SortFields/i },
    { key: 'clear', label: 'Limpia o borra datos', re: /\.ClearContents\b|\.Clear\b|\.Delete\b/i },
    { key: 'visibility', label: 'Muestra / oculta filas, columnas u hojas', re: /\.Hidden\s*=|\.Visible\s*=/i },
    { key: 'calc', label: 'Calcula / escribe fórmulas', re: /\.Formula(R1C1|Local)?\s*=|\.Calculate\b|WorksheetFunction/i },
    { key: 'loop', label: 'Recorre filas en bucle', re: /\bFor\s+Each\b|\bFor\s+\w+\s*=|\bDo\s+(While|Until)\b/i },
    { key: 'ui', label: 'Pide datos o muestra mensajes', re: /\bMsgBox\b|\bInputBox\b|\.Show\b/i },
  ];

  const EVENTS = /^(Workbook_(Open|BeforeClose|SheetChange|SheetActivate|BeforeSave)|Worksheet_(Change|Activate|SelectionChange|Calculate|PivotTableUpdate)|Auto_Open)$/i;

  function splitProcedures(code) {
    const lines = code.split(/\r?\n/);
    const procs = [];
    let cur = null;
    const startRe = /^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_À-ſ][\wÀ-ſ]*)/i;
    const endRe = /^\s*End\s+(Sub|Function|Property)\b/i;
    lines.forEach((line, i) => {
      if (/^Attribute\s/i.test(line)) return;
      const m = line.match(startRe);
      if (m && !cur) {
        cur = { name: m[3], kind: m[2].replace(/\s+/g, ' '), scope: m[1] || 'Public', start: i, lines: [line] };
        return;
      }
      if (cur) {
        cur.lines.push(line);
        if (endRe.test(line)) {
          cur.code = cur.lines.join('\n');
          delete cur.lines;
          procs.push(cur);
          cur = null;
        }
      }
    });
    return procs;
  }

  function stripComments(code) {
    return code
      .split(/\r?\n/)
      .map((l) => {
        let inStr = false;
        for (let i = 0; i < l.length; i++) {
          if (l[i] === '"') inStr = !inStr;
          else if (l[i] === "'" && !inStr) return l.slice(0, i);
        }
        return l.replace(/^\s*Rem\s.*$/i, '');
      })
      .join('\n');
  }

  function analyzeProcedure(proc, ctx) {
    const code = stripComments(proc.code);
    const sheets = new Set();
    const nav = new Set();
    const sheetRe = /(?:Sheets|Worksheets|Charts)\s*\(\s*"([^"]+)"\s*\)/gi;
    let m;
    while ((m = sheetRe.exec(code))) {
      const name = ctx.matchSheet(m[1]);
      if (!name) continue;
      sheets.add(name);
      const tail = code.slice(m.index + m[0].length, m.index + m[0].length + 30);
      if (/^\s*\.\s*(Select|Activate)\b/i.test(tail)) nav.add(name);
    }
    // referencias por nombre de código (Hoja1.Range..., Sheet2.Activate)
    for (const [codeName, sheet] of Object.entries(ctx.codeNames)) {
      const re = new RegExp('\\b' + codeName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*\\.\\s*(\\w+)', 'gi');
      while ((m = re.exec(code))) {
        sheets.add(sheet);
        if (/^(Select|Activate)$/i.test(m[1])) nav.add(sheet);
      }
    }
    const gotoRe = /Application\.Goto\s+(?:Reference:=)?\s*(?:Sheets|Worksheets)\("([^"]+)"\)/gi;
    while ((m = gotoRe.exec(code))) {
      const n = ctx.matchSheet(m[1]);
      if (n) nav.add(n);
    }
    const ranges = Array.from(new Set((code.match(/Range\(\s*"([^"]+)"\s*\)/gi) || []).map((r) => r.replace(/^Range\(\s*"|"\s*\)$/g, '')))).slice(0, 12);
    const pivots = Array.from(new Set((code.match(/PivotTables\(\s*"([^"]+)"\s*\)/gi) || []).map((r) => r.replace(/^PivotTables\(\s*"|"\s*\)$/gi, ''))));
    const calls = Array.from(new Set((code.match(/\b(?:Call\s+)([A-Za-z_]\w*)/gi) || []).map((c) => c.replace(/^Call\s+/i, ''))));
    const actions = ACTIONS.filter((a) => a.re.test(code)).map((a) => a.label);
    const messages = Array.from(new Set((code.match(/MsgBox\s*\(?\s*"([^"]+)"/gi) || []).map((x) => x.replace(/^MsgBox\s*\(?\s*"/i, '').replace(/"$/, '')))).slice(0, 3);
    return {
      ...proc,
      sheets: Array.from(sheets),
      navigatesTo: Array.from(nav),
      ranges,
      pivots,
      calls,
      actions,
      messages,
      isEvent: EVENTS.test(proc.name),
      lineCount: proc.code.split('\n').length,
    };
  }

  function analyze(modules, workbookSheets, codeNames) {
    const lower = new Map(workbookSheets.map((s) => [s.toLowerCase().trim(), s]));
    const ctx = {
      matchSheet: (n) => lower.get(String(n).toLowerCase().trim()) || null,
      codeNames: codeNames || {},
    };
    const result = modules.map((mod) => {
      const procs = splitProcedures(mod.code).map((p) => analyzeProcedure(p, ctx));
      const body = mod.code.split(/\r?\n/).filter((l) => !/^Attribute\s/i.test(l)).join('\n').trim();
      const owner = ctx.codeNames[mod.name] || null;
      return { name: mod.name, kind: owner ? 'sheet' : mod.kind === 'document' ? 'class' : 'standard', owner, code: body, procedures: procs };
    });
    return result.filter((m) => m.code.replace(/Option\s+Explicit/gi, '').trim());
  }

  function readMacros(wb, rawBuffer) {
    let source = wb.vbaraw || null;
    if (!source) {
      // .xls: las macros viven dentro del mismo contenedor OLE
      try {
        const bytes = new Uint8Array(rawBuffer);
        if (bytes[0] === 0xd0 && bytes[1] === 0xcf) source = bytes;
      } catch (e) {
        source = null;
      }
    }
    if (!source) return { modules: [], available: false };
    try {
      const sheetNames = wb.SheetNames.slice();
      const codeNames = {};
      ((wb.Workbook && wb.Workbook.Sheets) || []).forEach((s, i) => {
        if (s && s.CodeName) codeNames[s.CodeName] = sheetNames[i];
      });
      const modules = analyze(extractModules(source), sheetNames, codeNames);
      return { modules, available: true };
    } catch (e) {
      console.warn('No se pudieron leer las macros', e);
      return { modules: [], available: true, error: e.message };
    }
  }

  global.VBA = { readMacros, decompress, analyze };
})(window);

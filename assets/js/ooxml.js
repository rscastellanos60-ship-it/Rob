/*
 * ooxml.js — Lee las partes del archivo Excel que SheetJS no expone:
 * gráficos (c:chart y cx:chart), cuadros de texto, imágenes, botones de
 * formulario (VML) y los colores del tema. Todo se devuelve como objetos
 * planos para que app.js los dibuje con Chart.js.
 */
(function (global) {
  'use strict';

  /* ---------- utilidades XML ---------- */
  const parseXml = (text) => new DOMParser().parseFromString(text, 'application/xml');
  const all = (node, name) => (node ? Array.from(node.getElementsByTagNameNS('*', name)) : []);
  const first = (node, name) => (node ? node.getElementsByTagNameNS('*', name)[0] || null : null);
  const kids = (node, name) => (node ? Array.from(node.children).filter((c) => !name || c.localName === name) : []);
  const kid = (node, name) => kids(node, name)[0] || null;
  const val = (node, name) => {
    const c = kid(node, name);
    return c ? c.getAttribute('val') : null;
  };
  const relAttr = (el, local) => {
    if (!el) return null;
    for (const a of el.attributes) {
      if (a.localName === local && a.namespaceURI && a.namespaceURI.includes('relationships')) return a.value;
    }
    return null;
  };

  function resolvePath(base, target) {
    if (!target) return null;
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/');
    parts.pop();
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.' && seg !== '') parts.push(seg);
    }
    return parts.join('/');
  }

  function relsPath(p) {
    const i = p.lastIndexOf('/');
    return (i < 0 ? '' : p.slice(0, i + 1)) + '_rels/' + p.slice(i + 1) + '.rels';
  }

  async function readText(zip, path) {
    const f = path && zip.file(path);
    return f ? f.async('text') : null;
  }

  async function readRels(zip, partPath) {
    const text = await readText(zip, relsPath(partPath));
    const map = {};
    if (!text) return map;
    for (const r of all(parseXml(text), 'Relationship')) {
      const external = r.getAttribute('TargetMode') === 'External';
      map[r.getAttribute('Id')] = {
        type: (r.getAttribute('Type') || '').split('/').pop(),
        target: external ? null : resolvePath(partPath, r.getAttribute('Target')),
      };
    }
    return map;
  }

  /* ---------- colores ---------- */
  function hexToRgb(hex) {
    const n = parseInt(hex.replace('#', ''), 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  function rgbToHex([r, g, b]) {
    return '#' + [r, g, b].map((x) => Math.max(0, Math.min(255, Math.round(x))).toString(16).padStart(2, '0')).join('');
  }
  function rgbToHsl([r, g, b]) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h /= 6;
    }
    return [h, s, l];
  }
  function hslToRgb([h, s, l]) {
    if (s === 0) return [l * 255, l * 255, l * 255];
    const hue = (p, q, t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
  }

  // Convierte un nodo a:srgbClr / a:schemeClr / a:sysClr (con lumMod/lumOff) en hex.
  function colorOf(clrEl, theme) {
    if (!clrEl) return null;
    let hex = null;
    if (clrEl.localName === 'srgbClr') hex = '#' + clrEl.getAttribute('val');
    else if (clrEl.localName === 'sysClr') hex = '#' + (clrEl.getAttribute('lastClr') || '000000');
    else if (clrEl.localName === 'schemeClr') {
      const key = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }[clrEl.getAttribute('val')] || clrEl.getAttribute('val');
      hex = theme && theme.colors[key];
    }
    if (!hex) return null;
    const mod = val(clrEl, 'lumMod'), off = val(clrEl, 'lumOff'), shade = val(clrEl, 'shade'), tint = val(clrEl, 'tint');
    if (mod || off || shade || tint) {
      const hsl = rgbToHsl(hexToRgb(hex));
      if (mod) hsl[2] *= +mod / 100000;
      if (off) hsl[2] += +off / 100000;
      if (shade) hsl[2] *= +shade / 100000;
      if (tint) hsl[2] = hsl[2] + (1 - hsl[2]) * (1 - +tint / 100000);
      hsl[2] = Math.max(0, Math.min(1, hsl[2]));
      hex = rgbToHex(hslToRgb(hsl));
    }
    return hex;
  }

  function fillColor(spPr, theme) {
    if (!spPr) return null;
    const solid = kid(spPr, 'solidFill');
    if (solid) return colorOf(solid.firstElementChild, theme);
    const grad = kid(spPr, 'gradFill');
    if (grad) return colorOf(first(grad, 'gs') && first(grad, 'gs').firstElementChild, theme);
    const ln = kid(spPr, 'ln');
    const lnSolid = ln && kid(ln, 'solidFill');
    if (lnSolid) return colorOf(lnSolid.firstElementChild, theme);
    return null;
  }

  async function readTheme(zip, wbPath, wbRels) {
    const rel = Object.values(wbRels).find((r) => r.type === 'theme');
    const text = rel && (await readText(zip, rel.target));
    const colors = {};
    if (text) {
      const scheme = first(parseXml(text), 'clrScheme');
      for (const c of kids(scheme)) {
        const inner = c.firstElementChild;
        if (!inner) continue;
        colors[c.localName] = '#' + (inner.getAttribute('val') && inner.localName === 'srgbClr' ? inner.getAttribute('val') : inner.getAttribute('lastClr') || '000000');
      }
    }
    return { colors };
  }

  /* ---------- texto enriquecido ---------- */
  function richText(node) {
    if (!node) return '';
    const paras = all(node, 'p');
    if (!paras.length) return all(node, 't').map((t) => t.textContent).join('').trim();
    return paras
      .map((p) => all(p, 't').map((t) => t.textContent).join(''))
      .filter((s) => s.trim())
      .join('\n')
      .trim();
  }

  /* ---------- anclajes ---------- */
  function anchorPos(a) {
    const cell = (el) => (el ? { col: +(first(el, 'col') || {}).textContent || 0, row: +(first(el, 'row') || {}).textContent || 0 } : null);
    const from = cell(kid(a, 'from'));
    let to = cell(kid(a, 'to'));
    if (!to) {
      const ext = kid(a, 'ext');
      const cx = ext ? +ext.getAttribute('cx') : 0, cy = ext ? +ext.getAttribute('cy') : 0;
      // ~ 64px por columna, 20px por fila; 9525 EMU por píxel
      const base = from || { col: 0, row: 0 };
      to = { col: base.col + Math.max(1, Math.round(cx / 9525 / 64)), row: base.row + Math.max(1, Math.round(cy / 9525 / 20)) };
    }
    return { from: from || { col: 0, row: 0 }, to };
  }

  /* ---------- gráficos clásicos (c:chart) ---------- */
  function readRefData(el) {
    if (!el) return null;
    const ref = el.firstElementChild;
    if (!ref) return null;
    const out = { f: null, cache: null, formatCode: null, multi: false };
    const fEl = kid(ref, 'f');
    if (fEl) out.f = fEl.textContent.trim();
    let cache = null;
    if (/Lit$/.test(ref.localName)) cache = ref;
    else cache = kid(ref, 'numCache') || kid(ref, 'strCache') || kid(ref, 'multiLvlStrCache');
    if (cache) {
      out.formatCode = (kid(cache, 'formatCode') || {}).textContent || null;
      const count = +(val(cache, 'ptCount') || 0);
      if (cache.localName === 'multiLvlStrCache') {
        out.multi = true;
        const lvls = kids(cache, 'lvl');
        const arr = new Array(count).fill('');
        // lvl[0] es el nivel interno (hoja); los externos se agregan como prefijo
        lvls.forEach((lvl, li) => {
          const filled = new Array(count).fill(null);
          for (const pt of kids(lvl, 'pt')) filled[+pt.getAttribute('idx')] = (kid(pt, 'v') || {}).textContent || '';
          let last = '';
          for (let i = 0; i < count; i++) {
            if (filled[i] != null) last = filled[i];
            const piece = li === 0 ? (filled[i] || '') : last;
            arr[i] = li === 0 ? piece : piece && arr[i] ? piece + ' · ' + arr[i] : arr[i] || piece;
          }
        });
        out.cache = arr;
      } else {
        const arr = new Array(count).fill(null);
        for (const pt of kids(cache, 'pt')) {
          const i = +pt.getAttribute('idx');
          if (i < count) arr[i] = (kid(pt, 'v') || {}).textContent;
        }
        out.cache = arr;
        out.numeric = cache.localName.startsWith('num');
      }
    }
    return out;
  }

  function readSeries(s, theme) {
    const tx = kid(s, 'tx');
    const name = { f: null, text: null };
    if (tx) {
      const fEl = first(tx, 'f');
      if (fEl) name.f = fEl.textContent.trim();
      const v = first(tx, 'v');
      if (v) name.text = v.textContent;
    }
    const dPt = {};
    for (const p of kids(s, 'dPt')) {
      const c = fillColor(kid(p, 'spPr'), theme);
      if (c) dPt[+val(p, 'idx')] = c;
    }
    const dl = kid(s, 'dLbls');
    return {
      idx: +(val(s, 'idx') || 0),
      order: +(val(s, 'order') || 0),
      name,
      color: fillColor(kid(s, 'spPr'), theme),
      dPt,
      cat: readRefData(kid(s, 'cat') || kid(s, 'xVal')),
      val: readRefData(kid(s, 'val') || kid(s, 'yVal')),
      size: readRefData(kid(s, 'bubbleSize')),
      smooth: val(s, 'smooth') === '1',
      noMarker: (() => {
        const m = kid(s, 'marker');
        return !!m && val(m, 'symbol') === 'none';
      })(),
      labels: dl ? { showVal: val(dl, 'showVal') === '1', showPercent: val(dl, 'showPercent') === '1' } : null,
      trend: !!kid(s, 'trendline'),
    };
  }

  function readAxis(ax) {
    const nf = kid(ax, 'numFmt');
    const scaling = kid(ax, 'scaling');
    return {
      id: val(ax, 'axId'),
      kind: ax.localName,
      pos: val(ax, 'axPos'),
      deleted: val(ax, 'delete') === '1',
      format: nf && nf.getAttribute('sourceLinked') !== '1' ? nf.getAttribute('formatCode') : null,
      title: kid(ax, 'title') ? richText(kid(ax, 'title')) : '',
      min: scaling && val(scaling, 'min') != null ? +val(scaling, 'min') : null,
      max: scaling && val(scaling, 'max') != null ? +val(scaling, 'max') : null,
      reversed: scaling && val(scaling, 'orientation') === 'maxMin',
      grid: !!kid(ax, 'majorGridlines'),
    };
  }

  function readTitle(t) {
    if (!t) return { text: '', f: null };
    const fEl = first(t, 'f');
    return { text: richText(first(t, 'rich') || t), f: fEl ? fEl.textContent.trim() : null };
  }

  async function parseChart(zip, path, theme) {
    const text = await readText(zip, path);
    if (!text) return null;
    const doc = parseXml(text);
    const chartEl = kid(doc.documentElement, 'chart');
    if (!chartEl) return null;
    const plot = kid(chartEl, 'plotArea');
    const groups = [];
    for (const g of kids(plot)) {
      if (!/Chart$/.test(g.localName)) continue;
      const dl = kid(g, 'dLbls');
      groups.push({
        type: g.localName,
        barDir: val(g, 'barDir'),
        grouping: val(g, 'grouping'),
        varyColors: val(g, 'varyColors') === '1',
        holeSize: +(val(g, 'holeSize') || 50),
        scatterStyle: val(g, 'scatterStyle'),
        radarStyle: val(g, 'radarStyle'),
        axIds: kids(g, 'axId').map((a) => a.getAttribute('val')),
        series: kids(g, 'ser').map((s) => readSeries(s, theme)),
        labels: dl ? { showVal: val(dl, 'showVal') === '1', showPercent: val(dl, 'showPercent') === '1' } : null,
      });
    }
    const axes = kids(plot).filter((a) => /Ax$/.test(a.localName)).map(readAxis);
    const legend = kid(chartEl, 'legend');
    const pivot = !!kid(doc.documentElement, 'pivotSource');
    return {
      kind: 'chart',
      path,
      title: readTitle(kid(chartEl, 'title')),
      autoTitleDeleted: val(chartEl, 'autoTitleDeleted') === '1',
      groups,
      axes,
      legend: legend ? { pos: val(legend, 'legendPos') || 'r' } : null,
      pivot,
    };
  }

  /* ---------- gráficos nuevos de Excel 2016 (cx:chart) ---------- */
  async function parseChartEx(zip, path) {
    const text = await readText(zip, path);
    if (!text) return null;
    const doc = parseXml(text);
    const data = {};
    for (const d of all(doc, 'data')) {
      const dims = {};
      for (const dim of kids(d).filter((x) => /Dim$/.test(x.localName))) {
        const lvl = kid(dim, 'lvl');
        const count = lvl ? +(lvl.getAttribute('ptCount') || 0) : 0;
        const arr = new Array(count).fill(null);
        if (lvl) for (const pt of kids(lvl, 'pt')) arr[+pt.getAttribute('idx')] = pt.textContent;
        const fEl = kid(dim, 'f');
        dims[dim.getAttribute('type')] = {
          f: fEl ? fEl.textContent.trim() : null,
          cache: arr,
          formatCode: lvl && lvl.getAttribute('formatCode'),
          numeric: dim.localName === 'numDim',
        };
      }
      data[d.getAttribute('id')] = dims;
    }
    const chartEl = first(doc, 'chart');
    const series = all(doc, 'series').map((s) => {
      const id = val(s, 'dataId');
      const dims = data[id] || {};
      const tx = kid(s, 'tx');
      return {
        layout: s.getAttribute('layoutId'),
        hidden: s.getAttribute('hidden') === '1',
        name: { f: tx && first(tx, 'f') ? first(tx, 'f').textContent : null, text: tx && first(tx, 'v') ? first(tx, 'v').textContent : null },
        cat: dims.cat || null,
        val: dims.val || dims.size || null,
        color: null,
        dPt: {},
      };
    });
    const titleEl = chartEl && kid(chartEl, 'title');
    return {
      kind: 'chartex',
      path,
      title: { text: titleEl ? richText(titleEl) : '', f: titleEl && first(titleEl, 'f') ? first(titleEl, 'f').textContent : null },
      layout: series[0] ? series[0].layout : 'clusteredColumn',
      groups: [{ type: 'chartEx', series: series.filter((s) => !s.hidden && s.layout !== 'paretoLine') }],
      axes: [],
      legend: null,
    };
  }

  /* ---------- dibujos (drawingN.xml) ---------- */
  async function parseDrawing(zip, path, theme) {
    const text = await readText(zip, path);
    if (!text) return [];
    const doc = parseXml(text);
    const rels = await readRels(zip, path);
    const items = [];

    const handle = async (el, pos, groupName) => {
      // AlternateContent: usar la primera opción (Choice) que Excel moderno entiende
      if (el.localName === 'AlternateContent') {
        const choice = kid(el, 'Choice') || kid(el, 'Fallback');
        for (const c of kids(choice)) await handle(c, pos, groupName);
        return;
      }
      if (el.localName === 'grpSp') {
        for (const c of kids(el)) await handle(c, pos, groupName || (first(el, 'cNvPr') || { getAttribute: () => '' }).getAttribute('name'));
        return;
      }
      const cNvPr = first(el, 'cNvPr');
      const name = cNvPr ? cNvPr.getAttribute('name') || '' : '';
      const descr = cNvPr ? cNvPr.getAttribute('descr') || '' : '';
      const hidden = cNvPr && cNvPr.getAttribute('hidden') === '1';
      if (hidden) return;
      if (el.localName === 'graphicFrame') {
        const gd = first(el, 'graphicData');
        const chartRef = gd && kids(gd).find((c) => c.localName === 'chart');
        const id = relAttr(chartRef, 'id');
        const rel = id && rels[id];
        if (!rel || !rel.target) return;
        const isEx = (chartRef.namespaceURI || '').includes('chartex') || /chartEx/i.test(rel.target);
        const chart = isEx ? await parseChartEx(zip, rel.target) : await parseChart(zip, rel.target, theme);
        if (chart) items.push({ type: 'chart', name, pos, chart });
      } else if (el.localName === 'sp') {
        const txt = richText(kid(el, 'txBody'));
        const textlink = el.getAttribute('textlink') || null;
        const macro = el.getAttribute('macro') || null;
        if (!txt && !textlink && !macro) return;
        items.push({
          type: 'shape',
          name,
          descr,
          pos,
          text: txt,
          textlink,
          macro,
          fill: fillColor(kid(el, 'spPr'), theme),
          group: groupName || null,
        });
      } else if (el.localName === 'pic') {
        const blip = first(el, 'blip');
        const id = relAttr(blip, 'embed');
        const rel = id && rels[id];
        if (!rel || !rel.target || !zip.file(rel.target)) return;
        const ext = rel.target.split('.').pop().toLowerCase();
        const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp', webp: 'image/webp' }[ext];
        if (!mime) return;
        const b64 = await zip.file(rel.target).async('base64');
        items.push({ type: 'image', name, descr, pos, src: `data:${mime};base64,${b64}`, macro: el.getAttribute('macro') || null });
      }
    };

    for (const a of kids(doc.documentElement)) {
      if (!/Anchor$/.test(a.localName)) continue;
      const pos = anchorPos(a);
      for (const c of kids(a)) {
        if (['from', 'to', 'ext', 'pos', 'clientData'].includes(c.localName)) continue;
        await handle(c, pos, null);
      }
    }
    return items;
  }

  /* ---------- controles de formulario (VML heredado) ---------- */
  function decodeEntities(s) {
    return s
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n))
      .replace(/&amp;/g, '&')
      .trim();
  }

  async function parseVml(zip, path) {
    const text = await readText(zip, path);
    if (!text) return [];
    const out = [];
    const shapeRe = /<v:shape\b[\s\S]*?<\/v:shape>/gi;
    let m;
    while ((m = shapeRe.exec(text))) {
      const s = m[0];
      const type = (s.match(/ObjectType="([^"]+)"/i) || [])[1];
      if (!type || type === 'Note') continue;
      const tag = (t) => {
        const r = s.match(new RegExp(`<x:${t}>([\\s\\S]*?)</x:${t}>`, 'i'));
        return r ? decodeEntities(r[1]) : null;
      };
      const anchor = (tag('Anchor') || '').split(',').map((x) => +x.trim());
      const tb = s.match(/<v:textbox[\s\S]*?>([\s\S]*?)<\/v:textbox>/i);
      out.push({
        type: 'control',
        control: type, // Button, Drop, Checkbox, Spin, List, Radio...
        text: tb ? decodeEntities(tb[1]) : '',
        macro: tag('FmlaMacro'),
        link: tag('FmlaLink'),
        range: tag('FmlaRange'),
        pos: anchor.length >= 8 ? { from: { col: anchor[0], row: anchor[2] }, to: { col: anchor[4], row: anchor[6] } } : { from: { col: 0, row: 0 }, to: { col: 2, row: 2 } },
      });
    }
    return out;
  }

  /* ---------- punto de entrada ---------- */
  async function readWorkbookParts(buffer) {
    if (typeof JSZip === 'undefined') return null;
    let zip;
    try {
      zip = await JSZip.loadAsync(buffer);
    } catch (e) {
      return null; // .xls binario u otro formato: no hay partes OOXML
    }
    const rootRels = await readRels(zip, '');
    const wbRel = Object.values(rootRels).find((r) => r.type === 'officeDocument');
    const wbPath = wbRel ? wbRel.target : 'xl/workbook.xml';
    const wbText = await readText(zip, wbPath);
    if (!wbText || !wbPath.endsWith('.xml')) return { sheets: {}, theme: { colors: {} } };
    const wbRels = await readRels(zip, wbPath);
    const theme = await readTheme(zip, wbPath, wbRels);
    const sheets = {};
    for (const s of all(parseXml(wbText), 'sheet')) {
      const name = s.getAttribute('name');
      const rel = wbRels[relAttr(s, 'id')];
      const entry = { name, state: s.getAttribute('state') || 'visible', kind: rel ? rel.type : 'worksheet', items: [] };
      sheets[name] = entry;
      if (!rel || !rel.target) continue;
      const sRels = await readRels(zip, rel.target);
      for (const r of Object.values(sRels)) {
        if (!r.target) continue;
        if (r.type === 'drawing') entry.items.push(...(await parseDrawing(zip, r.target, theme)));
        else if (r.type === 'vmlDrawing') entry.items.push(...(await parseVml(zip, r.target)));
      }
    }
    return { sheets, theme };
  }

  global.OOXML = { readWorkbookParts, colorOf, hexToRgb, rgbToHex };
})(window);

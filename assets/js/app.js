/*
 * app.js — Tablero de indicadores a partir de un archivo Excel.
 *
 * Flujo: se carga el Excel (subido en el navegador o publicado en /data),
 * se leen datos, gráficos, cuadros de texto, botones y macros, y se arma una
 * página con una sección por hoja + resumen (hoja Dashboard) + macros.
 */
(function () {
  'use strict';

  const { Analysis: A, OOXML, VBA } = window;
  const $ = (sel, root = document) => root.querySelector(sel);
  const el = (tag, attrs = {}, ...children) => {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === 'class') node.className = v;
      else if (k === 'html') node.innerHTML = v;
      else if (k === 'style' && typeof v === 'object') Object.assign(node.style, v);
      else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
      else node.setAttribute(k, v === true ? '' : v);
    }
    for (const c of children.flat()) if (c != null && c !== false) node.append(c.nodeType ? c : document.createTextNode(String(c)));
    return node;
  };
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const slug = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'seccion';

  const ICONS = {
    dashboard: '<path d="M3 13h8V3H3v10zm0 8h8v-6H3v6zm10 0h8V11h-8v10zm0-18v6h8V3h-8z"/>',
    sheet: '<path d="M4 3h16a1 1 0 0 1 1 1v16a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1zm1 5v4h6V8H5zm8 0v4h6V8h-6zm-8 6v4h6v-4H5zm8 0v4h6v-4h-6z"/>',
    chart: '<path d="M5 9h3v11H5zM10.5 4h3v16h-3zM16 13h3v7h-3z"/>',
    macro: '<path d="M8.6 16.6 4 12l4.6-4.6L10 8.8 6.8 12l3.2 3.2-1.4 1.4zm6.8 0L14 15.2l3.2-3.2L14 8.8l1.4-1.4L20 12l-4.6 4.6z"/>',
    file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zm-1 7V3.5L18.5 9H13z"/>',
    hidden: '<path d="M12 6a9.8 9.8 0 0 1 9 6 9.9 9.9 0 0 1-2.4 3.4l-1.4-1.4A8 8 0 0 0 18.8 12 7.8 7.8 0 0 0 12 8c-.6 0-1.2.1-1.8.2L8.6 6.6A10 10 0 0 1 12 6zM3.3 3.3 2 4.6l2.7 2.7A10 10 0 0 0 3 12a9.8 9.8 0 0 0 12.4 5.7l3 3 1.3-1.3L3.3 3.3zM12 16a4 4 0 0 1-4-4c0-.6.1-1.1.4-1.6l5.2 5.2c-.5.3-1 .4-1.6.4z"/>',
    upload: '<path d="M5 20h14v-2H5v2zM12 3 6.5 8.5l1.4 1.4L11 6.8V16h2V6.8l3.1 3.1 1.4-1.4L12 3z"/>',
    print: '<path d="M7 3h10v4H7zM5 8h14a2 2 0 0 1 2 2v6h-4v4H7v-4H3v-6a2 2 0 0 1 2-2zm4 7v3h6v-3H9z"/>',
    moon: '<path d="M12.4 3a7 7 0 1 0 8.6 8.6A8.5 8.5 0 1 1 12.4 3z"/>',
    sun: '<path d="M12 7a5 5 0 1 0 0 10 5 5 0 0 0 0-10zM11 1h2v3h-2zm0 19h2v3h-2zM3.5 4.9l1.4-1.4 2.1 2.1-1.4 1.4zm12.9 12.9 1.4-1.4 2.1 2.1-1.4 1.4zM1 11h3v2H1zm19 0h3v2h-3zM4.9 20.5l-1.4-1.4 2.1-2.1 1.4 1.4zM17.8 7.6l-1.4-1.4 2.1-2.1 1.4 1.4z"/>',
    table: '<path d="M3 4h18v16H3V4zm2 4v4h5V8H5zm7 0v4h7V8h-7zm-7 6v4h5v-4H5zm7 0v4h7v-4h-7z"/>',
    download: '<path d="M5 20h14v-2H5v2zm7-3 5.5-5.5-1.4-1.4L13 13.2V4h-2v9.2L7.9 10.1 6.5 11.5 12 17z"/>',
    image: '<path d="M21 19V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2zM8.5 13.5l2.5 3 3.5-4.5 4.5 6H5l3.5-4.5z"/>',
    palette: '<path d="M12 3a9 9 0 0 0 0 18c.8 0 1.5-.7 1.5-1.5 0-.4-.1-.7-.4-1-.2-.3-.4-.6-.4-1 0-.8.7-1.5 1.5-1.5H16a5 5 0 0 0 5-5c0-4.4-4-8-9-8zm-5.5 9a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3zm3-4a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3zm5 0a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3zm3 4a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3z"/>',
    chev: '<path d="M9 6l6 6-6 6" fill="none" stroke="currentColor" stroke-width="2"/>',
    leaf: '<path d="M12 21c0-5 .5-9 3-12" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/><path d="M12 12c-3.5-1-6.5-.2-9 2 3 .8 6.2.6 9-2zm0-2.5C10.6 6 8 4 4.5 3.5c1 3.4 3.8 5.6 7.5 6zm2.6-1C15.6 5.2 18 3.4 21.5 3c-.7 3.5-3.1 5.8-6.9 5.5zM15 9c3.4-.6 6 .6 7.5 3-3.1.7-5.8-.3-7.5-3z" fill="currentColor"/>',
  };
  Object.assign(ICONS, {
    woman: '<circle cx="12" cy="3.6" r="2.3"/><path d="M10.2 7.2h3.6c.63 0 1.18.42 1.35 1.03L17.2 15h-2.7v7h-1.7v-7h-1.6v7H9.5v-7H6.8l2.05-6.77c.17-.61.72-1.03 1.35-1.03z"/>',
    man: '<circle cx="12" cy="3.6" r="2.3"/><path d="M9.6 7.2h4.8c.88 0 1.6.72 1.6 1.6V15h-1.9v7h-1.65v-6.2h-.9V22H9.9v-7H8V8.8c0-.88.72-1.6 1.6-1.6z"/>',
    people: '<circle cx="8" cy="5" r="2.3"/><circle cx="16.5" cy="6" r="2"/><path d="M5 9h6a2 2 0 0 1 2 2v4h-1.6v7H8.9v-5H7.1v5H4.6v-7H3v-4a2 2 0 0 1 2-2zm9.6 1.2h3.8a1.8 1.8 0 0 1 1.8 1.8v3.6h-1.4V22h-2.2v-5.4h-.8V22h-1.6v-6.4H14v-4h.6z"/>',
    org: '<path d="M9.5 2h5v4.5h-1.6V9H19v3.5h1.5V17h-4.5v-4.5H17V11h-4.1v1.5h1.6V17h-5v-4.5h1.6V11H7v1.5h1.5V17H4v-4.5h1V9h6.1V6.5H9.5z"/>',
    cycle: '<path d="M12 4V1L8 5l4 4V6a6 6 0 0 1 6 6c0 1-.25 1.95-.7 2.78l1.46 1.46A8 8 0 0 0 20 12a8 8 0 0 0-8-8zm-6 8c0-1 .25-1.95.7-2.78L5.24 7.76A8 8 0 0 0 4 12a8 8 0 0 0 8 8v3l4-4-4-4v3a6 6 0 0 1-6-6z"/>',
    exit: '<path d="M4 3h10v2H6v14h8v2H4zm12.2 4.2L21 12l-4.8 4.8-1.4-1.4 2.4-2.4H9v-2h8.2l-2.4-2.4z"/>',
    usercheck: '<circle cx="9" cy="7" r="3.5"/><path d="M2 20c0-3.6 3.1-6 7-6s7 2.4 7 6v1H2zm14.5-9.6 1.4-1.4 1.6 1.6 3.1-3.1 1.4 1.4-4.5 4.5z"/>',
    target: '<path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm0 18a8 8 0 1 1 8-8 8 8 0 0 1-8 8zm0-14a6 6 0 1 0 6 6 6 6 0 0 0-6-6zm0 10a4 4 0 1 1 4-4 4 4 0 0 1-4 4zm0-6a2 2 0 1 0 2 2 2 2 0 0 0-2-2z"/>',
    hourglass: '<path d="M6 2h12v2h-1v3.2c0 .8-.32 1.56-.88 2.12L13.4 12l2.72 2.68c.56.56.88 1.32.88 2.12V20h1v2H6v-2h1v-3.2c0-.8.32-1.56.88-2.12L10.6 12 7.88 9.32A3 3 0 0 1 7 7.2V4H6zm3 2v3.2l3 3 3-3V4z"/>',
    beach: '<path d="M12 2a9 9 0 0 1 9 9h-8v9h3v2H8v-2h3v-9H3a9 9 0 0 1 9-9z"/>',
    calendarx: '<path d="M7 2h2v2h6V2h2v2h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2zM5 9v11h14V9zm4.4 2.6L12 14.2l2.6-2.6 1.4 1.4-2.6 2.6 2.6 2.6-1.4 1.4-2.6-2.6-2.6 2.6L8 17.6l2.6-2.6L8 12.4z"/>',
    balance: '<path d="M11 3h2v2.1l5.6 1.4L21.5 13A3.5 3.5 0 0 1 15 13l2.6-5.2L13 6.7V19h4v2H7v-2h4V6.7L6.4 7.8 9 13a3.5 3.5 0 0 1-6.5 0l2.9-6.5L11 5.1zm-5.5 6.3L4 12.5h3zm13 0L17 12.5h3z"/>',
    home: '<path d="M12 3 2 11h3v10h5v-6h4v6h5V11h3z"/>',
    grad: '<path d="M12 3 1 9l11 6 9-4.9V17h2V9zM5 13.2v4L12 21l7-3.8v-4L12 17z"/>',
    shield: '<path d="M12 2 4 5v6c0 5 3.4 9.7 8 11 4.6-1.3 8-6 8-11V5zm-1.2 14.2-3.5-3.5 1.4-1.4 2.1 2.1 4.8-4.8 1.4 1.4z"/>',
    clipboard: '<path d="M9 2h6v2h3a1 1 0 0 1 1 1v16a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h3zm0 4v1.5h6V6zm-1 5v2h8v-2zm0 4v2h6v-2z"/>',
    clock: '<path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm0 18a8 8 0 1 1 8-8 8 8 0 0 1-8 8zm1-13h-2v6l5 3 1-1.6-4-2.4z"/>',
    map: '<path d="M12 2a7 7 0 0 1 7 7c0 5.2-7 13-7 13S5 14.2 5 9a7 7 0 0 1 7-7zm0 4.5A2.5 2.5 0 1 0 14.5 9 2.5 2.5 0 0 0 12 6.5z"/>',
    heart: '<path d="M12 21 10.6 19.7C5.4 15 2 11.9 2 8.1A5 5 0 0 1 7.1 3 5.6 5.6 0 0 1 12 5.5 5.6 5.6 0 0 1 16.9 3 5 5 0 0 1 22 8.1c0 3.8-3.4 6.9-8.6 11.6z"/>',
    briefcase: '<path d="M9 3h6a2 2 0 0 1 2 2v2h3a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V9a2 2 0 0 1 2-2h3V5a2 2 0 0 1 2-2zm0 4h6V5H9z"/>',
  });

  // Ícono según el tema del texto (secciones, indicadores, gráficos)
  const TOPICS = [
    [/g[eé]nero|sexo|mujer|femenin/i, 'woman'],
    [/estructura|cargos|ocupaci[oó]n|hc\b|planta/i, 'org'],
    [/mapa social|generaci[oó]n|social/i, 'people'],
    [/rotaci[oó]n/i, 'cycle'],
    [/control/i, 'clipboard'],
    [/retiro/i, 'exit'],
    [/cobertura/i, 'target'],
    [/selecci[oó]n|\bans\b|vacante/i, 'usercheck'],
    [/pensi[oó]n|prepension|antig[uü]edad/i, 'hourglass'],
    [/vacacion/i, 'beach'],
    [/ausent/i, 'calendarx'],
    [/equidad/i, 'balance'],
    [/teletrabajo|remoto/i, 'home'],
    [/desarrollo|inducci[oó]n|entrenamiento|formaci[oó]n|desempe[nñ]o|capacitaci/i, 'grad'],
    [/seguridad|salud|sst|accident/i, 'shield'],
    [/geogr|zona|sede/i, 'map'],
    [/estado civil|engagement|bienestar|enps/i, 'heart'],
    [/contrataci[oó]n|salario|contrato/i, 'briefcase'],
  ];
  const topicIcon = (text, fallback) => {
    const hit = TOPICS.find(([re]) => re.test(String(text || '')));
    return hit ? hit[1] : fallback;
  };

  const icon = (name, size = 18) => {
    const span = document.createElement('span');
    span.style.display = 'inline-flex';
    span.innerHTML = `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">${ICONS[name] || ''}</svg>`;
    return span.firstChild;
  };

  /* ---------- preferencias (por navegador) ---------- */
  const prefs = {
    get(k, d) {
      try {
        const v = localStorage.getItem('tablero:' + k);
        return v == null ? d : v;
      } catch (e) {
        return d;
      }
    },
    set(k, v) {
      try {
        localStorage.setItem('tablero:' + k, v);
      } catch (e) {
        /* almacenamiento bloqueado */
      }
    },
  };

  /* ---------- almacenamiento del último archivo (IndexedDB) ---------- */
  const store = {
    db: null,
    async open() {
      if (this.db) return this.db;
      this.db = await new Promise((resolve, reject) => {
        const req = indexedDB.open('tablero-excel', 1);
        req.onupgradeneeded = () => req.result.createObjectStore('files');
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
      return this.db;
    },
    async get() {
      try {
        const db = await this.open();
        return await new Promise((resolve) => {
          const req = db.transaction('files').objectStore('files').get('actual');
          req.onsuccess = () => resolve(req.result || null);
          req.onerror = () => resolve(null);
        });
      } catch (e) {
        return null;
      }
    },
    async put(rec) {
      try {
        const db = await this.open();
        await new Promise((resolve) => {
          const tx = db.transaction('files', 'readwrite');
          tx.objectStore('files').put(rec, 'actual');
          tx.oncomplete = resolve;
          tx.onerror = resolve;
        });
      } catch (e) {
        /* sin persistencia: la página sigue funcionando */
      }
    },
    async clear() {
      try {
        const db = await this.open();
        await new Promise((resolve) => {
          const tx = db.transaction('files', 'readwrite');
          tx.objectStore('files').delete('actual');
          tx.oncomplete = resolve;
          tx.onerror = resolve;
        });
      } catch (e) {
        /* nada que borrar */
      }
    },
  };

  /* ---------- estado ---------- */
  const state = {
    // seccionesTorta: secciones del Dashboard cuyos gráficos de categorías se muestran como torta
    config: { organizacion: '', titulo: 'Tablero de indicadores', archivo: null, seccionesTorta: ['Mapa Social'] },
    wb: null,
    parts: null,
    sheets: {},
    macros: { modules: [] },
    sections: [],
    file: null,
    repoFile: null,
    charts: new Map(), // id -> { chart, build }
    rendered: new Set(),
  };

  /* ---------- tema y paleta ---------- */
  function applyTheme(mode) {
    if (mode === 'light' || mode === 'dark') document.documentElement.setAttribute('data-theme', mode);
    else document.documentElement.removeAttribute('data-theme');
    prefs.set('tema', mode);
  }
  function isDark() {
    const t = document.documentElement.getAttribute('data-theme');
    return t ? t === 'dark' : window.matchMedia('(prefers-color-scheme: dark)').matches;
  }
  function tokens() {
    const cs = getComputedStyle(document.documentElement);
    const v = (n) => cs.getPropertyValue(n).trim();
    return {
      text: v('--text'),
      text2: v('--text-2'),
      muted: v('--muted'),
      grid: v('--grid'),
      surface: v('--surface'),
      series: [1, 2, 3, 4, 5, 6].map((i) => v('--series-' + i)),
      other: v('--series-other'),
      font: v('--font') || 'system-ui',
    };
  }
  const useExcelColors = () => prefs.get('colores', 'verde') === 'excel';
  function seriesColor(i, t, excelColor) {
    if (useExcelColors() && excelColor) return excelColor;
    return i < t.series.length ? t.series[i] : t.other;
  }
  function alpha(hex, a) {
    if (!hex || !hex.startsWith('#')) return hex;
    const [r, g, b] = OOXML.hexToRgb(hex);
    return `rgba(${r},${g},${b},${a})`;
  }

  /* ---------- formato de números ---------- */
  function fmtValue(v, fmt) {
    if (v == null || Number.isNaN(v)) return '—';
    return A.formatNumber(v, fmt);
  }
  function fmtTick(v, fmt) {
    if (typeof v !== 'number') return v;
    if (fmt && /%/.test(fmt)) return A.formatNumber(v, fmt.replace(/\.0+/, ''));
    const abs = Math.abs(v);
    if (abs >= 10000) return new Intl.NumberFormat('es-CO', { notation: 'compact', maximumFractionDigits: 1 }).format(v);
    return new Intl.NumberFormat('es-CO', { maximumFractionDigits: abs < 10 ? 2 : 0 }).format(v);
  }

  /* =======================================================================
     Gráficos (Chart.js)
     ======================================================================= */

  // Dibuja el valor sobre las barras / porcentaje en tortas cuando Excel lo tenía activado
  const valueLabels = {
    id: 'valueLabels',
    afterDatasetsDraw(chart, _args, opts) {
      if (!opts || !opts.enabled) return;
      const { ctx } = chart;
      const t = tokens();
      ctx.save();
      ctx.font = `600 11px ${t.font}`;
      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        const total = ds.data.reduce((s, x) => s + (typeof x === 'number' ? x : 0), 0);
        meta.data.forEach((elm, i) => {
          let raw = ds.data[i];
          if (Array.isArray(raw)) {
            const shown = ds._values ? ds._values[i] : raw[1] - raw[0];
            const p = elm.getProps(['x', 'y', 'base'], true);
            const text = fmtTick(shown, ds._fmt);
            const horizontal = chart.options.indexAxis === 'y';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            if (opts.inside && horizontal) {
              ctx.fillStyle = '#fff';
              ctx.fillText(text, (p.x + p.base) / 2, p.y);
            } else {
              ctx.fillStyle = t.text2;
              ctx.textBaseline = 'bottom';
              ctx.fillText(text, p.x, Math.min(p.y, p.base) - 4);
            }
            return;
          }
          if (raw == null || typeof raw !== 'number') return;
          let text;
          if (meta.type === 'pie' || meta.type === 'doughnut') {
            if (!total || raw / total < 0.07) return;
            text = opts.percent ? Math.round((raw / total) * 100) + '%' : fmtTick(raw, ds._fmt);
            const p = elm.tooltipPosition();
            ctx.fillStyle = '#fff';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(text, p.x, p.y);
            return;
          }
          text = fmtTick(raw, ds._fmt);
          const horizontal = chart.options.indexAxis === 'y';
          ctx.fillStyle = t.text2;
          ctx.textAlign = horizontal ? 'left' : 'center';
          ctx.textBaseline = horizontal ? 'middle' : 'bottom';
          const { x, y } = elm.getProps(['x', 'y'], true);
          ctx.fillText(text, horizontal ? x + 4 : x, horizontal ? y : y - 4);
        });
      });
      ctx.restore();
    },
  };

  // Total escrito en el centro de los anillos (no se grafica como porción)
  const centerTotal = {
    id: 'centerTotal',
    afterDraw(chart, _args, opts) {
      if (!opts || !opts.text) return;
      const arc = chart.getDatasetMeta(0).data[0];
      if (!arc) return;
      const { x, y, innerRadius } = arc.getProps(['x', 'y', 'innerRadius'], true);
      if (!innerRadius || innerRadius < 22) return;
      const t = tokens();
      const ctx = chart.ctx;
      ctx.save();
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      const size = Math.max(14, Math.min(30, innerRadius * 0.5));
      ctx.fillStyle = t.text;
      ctx.font = `750 ${size}px ${t.font}`;
      ctx.fillText(opts.text, x, y - size * 0.25);
      ctx.fillStyle = t.muted;
      ctx.font = `600 ${Math.max(10, size * 0.42)}px ${t.font}`;
      ctx.fillText(opts.label, x, y + size * 0.55);
      ctx.restore();
    },
  };

  // Línea vertical de referencia al pasar el cursor
  const crosshair = {
    id: 'crosshair',
    afterDraw(chart) {
      const active = chart.tooltip && chart.tooltip.getActiveElements();
      if (!active || !active.length || chart.config.type !== 'line') return;
      const x = active[0].element.x;
      const { top, bottom } = chart.chartArea;
      const ctx = chart.ctx;
      ctx.save();
      ctx.strokeStyle = tokens().muted;
      ctx.globalAlpha = 0.5;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(x, top);
      ctx.lineTo(x, bottom);
      ctx.stroke();
      ctx.restore();
    },
  };

  /*
   * Modelo normalizado de gráfico:
   * { type: bar|line|area|pie|doughnut|scatter|bubble|radar|waterfall,
   *   labels: [], series: [{ name, values, color, fmt, kind?, points?, slices? }],
   *   horizontal, stacked, percent, cutout, showValues, showPercent, yFormat, yTitle, xTitle, min, max }
   */
  function buildConfig(model) {
    const t = tokens();
    const pie = model.type === 'pie' || model.type === 'doughnut';
    const many = model.labels.length > 24;
    let labels = model.labels.slice();
    let series = model.series;

    if (pie) {
      // Torta: una sola serie; más de 6 porciones se agrupan en "Otros"
      const s = series[0] || { values: [] };
      let pairs = labels.map((l, i) => ({ l, v: s.values[i], c: s.slices && s.slices[i] })).filter((p) => typeof p.v === 'number' && p.v > 0);
      if (pairs.length > 7) {
        const sorted = pairs.slice().sort((a, b) => b.v - a.v);
        const keep = sorted.slice(0, 6);
        const rest = sorted.slice(6).reduce((sum, p) => sum + p.v, 0);
        pairs = keep.concat([{ l: 'Otros', v: rest, other: true }]);
      }
      labels = pairs.map((p) => p.l);
      const ds = {
        label: s.name,
        data: pairs.map((p) => p.v),
        backgroundColor: pairs.map((p, i) => (p.other ? t.other : seriesColor(i, t, p.c))),
        borderColor: t.surface,
        borderWidth: 2,
        hoverOffset: 6,
        _fmt: s.fmt,
      };
      return {
        type: model.type,
        data: { labels, datasets: [ds] },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: model.type === 'doughnut' ? (model.cutout || 55) + '%' : 0,
          layout: { padding: 6 },
          // leyenda abajo cuando la tarjeta es angosta
          onResize: (chart, size) => {
            const pos = size.width < 380 ? 'bottom' : 'right';
            if (chart.options.plugins.legend.position !== pos) {
              chart.options.plugins.legend.position = pos;
              chart.update('none');
            }
          },
          plugins: {
            centerTotal: model.type === 'doughnut' && model.totalText ? { text: model.totalText, label: model.totalLabel || 'Total' } : null,
            legend: { position: 'right', labels: { color: t.text2, usePointStyle: true, pointStyle: 'rectRounded', boxWidth: 10, boxHeight: 10, padding: 12, font: { size: 12 } } },
            tooltip: {
              callbacks: {
                label: (ctx) => {
                  const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                  return ` ${ctx.label}: ${fmtValue(ctx.raw, s.fmt)} (${((ctx.raw / total) * 100).toFixed(1)}%)`;
                },
              },
            },
            valueLabels: { enabled: model.showValues || model.showPercent || pairs.length <= 6, percent: true },
          },
        },
        plugins: [valueLabels, centerTotal],
      };
    }

    if (model.type === 'funnel') {
      const s = series[0] || { values: [] };
      const vals = s.values.map((v) => (typeof v === 'number' ? v : 0));
      const max = Math.max(...vals.map(Math.abs), 1);
      return {
        type: 'bar',
        data: {
          labels,
          datasets: [{
            label: s.name,
            data: vals.map((v) => [(max - Math.abs(v)) / 2, (max + Math.abs(v)) / 2]),
            _values: vals,
            _fmt: s.fmt,
            backgroundColor: vals.map((_, i) => (useExcelColors() && s.slices && s.slices[i]) || t.series[0]),
            borderRadius: 4,
            borderSkipped: false,
            categoryPercentage: 0.86,
            barPercentage: 0.92,
          }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          indexAxis: 'y',
          scales: {
            x: { display: false, min: 0, max },
            y: { grid: { display: false }, border: { display: false }, afterFit: fitLabels(labels, t.font), ticks: { color: t.text2, font: { size: 11.5 }, callback: (v, i) => wrapLabel(labels[i], 20) } },
          },
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: (ctx) => ` ${ctx.dataset.label || ''}: ${fmtValue(vals[ctx.dataIndex], s.fmt)}` } },
            valueLabels: { enabled: true, inside: true },
          },
        },
        plugins: [valueLabels],
      };
    }

    if (model.percent) {
      const totals = labels.map((_, i) => series.reduce((s, x) => s + Math.abs(x.values[i] || 0), 0));
      series = series.map((s) => ({ ...s, values: s.values.map((v, i) => (v == null || !totals[i] ? null : v / totals[i])), fmt: '0%' }));
    }

    const isLineish = (k) => k === 'line' || k === 'area';
    const baseType = model.type === 'area' ? 'line' : model.type === 'waterfall' ? 'bar' : model.type;
    const datasets = series.map((s, i) => {
      const color = seriesColor(s.colorIndex ?? i, t, s.color);
      const kind = s.kind || model.type;
      const ds = {
        label: s.name,
        data: s.points || s.values,
        _fmt: s.fmt,
        borderColor: color,
        backgroundColor: color,
      };
      if (s.kind && s.kind !== model.type) ds.type = s.kind === 'area' ? 'line' : s.kind;
      if (isLineish(kind)) {
        Object.assign(ds, {
          borderWidth: 2,
          tension: s.smooth ? 0.35 : 0,
          pointRadius: many || s.noMarker ? 0 : 3,
          pointHoverRadius: 5,
          pointBackgroundColor: color,
          pointBorderColor: t.surface,
          pointBorderWidth: 2,
          spanGaps: true,
          fill: kind === 'area' ? (model.stacked && i > 0 ? '-1' : 'origin') : false,
          backgroundColor: kind === 'area' ? alpha(color, model.stacked ? 0.55 : 0.18) : color,
          order: 0,
        });
      } else if (kind === 'bar' || kind === 'waterfall') {
        const stackedTop = model.stacked && i === series.length - 1;
        Object.assign(ds, {
          borderRadius: model.stacked ? (stackedTop ? 4 : 0) : 4,
          borderSkipped: 'start',
          borderWidth: model.stacked ? { top: 2 } : 0,
          borderColor: model.stacked ? t.surface : color,
          maxBarThickness: 46,
          categoryPercentage: 0.72,
          barPercentage: series.length > 1 && !model.stacked ? 0.9 : 0.8,
          order: 1,
        });
      } else if (kind === 'scatter' || kind === 'bubble') {
        Object.assign(ds, { pointRadius: 5, pointHoverRadius: 7, borderColor: t.surface, borderWidth: 2, backgroundColor: alpha(color, 0.85), showLine: !!s.showLine });
        if (s.showLine) Object.assign(ds, { borderColor: color, borderWidth: 2 });
      } else if (kind === 'radar') {
        Object.assign(ds, { borderWidth: 2, backgroundColor: alpha(color, 0.15), pointRadius: 3 });
      }
      if (model.type === 'waterfall' && s.waterfall) {
        ds.data = s.waterfall.bars;
        ds.backgroundColor = s.waterfall.kinds.map((k) => (k === 'total' ? t.other : k === 'up' ? t.series[0] : t.series[1]));
        ds.borderRadius = 3;
      }
      return ds;
    });

    const fmt = model.yFormat || (series[0] && series[0].fmt);
    // conteos (todos enteros): sin marcas decimales en el eje
    const allInts = !model.percent && series.every((x) => (x.values || []).every((v) => v == null || Number.isInteger(v)));
    const valueAxis = {
      beginAtZero: model.min == null,
      min: model.min ?? undefined,
      max: model.percent ? 1 : model.max ?? undefined,
      stacked: !!model.stacked,
      grid: { color: t.grid, drawTicks: false },
      border: { display: false },
      ticks: { color: t.text2, padding: 8, font: { size: 11.5 }, callback: (v) => fmtTick(v, fmt), maxTicksLimit: 7, precision: allInts ? 0 : undefined },
      title: model.yTitle ? { display: true, text: model.yTitle, color: t.muted, font: { size: 11.5 } } : undefined,
    };
    const catAxis = {
      stacked: !!model.stacked,
      grid: { display: false },
      border: { color: t.grid },
      ticks: { color: t.text2, font: { size: 11.5 }, maxRotation: 0, autoSkip: true, autoSkipPadding: 10 },
      title: model.xTitle ? { display: true, text: model.xTitle, color: t.muted, font: { size: 11.5 } } : undefined,
    };
    if (model.type === 'scatter' || model.type === 'bubble') {
      catAxis.type = 'linear';
      catAxis.grid = { color: t.grid };
      catAxis.ticks.callback = (v) => fmtTick(v, model.xFormat);
    }
    if (model.horizontal) {
      catAxis.ticks.callback = (v, i) => wrapLabel(labels[i], 20);
      catAxis.afterFit = fitLabels(labels, t.font);
    }
    const scales = model.type === 'radar'
      ? { r: { grid: { color: t.grid }, angleLines: { color: t.grid }, pointLabels: { color: t.text2 }, ticks: { display: false } } }
      : model.horizontal
        ? { y: { ...catAxis, reverse: false }, x: valueAxis }
        : { x: catAxis, y: valueAxis };

    return {
      type: baseType,
      data: { labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: model.horizontal ? 'y' : 'x',
        interaction: { mode: model.type === 'scatter' || model.type === 'bubble' ? 'nearest' : 'index', intersect: false },
        animation: { duration: 450 },
        layout: { padding: { top: model.showValues ? 18 : 4, right: model.horizontal && model.showValues ? 36 : 4 } },
        scales,
        plugins: {
          legend: {
            display: datasets.length > 1,
            position: 'top',
            align: 'start',
            labels: { color: t.text2, usePointStyle: true, pointStyle: 'rectRounded', boxWidth: 10, boxHeight: 10, padding: 14, font: { size: 12 } },
          },
          tooltip: {
            backgroundColor: isDark() ? '#13221a' : '#0f2418',
            borderColor: isDark() ? '#2d4a37' : '#0f2418',
            borderWidth: 1,
            padding: 10,
            cornerRadius: 8,
            usePointStyle: true,
            boxPadding: 4,
            titleFont: { weight: '700' },
            callbacks: {
              title: (items) => (items[0] ? items[0].label : ''),
              label: (ctx) => {
                const raw = ctx.raw;
                if (model.type === 'waterfall' && s0(series).waterfall) return ` ${ctx.dataset.label}: ${fmtValue(s0(series).waterfall.deltas[ctx.dataIndex], ctx.dataset._fmt)}`;
                if (raw && typeof raw === 'object') return ` ${ctx.dataset.label}: (${fmtValue(raw.x, model.xFormat)}, ${fmtValue(raw.y, ctx.dataset._fmt)})`;
                return ` ${ctx.dataset.label}: ${fmtValue(raw, ctx.dataset._fmt)}`;
              },
            },
          },
          valueLabels: { enabled: !!model.showValues && labels.length <= 16 && datasets.length <= 3 && !(model.stacked && datasets.length > 1) },
        },
      },
      plugins: [valueLabels, crosshair],
    };
  }
  const s0 = (series) => series[0] || {};

  // Ancho del eje de categorías según la etiqueta más larga (Chart.js las recorta en gráficos angostos)
  function fitLabels(labels, font) {
    return (scale) => {
      const ctx = scale.ctx;
      ctx.save();
      ctx.font = `11.5px ${font}`;
      let w = 0;
      labels.forEach((l) => [].concat(wrapLabel(l, 20)).forEach((line) => (w = Math.max(w, ctx.measureText(String(line)).width))));
      ctx.restore();
      scale.width = Math.min(w + 14, scale.chart.width * 0.5);
    };
  }

  // Etiquetas largas en varias líneas para que el eje no las corte
  function wrapLabel(text, max) {
    const words = String(text ?? '').split(/\s+/);
    const lines = [];
    let cur = '';
    for (const w of words) {
      if (cur && (cur + ' ' + w).length > max) {
        lines.push(cur);
        cur = w;
      } else cur = cur ? cur + ' ' + w : w;
    }
    if (cur) lines.push(cur);
    return lines.length > 1 ? lines.slice(0, 3) : lines[0] || '';
  }

  let chartSeq = 0;
  function mountChart(container, model, height) {
    const id = 'c' + ++chartSeq;
    const wrap = el('div', { class: 'chart-wrap', style: { minHeight: (height || 300) + 'px' } });
    const canvas = el('canvas', { role: 'img', 'aria-label': (model.title || 'Gráfico') + ': ' + model.series.map((s) => s.name).join(', ') });
    wrap.append(canvas);
    container.append(wrap);
    const entry = { model, canvas, chart: null };
    entry.build = () => {
      if (entry.chart) entry.chart.destroy();
      entry.chart = new Chart(canvas, buildConfig(entry.model));
    };
    entry.build();
    state.charts.set(id, entry);
    return entry;
  }
  function rebuildCharts() {
    for (const [id, entry] of state.charts) {
      if (!entry.canvas.isConnected) {
        if (entry.chart) entry.chart.destroy();
        state.charts.delete(id);
        continue;
      }
      entry.build();
    }
  }

  function dataTableFor(model) {
    const table = el('table', { class: 'data' });
    const head = el('tr', {}, el('th', {}, 'Categoría'), ...model.series.map((s) => el('th', { class: 'num' }, s.name)));
    table.append(el('thead', {}, head));
    const body = el('tbody');
    model.labels.forEach((l, i) => {
      body.append(el('tr', {}, el('td', {}, l), ...model.series.map((s) => {
        const v = s.points ? s.points[i] && s.points[i].y : s.values[i];
        return el('td', { class: 'num' }, fmtValue(v, s.fmt));
      })));
    });
    (model.totals || []).forEach((tot) => {
      body.append(el('tr', { class: 'is-total' }, el('td', {}, tot.label), ...model.series.map((s) => el('td', { class: 'num' }, fmtValue(s.total, s.fmt)))));
    });
    table.append(body);
    return el('div', { class: 'data-table-inline' }, table);
  }

  /* ---------- gráficos por sexo: pictograma con íconos de mujer y hombre ---------- */
  function genderInfo(model) {
    if (!model || model.panels || model.series.length !== 1 || model.labels.length !== 2) return null;
    const kind = (l) => (/femen|mujer|female|\bf\b/i.test(l) ? 'woman' : /mascul|hombre|male|\bm\b/i.test(l) ? 'man' : null);
    const kinds = model.labels.map(kind);
    if (!kinds.includes('woman') || !kinds.includes('man')) return null;
    const vals = model.series[0].values.map((v) => (typeof v === 'number' ? v : 0));
    const sum = vals.reduce((a, b) => a + b, 0);
    if (!sum || vals.some((v) => v < 0)) return null;
    const fmt = model.series[0].fmt;
    return {
      total: model.totalText || fmtValue(sum, fmt),
      items: model.labels.map((l, i) => ({ label: l, kind: kinds[i], value: vals[i], text: fmtValue(vals[i], fmt), pct: vals[i] / sum, color: `var(--series-${i + 1})` })),
    };
  }

  // Figura recortada a la silueta (más grande que el ícono normal)
  function figure(kind) {
    const span = document.createElement('span');
    span.innerHTML = `<svg viewBox="6.3 1 11.4 21.4" fill="currentColor" aria-hidden="true">${ICONS[kind]}</svg>`;
    return span.firstChild;
  }

  function genderBlock(model, compact) {
    const g = genderInfo(model);
    const wrap = el('div', { class: 'gender' + (compact ? ' gender--compact' : '') });
    // 10 figuras: cada una equivale al 10 %
    const picto = el('div', { class: 'gender__picto', role: 'img', 'aria-label': g.items.map((x) => `${x.label} ${Math.round(x.pct * 100)}%`).join(', ') });
    const nFirst = Math.round(g.items[0].pct * 10);
    for (let i = 0; i < 10; i++) {
      const it = i < nFirst ? g.items[0] : g.items[1];
      const f = el('span', { class: 'gender__fig', style: { color: it.color } }, figure(it.kind));
      f.style.animationDelay = i * 40 + 'ms';
      picto.append(f);
    }
    wrap.append(picto);
    const bar = el('div', { class: 'gender__bar', 'aria-hidden': 'true' }, ...g.items.map((x) => el('span', { style: { width: (x.pct * 100).toFixed(2) + '%', background: x.color } })));
    wrap.append(bar);
    const stats = el('div', { class: 'gender__stats' });
    g.items.forEach((x) => stats.append(el('div', { class: 'gender__stat' },
      el('span', { class: 'gender__badge', style: { background: x.color } }, icon(x.kind, compact ? 18 : 24)),
      el('div', {},
        el('div', { class: 'gender__value' }, x.text, el('span', { class: 'gender__pct' }, el('span', { class: 'gender__dot' }, ' · '), (x.pct * 100).toFixed(1).replace('.', ',') + '%')),
        el('div', { class: 'gender__label' }, x.label)))));
    wrap.append(stats);
    return wrap;
  }

  function genderCard(model) {
    const g = genderInfo(model);
    const card = el('article', { class: 'card chart-card gender-card' });
    const tools = el('div', { class: 'card__tools' });
    card.append(el('div', { class: 'card__head' },
      el('div', { class: 'card__titles' }, el('h3', { class: 'card__title' }, model.title || 'Distribución por sexo'), model.subtitle ? el('div', { class: 'card__sub' }, model.subtitle) : null),
      el('div', { class: 'chart-total' }, el('span', {}, model.totalLabel || 'Total'), el('b', {}, g.total)),
      tools));
    const body = el('div', { class: 'card__body' }, genderBlock(model, false));
    card.append(body);
    let tableEl = null;
    const tableBtn = el('button', { class: 'icon-btn icon-btn--plain', title: 'Ver datos', 'aria-label': 'Ver datos', 'aria-pressed': 'false' }, icon('table', 16));
    tableBtn.addEventListener('click', () => {
      if (tableEl) {
        tableEl.remove();
        tableEl = null;
      } else body.append((tableEl = dataTableFor(model)));
      tableBtn.setAttribute('aria-pressed', String(!!tableEl));
    });
    tools.append(tableBtn);
    return card;
  }

  function chartCard(model, opts = {}) {
    const card = el('article', { class: 'card chart-card' });
    const tools = el('div', { class: 'card__tools' });
    const head = el('div', { class: 'card__head' },
      el('div', { class: 'card__titles' }, el('h3', { class: 'card__title' }, model.title || 'Gráfico'), model.subtitle ? el('div', { class: 'card__sub' }, model.subtitle) : null),
      model.totalText && model.type !== 'doughnut' ? el('div', { class: 'chart-total', title: 'El total no se grafica para no competir con las categorías' }, el('span', {}, model.totalLabel || 'Total'), el('b', {}, model.totalText)) : null,
      tools);
    card.append(head);
    if (opts.controls) card.append(opts.controls);
    const body = el('div', { class: 'card__body' });
    card.append(body);
    const panels = model.panels || [model];
    const entries = panels.map((p) => {
      const holder = el('div', { class: 'chart-panel' });
      if (panels.length > 1 && p.panelLabel) holder.append(el('div', { class: 'chart-panel__label' }, p.panelLabel));
      body.append(holder);
      return mountChart(holder, p, panels.length > 1 ? Math.max(180, (opts.height || 300) * 0.62) : opts.height);
    });
    let tableEl = null;
    const tableBtn = el('button', { class: 'icon-btn icon-btn--plain', title: 'Ver datos del gráfico', 'aria-label': 'Ver datos del gráfico', 'aria-pressed': 'false' }, icon('table', 16));
    tableBtn.addEventListener('click', () => {
      if (tableEl) {
        tableEl.remove();
        tableEl = null;
        tableBtn.setAttribute('aria-pressed', 'false');
        return;
      }
      tableEl = el('div', {}, ...panels.map((p) => dataTableFor(p)));
      body.append(tableEl);
      tableBtn.setAttribute('aria-pressed', 'true');
    });
    const pngBtn = el('button', { class: 'icon-btn icon-btn--plain', title: 'Descargar imagen', 'aria-label': 'Descargar imagen' }, icon('download', 16));
    pngBtn.addEventListener('click', () => {
      const e = entries[0];
      const a = el('a', { href: e.chart.toBase64Image('image/png', 1), download: slug(model.title || 'grafico') + '.png' });
      a.click();
    });
    tools.append(tableBtn, pngBtn);
    card._entries = entries;
    return card;
  }

  /* ---------- gráfico de Excel -> modelo ---------- */
  const TYPE_MAP = {
    barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', stockChart: 'line',
    areaChart: 'area', area3DChart: 'area', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie',
    doughnutChart: 'doughnut', scatterChart: 'scatter', bubbleChart: 'bubble', radarChart: 'radar',
    surfaceChart: 'line', surface3DChart: 'line',
  };

  // Secciones configuradas como "torta": gráficos de una serie y pocas categorías -> anillo
  function preferPie(model, partTitle) {
    const wanted = (state.config.seccionesTorta || []).map((x) => Correcciones.fix(String(x)).toLowerCase().trim());
    if (!model || !partTitle || !wanted.includes(String(partTitle).toLowerCase().trim())) return model;
    if (model.panels || model.series.length !== 1 || !['bar', 'funnel', 'pie', 'doughnut'].includes(model.type)) return model;
    const vals = model.series[0].values;
    if (model.labels.length < 2 || model.labels.length > 8 || vals.some((v) => typeof v === 'number' && v < 0)) return model;
    return { ...model, type: 'doughnut', horizontal: false, stacked: false, percent: false, cutout: 58, showPercent: true };
  }

  // Total del anillo: el de Excel si existía, o la suma de las porciones si son conteos
  function withCenterTotal(model) {
    if (!model || model.type !== 'doughnut' || model.totalText) return model;
    const s = model.series[0];
    const vals = (s && s.values) || [];
    if (!vals.length || /%/.test(s.fmt || '') || vals.some((v) => typeof v === 'number' && !Number.isInteger(v))) return model;
    return { ...model, totalText: fmtValue(vals.reduce((a, v) => a + (v || 0), 0), s.fmt), totalLabel: 'Total' };
  }

  // Tarjeta para un gráfico que en Excel usa los mismos datos que otro
  function copyCard(model) {
    return el('div', { class: 'card text-card placeholder-card' }, icon('chart', 22), el('div', {}, el('b', {}, model.title),
      el('p', {}, `En el Excel este gráfico todavía usa los mismos datos que «${model.copyOf}», así que no se muestran aquí para no presentar cifras equivocadas. Para verlo, en Excel cambia sus datos (clic derecho sobre el gráfico → Seleccionar datos) por la tabla que le corresponde.`)));
  }

  // Gráficos de una hoja que apuntan exactamente a los mismos rangos que uno anterior
  function markChartCopies(items) {
    const seen = new Map();
    // el original es el primero en la hoja (de arriba abajo, de izquierda a derecha)
    const charts = items.filter((i) => i.type === 'chart').sort((x, y) => x.pos.from.row - y.pos.from.row || x.pos.from.col - y.pos.from.col);
    for (const it of charts) {
      const refs = it.chart.groups.flatMap((g) => g.series.map((s) => [s.cat && s.cat.f, s.val && s.val.f].join('|')));
      if (!refs.length || refs.some((r) => !r.split('|')[1])) continue;
      const key = refs.join(';');
      if (seen.has(key)) it.copyOf = seen.get(key);
      else seen.set(key, it);
    }
  }

  function chartModelFor(item, sheet, partTitle) {
    return withCenterTotal(preferPie(excelChartModel(item, sheet), partTitle));
  }

  // Categorías "Total" (por nombre, o la última si es la suma de las demás): se quitan del
  // gráfico para que no compitan con las demás y se muestran escritas
  function extractTotals(allSeries, labels) {
    if (allSeries.some((s) => ['waterfall', 'scatter', 'bubble', 'line', 'area'].includes(s.kind))) return [];
    const idx = new Set();
    labels.forEach((l, i) => A.isTotal(l) && idx.add(i));
    if (!idx.size && allSeries.length === 1 && labels.length >= 3 && !labels.some((l) => A.isTimeLike(l))) {
      const v = allSeries[0].values;
      const last = v.length - 1;
      const rest = v.slice(0, last).reduce((s, x) => s + (x || 0), 0);
      if (v[last] && rest && Math.abs(v[last] - rest) <= Math.abs(rest) * 0.01 + 1e-9) idx.add(last);
    }
    if (!idx.size || labels.length - idx.size < 2) return [];
    const out = Array.from(idx).map((i) => ({
      index: i,
      label: A.isTotal(labels[i]) ? labels[i] : 'Total',
      first: allSeries[0].values[i],
      fmt: allSeries[0].fmt,
    }));
    const firstIdx = Math.min(...idx);
    allSeries.forEach((s) => {
      s.total = s.values[firstIdx];
      s.values = s.values.filter((_, i) => !idx.has(i));
      if (s.slices) s.slices = s.slices.filter((_, i) => !idx.has(i));
    });
    return out;
  }

  // Nombre de serie ausente en Excel ("Serie1"): usar el encabezado sobre el rango de valores
  function headerAbove(f) {
    if (!f) return null;
    const m = String(f).match(/^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)/);
    if (!m) return null;
    const ws = state.wb.Sheets[m[1] ? m[1].replace(/''/g, "'") : m[2]];
    const row = +m[4] - 2;
    if (!ws || row < 0) return null;
    const c = XLSX.utils.decode_col(m[3]);
    const merge = (ws['!merges'] || []).find((g) => row >= g.s.r && row <= g.e.r && c >= g.s.c && c <= g.e.c);
    const cell = ws[XLSX.utils.encode_cell(merge ? merge.s : { r: row, c })];
    return cell && typeof cell.v === 'string' && cell.v.trim() ? Correcciones.fix(cell.v.trim()) : null;
  }

  function excelChartModel(item, sheetName) {
    const chart = item.chart;
    const wb = state.wb;
    const theme = state.parts && state.parts.theme;
    const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((k) => theme && theme.colors[k]);
    const sources = new Set();
    const allSeries = [];
    let labels = [];
    let xFormat = null;
    for (const g of chart.groups) {
      let kind = TYPE_MAP[g.type] || 'bar';
      if (g.type === 'chartEx') kind = chart.layout === 'waterfall' ? 'waterfall' : chart.layout === 'funnel' ? 'funnel' : 'bar';
      const valAxis = chart.axes.filter((a) => g.axIds.includes(a.id) && a.kind === 'valAx');
      const yAxis = valAxis.find((a) => a.pos === 'l' || a.pos === 'r') || valAxis[valAxis.length - 1] || null;
      for (const s of g.series.slice().sort((a, b) => a.order - b.order)) {
        [s.cat && s.cat.f, s.val && s.val.f, s.name && s.name.f].forEach((f) => A.refSheets(f).forEach((x) => sources.add(x)));
        const v = A.seriesValues(s.val, wb, true);
        if (!v.values) continue;
        let cats = A.seriesValues(s.cat, wb, kind === 'scatter' || kind === 'bubble');
        if (cats.values && cats.values.length > labels.length && kind !== 'scatter' && kind !== 'bubble') labels = cats.values.map((x) => (x == null ? '' : String(x).trim()));
        const entry = {
          name: A.seriesName(s.name, wb, headerAbove(s.val && s.val.f) || null),
          values: v.values,
          fmt: v.fmt || (s.val && s.val.formatCode) || null,
          color: s.color || accents[allSeries.length % 6] || null,
          slices: labels.map((_, i) => s.dPt[i] || accents[i % 6] || null),
          kind,
          axis: yAxis ? yAxis.id : 'main',
          axisInfo: yAxis,
          smooth: s.smooth,
          noMarker: s.noMarker,
          showValues: (s.labels && s.labels.showVal) || (g.labels && g.labels.showVal),
          showPercent: (s.labels && s.labels.showPercent) || (g.labels && g.labels.showPercent),
          group: g,
        };
        if (kind === 'scatter' || kind === 'bubble') {
          const xs = cats.values || v.values.map((_, i) => i + 1);
          xFormat = cats.fmt || xFormat;
          const sizes = kind === 'bubble' ? A.seriesValues(s.size, wb, true).values || [] : [];
          entry.points = v.values.map((y, i) => (y == null || xs[i] == null ? null : { x: +xs[i], y, r: kind === 'bubble' ? 4 + Math.sqrt(Math.abs(sizes[i] || 1)) * 2 : undefined })).filter(Boolean);
          entry.showLine = /line/i.test(g.scatterStyle || '') && !/marker$/i.test(g.scatterStyle || 'x');
          if (g.scatterStyle === 'lineMarker' || g.scatterStyle === 'smoothMarker') entry.showLine = true;
        }
        if (kind === 'waterfall') {
          let acc = 0;
          const bars = [], kinds = [];
          const totals = new Set(s.subtotals || []);
          v.values.forEach((d, i) => {
            if (totals.has(i)) {
              acc = d || 0;
              bars.push([0, acc]);
              kinds.push('total');
              return;
            }
            const start = acc;
            acc += d || 0;
            bars.push([start, acc]);
            kinds.push((d || 0) >= 0 ? 'up' : 'down');
          });
          entry.waterfall = { bars, kinds, deltas: v.values };
        }
        allSeries.push(entry);
      }
    }
    // series sin ningún valor (columnas vacías en Excel) no aportan y ensucian la leyenda
    for (let i = allSeries.length - 1; i >= 0 && allSeries.length > 1; i--)
      if (allSeries[i].values.every((v) => v == null || v === 0)) allSeries.splice(i, 1);
    if (!allSeries.length) return null;
    const totals = extractTotals(allSeries, labels);
    if (totals.length) labels = labels.filter((_, i) => !totals.some((t) => t.index === i));
    if (!labels.length) labels = allSeries[0].values.map((_, i) => String(i + 1));

    let title = chart.title && chart.title.text;
    if (chart.title && chart.title.f) {
      const cells = A.resolveRef(chart.title.f, wb);
      if (cells && cells[0]) title = A.cellText(cells[0]);
    }
    if (!title && !chart.autoTitleDeleted && allSeries.length === 1 && allSeries[0].name) title = allSeries[0].name;
    if (!title) title = item.name && !/^(chart|gráfico|grafico)\s*\d+$/i.test(item.name) ? item.name : allSeries.length === 1 && allSeries[0].name ? allSeries[0].name : 'Gráfico';
    // gráfico copiado de otro sin cambiarle los datos: no mostrar cifras que no son suyas
    if (item.copyOf) {
      const owner = excelChartModel(item.copyOf, sheetName);
      const norm = (t) => String(t || '').toLowerCase().trim();
      if (owner && norm(owner.title) !== norm(title)) return { title, subtitle: 'Hoja ' + sheetName, copyOf: owner.title, labels: [], series: [] };
    }

    // series sin nombre en Excel ("Serie1"): nombrarlas por lo que muestran
    allSeries.forEach((x, i) => {
      if (x.name) return;
      const pct = /%/.test(x.fmt || '');
      x.name = pct ? 'Porcentaje' : allSeries.filter((y) => !y.name).length === 1 || i === 0 ? title : `${title} (${i + 1})`;
    });
    const g0 = allSeries[0].group;
    const mainKind = allSeries[0].kind;
    const pie = mainKind === 'pie' || mainKind === 'doughnut';
    const base = {
      type: mainKind,
      labels,
      horizontal: chart.groups.some((g) => g.barDir === 'bar') && !pie,
      stacked: chart.groups.some((g) => g.grouping === 'stacked' || g.grouping === 'percentStacked'),
      percent: chart.groups.some((g) => g.grouping === 'percentStacked'),
      cutout: g0.holeSize,
      showValues: allSeries.some((s) => s.showValues),
      showPercent: allSeries.some((s) => s.showPercent),
      xFormat,
      totals,
      totalText: totals.length ? fmtValue(totals[0].first, totals[0].fmt) : null,
      totalLabel: totals.length ? totals[0].label : null,
    };
    const valAxes = chart.axes.filter((a) => a.kind === 'valAx');
    const catAxis = chart.axes.find((a) => a.kind === 'catAx' || a.kind === 'dateAx');
    if (catAxis && catAxis.title) base.xTitle = catAxis.title;

    const subtitle = (sources.size ? 'Fuente: ' + Array.from(sources).join(', ') : 'Hoja ' + sheetName) + (chart.pivot ? ' · tabla dinámica' : '');

    // Excel usa colores de serie únicos: asignar orden fijo de la paleta por serie
    allSeries.forEach((s, i) => (s.colorIndex = i));

    // Eje secundario -> paneles separados con el mismo eje X (no doble eje)
    const axisIds = Array.from(new Set(allSeries.map((s) => s.axis)));
    if (!pie && axisIds.length > 1 && mainKind !== 'scatter') {
      const panels = axisIds.map((axId) => {
        const ss = allSeries.filter((s) => s.axis === axId);
        const ax = ss[0].axisInfo;
        return {
          ...base,
          type: ss[0].kind,
          title,
          series: ss,
          panelLabel: ss.map((s) => s.name).join(' · ') + (ax && ax.title ? ' — ' + ax.title : ''),
          yFormat: ax && ax.format,
          yTitle: ax && ax.title,
          min: ax && ax.min,
          max: ax && ax.max,
          stacked: ss.some((s) => s.group.grouping === 'stacked' || s.group.grouping === 'percentStacked'),
        };
      });
      return { title, subtitle, labels, series: allSeries, panels, totals, totalText: base.totalText, totalLabel: base.totalLabel };
    }
    const ax = valAxes.find((a) => !a.deleted) || valAxes[0];
    return {
      ...base,
      title,
      subtitle,
      series: pie ? allSeries.slice(0, 1) : allSeries,
      yFormat: ax && ax.format,
      yTitle: ax && ax.title,
      min: ax && ax.min,
      max: ax && ax.max,
    };
  }

  /* =======================================================================
     Construcción de secciones
     ======================================================================= */
  const DASH_RE = /dash|tablero|resumen|panel|indicador|inicio|portada|kpi|gerencial|tablero/i;

  function findDashboardSheet() {
    const names = state.wb.SheetNames;
    const items = (n) => (state.parts && state.parts.sheets[n] ? state.parts.sheets[n].items : []);
    const visible = names.filter((n) => sheetState(n) === 'visible');
    const byName = visible.find((n) => DASH_RE.test(n));
    if (byName) return byName;
    let best = null, bestCount = 0;
    for (const n of visible) {
      const c = items(n).filter((i) => i.type === 'chart').length;
      if (c > bestCount) {
        best = n;
        bestCount = c;
      }
    }
    return bestCount >= 2 ? best : null;
  }

  function sheetState(name) {
    if (state.parts && state.parts.sheets[name]) return state.parts.sheets[name].state;
    const meta = ((state.wb.Workbook && state.wb.Workbook.Sheets) || [])[state.wb.SheetNames.indexOf(name)];
    return meta && meta.Hidden ? (meta.Hidden === 2 ? 'veryHidden' : 'hidden') : 'visible';
  }

  const DATA_RE = /^(bd|base|datos|data|param|maestr|temp|calc|aux)|\bbd\b|tabla(s)? din/i;
  const FALLBACK_RE = /no est[aá] disponible en su versi[oó]n|not available in your version|isn't available in your version/i;
  const sheetItems = (n) => (n && state.parts && state.parts.sheets[n] ? state.parts.sheets[n].items : []);

  // Divide la hoja Dashboard en partes usando sus rótulos (formas con títulos cortos)
  function dashboardModel(sheet) {
    const items = sheetItems(sheet).filter((i) => !(i.type === 'shape' && FALLBACK_RE.test(i.text || '')));
    const a = state.sheets[sheet] || { tables: [], kpis: [] };
    const span = (i) => ({ w: i.pos.to.col - i.pos.from.col, h: i.pos.to.row - i.pos.from.row });
    const images = items.filter((i) => i.type === 'image' && !i.macro);
    const logo = images.filter((i) => i.pos.from.row <= 3).sort((x, y) => x.pos.from.col - y.pos.from.col)[0] || null;
    const bigImages = images.filter((i) => i !== logo && span(i).w >= 6 && span(i).h >= 6 && !(i.unsupported && span(i).w * span(i).h < 60));
    const charts = items.filter((i) => i.type === 'chart');
    const covers = charts.concat(images);
    const covered = (r, c) => covers.some((i) => r >= i.pos.from.row && r < i.pos.to.row && c >= i.pos.from.col && c < i.pos.to.col);
    const kpiInfo = dashboardKpis(sheet, items);
    const headings = items
      .filter((i) => i.type === 'shape' && !i.textlink && !i.macro && i.text && i.text.length >= 3 && i.text.length <= 48 && !i.text.includes('\n') && span(i).h <= 4 && !kpiInfo.used.has(i))
      .sort((x, y) => x.pos.from.row - y.pos.from.row || x.pos.from.col - y.pos.from.col);
    const texts = items.filter((i) => i.type === 'shape' && !i.macro && !i.textlink && !kpiInfo.used.has(i) && !headings.includes(i) && i.text && i.text.length > 48);
    const visibleKpi = (k) => k.fromShape || (!k.hidden && !covered(k.r, k.c));
    // con secciones: solo los indicadores porcentuales destacados; sin secciones: todos los visibles
    const headline = headings.length >= 3
      ? kpiInfo.kpis.filter((k) => visibleKpi(k) && (k.fromShape || (k.format && /%/.test(k.format))))
      : kpiInfo.kpis.filter(visibleKpi).slice(0, 16);
    const tables = a.tables.filter((t) => !covered(t.bounds.r1, t.bounds.c1) && !covered(t.headerRow, t.bounds.c1));
    const useParts = headings.length >= 3;
    const parts = useParts ? headings.map((h) => ({ title: h.text.trim(), heading: h, items: [], tables: [], kpis: [], notes: [] })) : [];
    const general = { title: 'General', heading: null, items: [], tables: [], kpis: [], notes: [] };
    const assign = (r, c) => {
      let best = null;
      for (const p of parts) {
        const h = p.heading.pos.from;
        if (h.row > r + 1 || h.col > c + 3) continue;
        if (!best || h.row > best.heading.pos.from.row || (h.row === best.heading.pos.from.row && h.col > best.heading.pos.from.col)) best = p;
      }
      return best || general;
    };
    charts.concat(bigImages, texts).forEach((i) => assign(i.pos.from.row, i.pos.from.col).items.push(i));
    tables.forEach((t) => assign(t.headerRow, t.bounds.c1).tables.push(t));
    headline.forEach((k) => (k.r != null ? assign(k.r, k.c) : general).kpis.push(k));
    // notas de texto sueltas (fuentes, aclaraciones) dentro de cada parte
    const header0 = A.sheetHeader(a);
    (a.allTexts || a.texts || [])
      .filter((x) => x.text.length > 25 && x.text !== header0.title && !covered(x.r, x.c) && !/^etiquetas de/i.test(x.text))
      .forEach((x) => useParts && assign(x.r, x.c).notes.push(x));
    parts.forEach((p) => p.items.sort((x, y) => x.pos.from.row - y.pos.from.row || x.pos.from.col - y.pos.from.col));
    const all = parts.filter((p) => p.items.length || p.tables.length || p.kpis.length);
    if (general.items.length || general.tables.length || (!useParts && general.kpis.length)) all.unshift(general);
    const header = A.sheetHeader(a);
    return { sheet, parts: all, headline, logo, header, buttons: items.filter((i) => (i.type === 'control' && i.control === 'Button') || ((i.type === 'shape' || i.type === 'image') && i.macro)), split: useParts };
  }

  function buildSections() {
    const wb = state.wb;
    const dash = findDashboardSheet();
    const macroSheets = new Map(); // hoja -> procedimientos que la usan
    const navTargets = new Set();
    for (const m of state.macros.modules)
      for (const p of m.procedures) {
        p.sheets.forEach((s) => macroSheets.set(s, (macroSheets.get(s) || []).concat({ module: m.name, proc: p })));
        p.navigatesTo.forEach((s) => navTargets.add(s));
      }
    const chartSources = new Set();
    for (const it of sheetItems(dash))
      if (it.type === 'chart') for (const g of it.chart.groups) for (const s of g.series) [s.val && s.val.f, s.cat && s.cat.f].forEach((f) => A.refSheets(f).forEach((x) => x !== dash && chartSources.add(x)));

    const sections = [];
    const ids = new Set();
    const uid = (base) => {
      let id = slug(base), n = 2;
      while (ids.has(id)) id = slug(base) + '-' + n++;
      ids.add(id);
      return id;
    };
    state.dash = dash ? dashboardModel(dash) : null;
    sections.push({ id: uid('resumen'), kind: 'dashboard', title: 'Resumen', sheet: dash, icon: 'dashboard', group: 'main' });
    if (state.dash && state.dash.split)
      state.dash.parts.forEach((part, i) => {
        const n = part.items.filter((x) => x.type === 'chart').length;
        sections.push({ id: uid(part.title), kind: 'dashpart', title: Correcciones.fix(part.title), sheet: dash, part, index: i, icon: topicIcon(part.title, n ? 'chart' : 'table'), group: 'dash', badge: n || null });
      });

    const ordered = wb.SheetNames.filter((n) => n !== dash).map((n, i) => {
      const a = state.sheets[n];
      const charts = sheetItems(n).filter((x) => x.type === 'chart').length;
      const hasContent = (a && (a.tables.length || a.kpis.length)) || charts;
      // prioridad: hojas a las que llevan las macros/botones, luego las que alimentan el tablero
      const score = (navTargets.has(n) ? 0 : chartSources.has(n) ? 1 : 2) * 1000 + i;
      return { n, a, charts, hasContent, score };
    });
    ordered.sort((x, y) => x.score - y.score);
    for (const o of ordered) {
      if (!o.hasContent) continue;
      const st = sheetState(o.n);
      const info = state.parts && state.parts.sheets[o.n];
      const big = o.a && o.a.tables.some((t) => t.rows.length > 300);
      const isData = DATA_RE.test(o.n) || (info && info.pivots >= 3) || (big && !o.charts);
      sections.push({
        id: uid(o.n),
        kind: 'sheet',
        title: Correcciones.fix(o.n),
        sheet: o.n,
        icon: o.charts ? 'chart' : 'sheet',
        group: st !== 'visible' ? 'hidden' : isData ? 'data' : 'sheets',
        macroRefs: macroSheets.get(o.n) || [],
        navTarget: navTargets.has(o.n),
        feedsDashboard: chartSources.has(o.n),
        badge: o.a ? o.a.tables.length || null : null,
      });
    }
    sections.push({ id: uid('macros'), kind: 'macros', title: 'Macros', icon: 'macro', group: 'tools', badge: state.macros.modules.reduce((s, m) => s + m.procedures.length, 0) || null });
    sections.push({ id: uid('archivo'), kind: 'file', title: 'Archivo y actualización', icon: 'file', group: 'tools' });
    state.sections = sections;
    state.dashSheet = dash;
  }

  function sectionForSheet(name) {
    return state.sections.find((s) => s.sheet === name && s.kind !== 'dashpart') || null;
  }

  /* ---------- navegación ---------- */
  function renderNav() {
    const nav = $('#nav');
    nav.innerHTML = '';
    const link = (s) =>
      el('a', { href: '#/' + s.id, 'data-id': s.id }, icon(s.icon, 17), el('span', { class: 'nav-text' }, s.title), s.badge ? el('span', { class: 'nav-badge' }, s.badge) : null);
    const by = (g) => state.sections.filter((s) => s.group === g);
    const collapsible = (label, list, open) => {
      const d = el('details', {}, el('summary', { class: 'nav-label' }, `${label} (${list.length})`), ...list.map(link));
      if (open) d.open = true;
      return d;
    };
    nav.append(...by('main').map(link));
    if (by('dash').length) nav.append(el('div', { class: 'nav-label' }, 'Dashboard' + (state.dash && state.dash.header.period ? ' · ' + state.dash.header.period : '')), ...by('dash').map(link));
    if (by('sheets').length) nav.append(el('div', { class: 'nav-label' }, by('dash').length ? 'Otras hojas' : 'Hojas del archivo'), ...by('sheets').map(link));
    if (by('data').length) nav.append(collapsible('Bases de datos', by('data'), false));
    if (by('hidden').length) nav.append(collapsible('Hojas ocultas', by('hidden'), false));
    nav.append(el('div', { class: 'nav-label' }, 'Herramientas'), ...by('tools').map(link));
    // si la sección activa está dentro de un grupo plegado, abrirlo
    const active = nav.querySelector(`a[data-id="${currentId()}"]`);
    if (active && active.closest('details')) active.closest('details').open = true;
  }

  function currentId() {
    const h = decodeURIComponent(location.hash.replace(/^#\/?/, ''));
    return state.sections.some((s) => s.id === h) ? h : state.sections[0] && state.sections[0].id;
  }

  function showSection(id) {
    const sec = state.sections.find((s) => s.id === id);
    if (!sec) return;
    document.querySelectorAll('#nav a').forEach((a) => (a.dataset.id === id ? a.setAttribute('aria-current', 'page') : a.removeAttribute('aria-current')));
    const active = $('#nav a[aria-current="page"]');
    if (active && window.innerWidth <= 860) active.scrollIntoView({ inline: 'center', block: 'nearest' });
    document.querySelectorAll('.section').forEach((n) => n.classList.toggle('is-active', n.id === 'sec-' + id));
    let node = document.getElementById('sec-' + id);
    if (!node) {
      node = el('section', { class: 'section is-active', id: 'sec-' + id, 'aria-labelledby': 'h-' + id });
      $('#content').append(node);
    }
    if (!state.rendered.has(id)) {
      state.rendered.add(id);
      try {
        RENDER[sec.kind](node, sec);
      } catch (e) {
        console.error(e);
        node.append(el('div', { class: 'notice notice--error' }, 'No se pudo mostrar esta sección: ' + e.message));
      }
    }
    document.title = sec.title + ' · ' + (state.config.titulo || 'Tablero');
  }

  function goTo(id) {
    if (location.hash !== '#/' + id) location.hash = '#/' + id;
    else showSection(id);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function sectionHead(sec, eyebrow, subtitle, extra, iconName) {
    return el('header', { class: 'section-head' },
      iconName ? el('div', { class: 'section-icon', 'aria-hidden': 'true' }, icon(iconName, 30)) : null,
      el('div', { class: 'section-head__text' },
        el('div', { class: 'eyebrow' }, eyebrow),
        el('h1', { class: 'section-title', id: 'h-' + sec.id }, sec.title),
        subtitle ? el('p', { class: 'section-sub' }, subtitle) : null),
      extra || null);
  }

  /* ---------- tarjetas de indicador ---------- */
  function sparkline(values) {
    const w = 200, h = 34, pad = 3;
    const nums = values.filter((v) => typeof v === 'number');
    if (nums.length < 2) return null;
    const min = Math.min(...nums), max = Math.max(...nums);
    const sx = (i) => pad + (i / (values.length - 1)) * (w - pad * 2);
    const sy = (v) => h - pad - ((v - min) / (max - min || 1)) * (h - pad * 2);
    const pts = values.map((v, i) => (typeof v === 'number' ? `${sx(i).toFixed(1)},${sy(v).toFixed(1)}` : null)).filter(Boolean);
    const lastX = sx(values.length - 1), lastY = sy(values[values.length - 1]);
    const svg = `<svg class="kpi__spark" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" aria-hidden="true">
      <polyline points="${pts.join(' ')} ${lastX},${h} ${pad},${h}" fill="var(--brand-100)" stroke="none"/>
      <polyline points="${pts.join(' ')}" fill="none" stroke="var(--series-1)" stroke-width="2" vector-effect="non-scaling-stroke" stroke-linejoin="round"/>
      <circle cx="${lastX}" cy="${lastY}" r="3" fill="var(--series-1)" stroke="var(--surface)" stroke-width="2" vector-effect="non-scaling-stroke"/></svg>`;
    const d = document.createElement('div');
    d.innerHTML = svg;
    return d.firstElementChild;
  }

  function kpiCard(k, i) {
    const card = el('div', { class: 'card kpi' + (i % 4 === 1 ? ' kpi--gold' : '') });
    const ic = topicIcon(k.label + ' ' + (k.source || ''), null);
    if (ic) card.append(el('div', { class: 'kpi__icon', 'aria-hidden': 'true' }, icon(ic, 20)));
    card.append(el('div', { class: 'kpi__label', title: k.label }, k.label));
    const valueText = k.text || fmtValue(k.value, k.format);
    card.append(el('div', { class: 'kpi__value' }, valueText, k.unit ? el('span', { class: 'kpi__unit' }, k.unit) : null));
    const meta = el('div', { class: 'kpi__meta' });
    if (k.delta != null && isFinite(k.delta)) {
      const up = k.delta >= 0;
      meta.append(el('span', { class: 'delta ' + (up ? 'delta--up' : 'delta--down') }, (up ? '▲ +' : '▼ ') + (k.delta * 100).toFixed(1).replace('.', ',') + '%'));
      if (k.prevPeriod) meta.append(el('span', {}, 'vs. ' + k.prevPeriod));
    }
    if (k.period) meta.prepend(el('span', {}, k.period));
    if (k.source) meta.append(el('span', {}, k.source));
    if (meta.childNodes.length) card.append(meta);
    if (k.spark) {
      const sp = sparkline(k.spark);
      if (sp) card.append(sp);
    }
    return card;
  }

  /* ---------- tabla de datos ---------- */
  function dataTableCard(table, opts = {}) {
    const card = el('article', { class: 'card' });
    const cols = table.columns;
    const head = el('div', { class: 'card__head' },
      el('div', { class: 'card__titles' }, el('h3', { class: 'card__title' }, opts.title || table.title || 'Datos'), el('div', { class: 'card__sub' }, `${table.rows.length} filas · ${cols.length} columnas · rango ${table.range}`)));
    card.append(head);
    const search = el('input', { class: 'search', type: 'search', placeholder: 'Buscar en la tabla…', 'aria-label': 'Buscar en la tabla' });
    const csvBtn = el('button', { class: 'btn btn--sm' }, icon('download', 14), 'CSV');
    card.append(el('div', { class: 'table-tools' }, search, csvBtn));
    const scroll = el('div', { class: 'table-scroll' });
    card.append(scroll);
    const foot = el('div', { class: 'table-foot' });
    card.append(foot);
    let sortCol = -1, sortDir = 1, limit = opts.limit || 25;
    const draw = () => {
      const q = search.value.trim().toLowerCase();
      let rows = table.rows;
      if (q) rows = rows.filter((r) => r.cells.some((c) => A.cellText(c).toLowerCase().includes(q)));
      if (sortCol >= 0) {
        rows = rows.slice().sort((a, b) => {
          const x = a.cells[sortCol], y = b.cells[sortCol];
          const xv = x ? (x.v instanceof Date ? +x.v : x.v) : null, yv = y ? (y.v instanceof Date ? +y.v : y.v) : null;
          if (xv == null) return 1;
          if (yv == null) return -1;
          return (typeof xv === 'number' && typeof yv === 'number' ? xv - yv : String(xv).localeCompare(String(yv), 'es')) * sortDir;
        });
      }
      const t = el('table', { class: 'data' });
      const tr = el('tr');
      cols.forEach((c, i) => {
        const th = el('th', { class: c.type === 'number' ? 'num' : '', scope: 'col', 'aria-sort': sortCol === i ? (sortDir > 0 ? 'ascending' : 'descending') : null }, c.name, el('span', { class: 'sort' }, sortCol === i ? (sortDir > 0 ? '↑' : '↓') : '↕'));
        th.addEventListener('click', () => {
          sortDir = sortCol === i ? -sortDir : 1;
          sortCol = i;
          draw();
        });
        tr.append(th);
      });
      t.append(el('thead', {}, tr));
      const tb = el('tbody');
      rows.slice(0, limit).forEach((r) => {
        const total = A.isTotal(A.cellText(r.cells[0]));
        tb.append(el('tr', { class: total ? 'is-total' : '' }, ...r.cells.map((c, i) => el('td', { class: cols[i].type === 'number' ? 'num' : '' }, A.cellText(c)))));
      });
      t.append(tb);
      scroll.innerHTML = '';
      scroll.append(t);
      foot.innerHTML = '';
      foot.append(el('span', {}, `Mostrando ${Math.min(limit, rows.length)} de ${rows.length} filas`));
      if (rows.length > limit) {
        foot.append(el('button', { class: 'btn btn--sm', onclick: () => { limit += 100; draw(); } }, 'Ver más filas'));
      }
    };
    search.addEventListener('input', () => {
      limit = opts.limit || 25;
      draw();
    });
    csvBtn.addEventListener('click', () => {
      const lines = [cols.map((c) => c.name)].concat(table.rows.map((r) => r.cells.map((c) => (c ? (typeof c.v === 'number' ? c.v : A.cellText(c)) : ''))));
      const csv = lines.map((l) => l.map((x) => `"${String(x).replace(/"/g, '""')}"`).join(';')).join('\r\n');
      const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
      const a = el('a', { href: URL.createObjectURL(blob), download: slug(table.sheet + '-' + (table.title || table.range)) + '.csv' });
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
    });
    draw();
    return card;
  }

  /* ---------- tabla -> gráfico interactivo ---------- */
  function tableChartCard(table) {
    const spec = A.suggestChart(table);
    if (!spec) return null;
    const st = {
      type: spec.type,
      transpose: spec.transpose,
      active: new Set(
        (spec.numeric.filter((c) => !A.isTotal(c.name)).length ? spec.numeric.filter((c) => !A.isTotal(c.name)) : spec.numeric)
          .slice(0, spec.transpose ? 24 : 6)
          .map((c) => c.name),
      ),
    };
    const controls = el('div', { class: 'chart-controls' });
    const card = el('div');
    const title = table.title || `${table.sheet} · ${table.range}`;
    const render = () => {
      const columns = spec.numeric.filter((c) => st.active.has(c.name));
      if (!columns.length) return;
      const data = A.chartDataFromTable(table, spec, { columns, transpose: st.transpose });
      const type = st.type === 'stacked' ? 'bar' : st.type;
      const model = {
        title,
        subtitle: (spec.cat ? `Por ${spec.cat.name}` : '') + (st.transpose ? ' · series por fila' : ''),
        type,
        labels: data.labels,
        series: data.series.slice(0, type === 'doughnut' ? 1 : 8),
        stacked: st.type === 'stacked',
        horizontal: spec.horizontal && (type === 'bar') && !st.transpose,
        yFormat: data.series[0] && data.series[0].fmt,
      };
      const fresh = chartCard(model, { controls, height: 320 });
      card.replaceChildren(fresh);
    };
    const types = [['bar', 'Barras'], ['stacked', 'Apiladas'], ['line', 'Líneas'], ['area', 'Área'], ['doughnut', 'Dona']];
    const seg = el('div', { class: 'seg', role: 'group', 'aria-label': 'Tipo de gráfico' });
    types.forEach(([k, label]) => {
      const b = el('button', { type: 'button', 'aria-pressed': String(st.type === k) }, label);
      b.addEventListener('click', () => {
        st.type = k;
        seg.querySelectorAll('button').forEach((x) => x.setAttribute('aria-pressed', String(x === b)));
        render();
      });
      seg.append(b);
    });
    controls.append(seg);
    if (spec.cat) {
      const tr = el('button', { type: 'button', class: 'chip', 'aria-pressed': String(st.transpose), title: 'Intercambiar filas y columnas' }, '⇄ Transponer');
      tr.addEventListener('click', () => {
        st.transpose = !st.transpose;
        tr.setAttribute('aria-pressed', String(st.transpose));
        if (st.transpose) spec.numeric.forEach((c) => st.active.add(c.name));
        render();
      });
      controls.append(tr);
    }
    if (spec.numeric.length > 1) {
      const toggles = el('div', { class: 'series-toggles', role: 'group', 'aria-label': 'Columnas visibles' });
      spec.numeric.slice(0, 24).forEach((c) => {
        const b = el('button', { type: 'button', class: 'chip', 'aria-pressed': String(st.active.has(c.name)) }, c.name);
        b.addEventListener('click', () => {
          if (st.active.has(c.name)) {
            if (st.active.size > 1) st.active.delete(c.name);
          } else st.active.add(c.name);
          b.setAttribute('aria-pressed', String(st.active.has(c.name)));
          render();
        });
        toggles.append(b);
      });
      controls.append(toggles);
    }
    render();
    return card;
  }

  /* ---------- rejilla con la disposición de la hoja de Excel ---------- */
  function layoutItems(items) {
    const sorted = items.slice().sort((a, b) => a.pos.from.row - b.pos.from.row || a.pos.from.col - b.pos.from.col);
    const bands = [];
    for (const it of sorted) {
      const band = bands[bands.length - 1];
      if (band && it.pos.from.row < band.end - 1) {
        band.items.push(it);
        band.end = Math.max(band.end, it.pos.to.row);
      } else bands.push({ end: it.pos.to.row, items: [it] });
    }
    const out = [];
    for (const b of bands) {
      b.items.sort((x, y) => x.pos.from.col - y.pos.from.col);
      const widths = b.items.map((i) => Math.max(1, i.pos.to.col - i.pos.from.col));
      const total = widths.reduce((s, w) => s + w, 0);
      let spans = widths.map((w) => Math.max(3, Math.round((12 * w) / total)));
      let diff = 12 - spans.reduce((s, x) => s + x, 0);
      if (b.items.length <= 4) {
        for (let k = 0; diff !== 0 && k < 24; k++) {
          const idx = diff > 0 ? spans.indexOf(Math.min(...spans)) : spans.indexOf(Math.max(...spans));
          spans[idx] += diff > 0 ? 1 : -1;
          diff += diff > 0 ? -1 : 1;
        }
      }
      b.items.forEach((it, i) => out.push({ it, span: spans[i], height: Math.max(220, Math.min(460, (it.pos.to.row - it.pos.from.row) * 20)) }));
    }
    return out;
  }

  function textlinkValue(link, sheet) {
    const ref = link.replace(/^=/, '');
    const cells = A.resolveRef(ref.includes('!') ? ref : `'${sheet.replace(/'/g, "''")}'!${ref}`, state.wb);
    return cells && cells[0] ? cells[0] : null;
  }

  // Botones y formas con macro asignada -> accesos a secciones
  function macroTarget(macroRef) {
    if (!macroRef) return null;
    const name = macroRef.replace(/^\[\d+\]!|^.*!/, '').split('.').pop().replace(/'/g, '');
    for (const m of state.macros.modules)
      for (const p of m.procedures)
        if (p.name.toLowerCase() === name.toLowerCase()) {
          // si no navega directamente, revisar los procedimientos que llama
          const called = p.calls.map((c) => state.macros.modules.flatMap((mm) => mm.procedures).find((x) => x.name.toLowerCase() === c.toLowerCase())).filter(Boolean);
          const target = p.navigatesTo[0] || (called.find((c) => c.navigatesTo.length) || { navigatesTo: [] }).navigatesTo[0] || (p.sheets.length === 1 ? p.sheets[0] : null);
          return { proc: p, module: m, section: target ? sectionForSheet(target) : null };
        }
    return { proc: null, name, section: null };
  }

  function dashboardKpis(sheet, items) {
    const kpis = [];
    const shapes = items.filter((i) => i.type === 'shape');
    const used = new Set();
    // cuadros de texto vinculados a celdas: el valor viene de la celda, la etiqueta del cuadro vecino
    for (const s of shapes.filter((x) => x.textlink)) {
      const cell = textlinkValue(s.textlink, sheet);
      if (!cell) continue;
      used.add(s);
      const label = shapes
        .filter((o) => o !== s && !o.textlink && o.text && !o.macro && o.text.length < 80)
        .map((o) => {
          const overlapX = Math.min(o.pos.to.col, s.pos.to.col) - Math.max(o.pos.from.col, s.pos.from.col);
          const dy = s.pos.from.row - o.pos.to.row;
          const sameGroup = o.group && o.group === s.group;
          return { o, score: sameGroup ? -10 : overlapX >= 0 ? Math.abs(dy) : 99 };
        })
        .sort((a, b) => a.score - b.score)[0];
      if (label && label.score < 6) used.add(label.o);
      kpis.push({ label: label && label.score < 6 ? label.o.text : s.name, value: cell.v, text: A.cellText(cell), format: cell.z, fromShape: true, r: s.pos.from.row, c: s.pos.from.col });
    }
    const a = state.sheets[sheet];
    if (a) for (const k of a.kpis) if (!kpis.some((x) => x.text === k.text && x.label === k.label)) kpis.push(k);
    return { kpis, used };
  }

  // Dibuja imágenes EMF (tablas pegadas como imagen en Excel) como SVG
  function renderEmf(it) {
    if (it._svg !== undefined) return it._svg ? it._svg.cloneNode(true) : null;
    it._svg = null;
    if (!window.EMF || it.unsupported !== 'EMF' || !it.bytes) return null;
    try {
      const { svg } = EMF.render(it.bytes);
      if (!svg.childNodes.length) return null;
      svg.setAttribute('role', 'img');
      svg.setAttribute('aria-label', it.descr || it.name || 'Imagen del tablero');
      it._svg = svg;
      return svg.cloneNode(true);
    } catch (e) {
      console.warn('No se pudo dibujar la imagen EMF', e);
      return null;
    }
  }

  const isPieSection = (title) =>
    (state.config.seccionesTorta || []).some((x) => Correcciones.fix(String(x)).toLowerCase().trim() === String(title || '').toLowerCase().trim());

  // Secciones de tortas: filas completas de tarjetas (el pictograma ocupa doble ancho)
  function packCards(items, models) {
    const sorted = items.slice().sort((a, b) => a.pos.from.row - b.pos.from.row || a.pos.from.col - b.pos.from.col);
    const unit = (it) => (it.type === 'chart' && genderInfo(models.get(it)) ? 2 : it.type === 'chart' ? 1 : 4);
    const rows = [];
    for (const it of sorted) {
      const u = unit(it);
      const row = rows[rows.length - 1];
      if (row && row.units + u <= 4) {
        row.items.push(it);
        row.units += u;
      } else rows.push({ items: [it], units: u });
    }
    const out = [];
    for (const row of rows) {
      let used = 0;
      row.items.forEach((it, i) => {
        const span = i === row.items.length - 1 ? 12 - used : Math.round((12 * unit(it)) / row.units);
        used += span;
        out.push({ it, span, height: row.units >= 4 ? 380 : 420 });
      });
    }
    return out;
  }

  // Contenido de una parte del Dashboard: indicadores, gráficos en su disposición y tablas
  function renderPart(node, p, sheet, opts) {
    if (p.kpis.length && !opts.skipKpis) node.append(el('div', { class: 'kpi-grid', style: { marginBottom: '16px' } }, ...p.kpis.map(kpiCard)));
    if (p.items.length) {
      if (opts.skipKpis && p.title !== 'General') node.append(el('h2', { class: 'block-title' }, p.title));
      const grid = el('div', { class: 'dash-grid' });
      const models = new Map(p.items.filter((it) => it.type === 'chart').map((it) => [it, chartModelFor(it, sheet, p.title)]));
      for (let { it, span, height } of isPieSection(p.title) ? packCards(p.items, models) : layoutItems(p.items)) {
        let card = null;
        if (it.type === 'chart') {
          const model = models.get(it);
          if (!model) continue;
          if (model.copyOf) {
            card = copyCard(model);
            card.style.gridColumn = `span ${span}`;
            grid.append(card);
            continue;
          }
          const isGender = !!genderInfo(model);
          card = isGender ? genderCard(model) : chartCard(model, { height: height - 70 });
          // el pictograma necesita espacio para las 10 figuras
          if (isGender) span = Math.max(span, 6);
        } else if (it.type === 'image' && it.unsupported && renderEmf(it)) {
          card = el('div', { class: 'card image-card image-card--big emf-card' }, renderEmf(it));
        } else if (it.type === 'image' && it.unsupported) {
          card = el('div', { class: 'card text-card placeholder-card' }, icon('image', 22), el('div', {}, el('b', {}, 'Imagen pegada en el Excel'), el('p', {}, `Esta parte del Dashboard es una imagen ${it.unsupported} (por ejemplo, una tabla pegada como imagen) que el navegador no puede mostrar. Para verla aquí, pégala en Excel como imagen PNG o como tabla con valores.`)));
        } else if (it.type === 'image') {
          card = el('div', { class: 'card image-card image-card--big' }, el('img', { src: it.src, alt: it.descr || it.name || 'Imagen del tablero', loading: 'lazy' }));
        } else {
          card = el('div', { class: 'card text-card' }, it.text);
        }
        card.style.gridColumn = `span ${span}`;
        grid.append(card);
      }
      node.append(grid);
    }
    if (p.notes && p.notes.length) node.append(el('div', { class: 'part-notes' }, ...p.notes.map((n) => el('p', {}, n.text))));
    p.tables.forEach((t) => {
      node.append(el('h2', { class: 'block-title' }, t.title || 'Tabla', el('span', { class: 'muted' }, t.range)));
      node.append(el('div', { class: 'table-block' }, dataTableCard(t, { limit: 15, title: 'Datos' })));
    });
  }

  /* =======================================================================
     Renderizadores de sección
     ======================================================================= */
  const RENDER = {
    dashboard(node, sec) {
      const sheet = sec.sheet;
      const d = state.dash;
      const props = state.wb.Props || {};
      if (state.file.source === 'local' && state.repoFile)
        node.append(el('div', { class: 'notice' }, 'Estás viendo un archivo cargado en este navegador.', el('button', { class: 'btn btn--sm', onclick: useRepoFile }, 'Volver al archivo publicado')));

      if (!d) {
        node.append(sectionHead(sec, 'Resumen', `Resumen generado automáticamente a partir de las hojas de ${state.file.name}.`));
        const kpis = [];
        const autos = [];
        for (const n of state.wb.SheetNames) {
          const a = state.sheets[n];
          if (!a || sheetState(n) !== 'visible') continue;
          for (const t of a.tables) {
            const ks = A.timeKpis(t, A.suggestChart(t));
            if (ks[0] && kpis.length < 8) kpis.push({ ...ks[0], source: n });
          }
          if (a.tables[0] && autos.length < 6) autos.push(a.tables[0]);
        }
        if (kpis.length) node.append(el('h2', { class: 'block-title' }, 'Indicadores clave'), el('div', { class: 'kpi-grid' }, ...kpis.map(kpiCard)));
        if (autos.length) {
          node.append(el('h2', { class: 'block-title' }, 'Panorama por hoja'));
          const grid = el('div', { class: 'dash-grid' });
          autos.forEach((t) => {
            const c = tableChartCard(t);
            if (c) {
              c.style.gridColumn = 'span 6';
              grid.append(c);
            }
          });
          node.append(grid);
        }
        return;
      }

      // Banda de encabezado como la del Excel
      const title = d.header.title || props.Title || sheet;
      node.append(el('header', { class: 'hero' },
        d.logo ? el('img', { class: 'hero__logo', src: d.logo.src, alt: '' }) : null,
        el('div', { class: 'hero__text' },
          el('div', { class: 'eyebrow eyebrow--light' }, 'Resumen del tablero'),
          el('h1', { class: 'hero__title', id: 'h-' + sec.id }, title),
          el('p', { class: 'hero__sub' }, `Hoja «${sheet}» · ${state.file.name}`)),
        d.header.period ? el('div', { class: 'hero__period' }, el('span', {}, 'Corte'), el('b', {}, d.header.period)) : null));

      // Accesos (botones con macro de la hoja)
      const links = [];
      for (const b of d.buttons) {
        const t = macroTarget(b.macro);
        const label = (b.text || b.name || (t && (t.proc ? t.proc.name : t.name)) || 'Macro').split('\n')[0];
        if (t && t.section) links.push(el('a', { class: 'chip chip--macro', href: '#/' + t.section.id, title: 'Ejecutaba la macro ' + (t.proc ? t.proc.name : '') }, icon('chev', 12), label));
        else if (t) links.push(el('a', { class: 'chip chip--macro', href: '#/' + (state.sections.find((s) => s.kind === 'macros') || {}).id, title: 'Ver la macro' }, icon('macro', 13), label));
      }
      if (links.length) node.append(el('div', { class: 'quick-links', 'aria-label': 'Botones del tablero' }, ...links));

      // Indicadores destacados de todas las secciones
      const partOf = (k) => d.parts.find((p) => p.kpis.includes(k));
      const kpis = d.split ? d.headline.map((k) => ({ ...k, source: partOf(k) && partOf(k).title !== 'General' ? partOf(k).title : null })) : d.headline;
      if (kpis.length) {
        node.append(el('h2', { class: 'block-title' }, 'Indicadores clave'));
        node.append(el('div', { class: 'kpi-grid' }, ...kpis.map(kpiCard)));
      }

      if (!d.split) {
        if (kpis.length) node.append(el('div', { style: { height: '16px' } }));
        d.parts.forEach((p) => renderPart(node, p, sheet, { skipKpis: true }));
      } else {
        node.append(el('h2', { class: 'block-title' }, 'Secciones del tablero', el('span', { class: 'muted' }, `${d.parts.length} secciones`)));
        const grid = el('div', { class: 'dash-grid section-cards' });
        for (const s of state.sections.filter((x) => x.kind === 'dashpart')) {
          const p = s.part;
          const nCharts = p.items.filter((x) => x.type === 'chart').length;
          const card = el('article', { class: 'card section-card' });
          card.append(el('a', { class: 'section-card__head', href: '#/' + s.id },
            el('span', { class: 'section-card__icon', 'aria-hidden': 'true' }, icon(s.icon, 20)),
            el('div', {}, el('h3', { class: 'card__title' }, p.title), el('div', { class: 'card__sub' }, [nCharts ? `${nCharts} gráfico${nCharts === 1 ? '' : 's'}` : null, p.tables.length ? `${p.tables.length} tabla${p.tables.length === 1 ? '' : 's'}` : null, p.kpis.length ? `${p.kpis.length} indicador${p.kpis.length === 1 ? '' : 'es'}` : null].filter(Boolean).join(' · '))),
            el('span', { class: 'section-card__go' }, 'Ver', icon('chev', 14))));
          const model = p.items.filter((x) => x.type === 'chart').map((x) => chartModelFor(x, sheet, p.title)).find((m) => m && !m.copyOf);
          if (p.kpis[0] && !model) card.append(el('div', { class: 'section-card__kpi' }, el('b', {}, p.kpis[0].text), el('span', {}, p.kpis[0].label)));
          if (model && genderInfo(model)) {
            card.append(el('div', { class: 'section-card__chart' }, genderBlock(model, true)));
          } else if (model) {
            const holder = el('div', { class: 'section-card__chart' });
            card.append(holder);
            mountChart(holder, { ...model, panels: undefined, ...(model.panels ? model.panels[0] : {}) }, 200);
          } else if (!p.kpis[0]) {
            const img = p.items.find((x) => x.type === 'image');
            const emf = img && img.unsupported ? renderEmf(img) : null;
            if (emf) card.append(el('div', { class: 'image-card emf-card section-card__emf' }, emf));
            else if (img && img.src) card.append(el('div', { class: 'image-card' }, el('img', { src: img.src, alt: '' })));
            else if (p.tables[0]) card.append(el('p', { class: 'muted', style: { padding: '0 16px 16px', margin: 0 } }, p.tables[0].title || 'Tabla de datos'));
          }
          card.style.gridColumn = 'span 4';
          grid.append(card);
        }
        node.append(grid);
      }

      // Otras hojas
      const others = state.sections.filter((s) => s.kind === 'sheet' && s.group === 'sheets');
      if (others.length) {
        node.append(el('h2', { class: 'block-title' }, 'Otras hojas del archivo'));
        node.append(el('div', { class: 'quick-links' }, ...others.map((s) => el('a', { class: 'chip', href: '#/' + s.id }, icon(s.icon, 14), s.title, s.feedsDashboard ? el('span', { class: 'pill pill--brand' }, 'alimenta el tablero') : null))));
      }
    },

    dashpart(node, sec) {
      const d = state.dash;
      const p = sec.part;
      node.append(sectionHead(sec, 'Dashboard' + (d.header.period ? ' · ' + d.header.period : ''), d.header.title ? `${d.header.title} — hoja «${sec.sheet}»` : `Hoja «${sec.sheet}»`, null, sec.icon));
      renderPart(node, p, sec.sheet, {});
      const parts = state.sections.filter((x) => x.kind === 'dashpart');
      const i = parts.indexOf(sec);
      const prev = parts[i - 1], next = parts[i + 1];
      node.append(el('nav', { class: 'pager', 'aria-label': 'Secciones del tablero' },
        prev ? el('a', { class: 'btn', href: '#/' + prev.id }, '← ' + prev.title) : el('span'),
        next ? el('a', { class: 'btn btn--primary', href: '#/' + next.id }, next.title + ' →') : el('a', { class: 'btn', href: '#/' + state.sections[0].id }, 'Volver al resumen')));
    },

    sheet(node, sec) {
      const a = state.sheets[sec.sheet];
      const items = state.parts && state.parts.sheets[sec.sheet] ? state.parts.sheets[sec.sheet].items : [];
      const charts = items.filter((i) => i.type === 'chart');
      const bits = [];
      if (a.tables.length) bits.push(`${a.tables.length} tabla${a.tables.length === 1 ? '' : 's'}`);
      if (charts.length) bits.push(`${charts.length} gráfico${charts.length === 1 ? '' : 's'} de Excel`);
      if (sec.group === 'hidden') bits.push('hoja oculta en el archivo');
      const extra = el('div', { class: 'quick-links' });
      if (sec.feedsDashboard) extra.append(el('span', { class: 'pill pill--brand' }, 'Alimenta el tablero'));
      if (sec.navTarget) extra.append(el('span', { class: 'pill pill--gold' }, 'Destino de un botón con macro'));
      node.append(sectionHead(sec, 'Hoja', bits.join(' · '), extra.childNodes.length ? extra : null));

      if (sec.macroRefs && sec.macroRefs.length) {
        const names = Array.from(new Set(sec.macroRefs.map((r) => r.proc.name)));
        const macrosId = (state.sections.find((s) => s.kind === 'macros') || {}).id;
        node.append(el('div', { class: 'notice' }, icon('macro', 16), el('span', {}, 'Macros que trabajan con esta hoja:'), ...names.slice(0, 8).map((n) => el('a', { class: 'chip chip--macro', href: '#/' + macrosId }, n))));
      }
      if (a.truncated) node.append(el('div', { class: 'notice' }, 'La hoja es muy grande: se analizaron las primeras 6.000 filas y 160 columnas.'));

      if (charts.length) {
        node.append(el('h2', { class: 'block-title' }, 'Gráficos de la hoja'));
        const grid = el('div', { class: 'dash-grid' });
        for (const { it, span, height } of layoutItems(charts)) {
          const model = withCenterTotal(excelChartModel(it, sec.sheet));
          if (!model) continue;
          const card = model.copyOf ? copyCard(model) : chartCard(model, { height: height - 70 });
          card.style.gridColumn = `span ${span}`;
          grid.append(card);
        }
        node.append(grid);
      }

      if (a.kpis.length) {
        node.append(el('h2', { class: 'block-title' }, 'Indicadores'));
        node.append(el('div', { class: 'kpi-grid' }, ...a.kpis.slice(0, 12).map(kpiCard)));
      }

      a.tables.slice(0, 15).forEach((t, i) => {
        const block = el('div', { class: 'table-block' });
        node.append(el('h2', { class: 'block-title' }, t.title || (a.tables.length > 1 ? `Tabla ${i + 1}` : 'Datos'), el('span', { class: 'muted' }, t.range)));
        const spec = A.suggestChart(t);
        const tk = A.timeKpis(t, spec);
        if (tk.length) block.append(el('div', { class: 'kpi-grid' }, ...tk.map(kpiCard)));
        const grid = el('div', { class: 'table-block__grid' });
        const coveredByExcel = charts.length && spec && spec.numeric.length <= 2 && charts.some((c) => c.chart.groups.some((g) => g.series.some((s) => s.val && s.val.f && A.refSheets(s.val.f).includes(sec.sheet))));
        const chart = coveredByExcel ? null : tableChartCard(t);
        if (chart) grid.append(chart);
        grid.append(dataTableCard(t, { title: 'Tabla de datos', limit: chart ? 12 : 25 }));
        block.append(grid);
        node.append(block);
      });
      if (a.tables.length > 15) node.append(el('p', { class: 'muted' }, `Hay ${a.tables.length - 15} tablas más en esta hoja que no se muestran.`));
      if (!a.tables.length && !charts.length && !a.kpis.length) node.append(el('p', { class: 'muted' }, 'Esta hoja no tiene tablas con datos numéricos.'));
    },

    macros(node, sec) {
      const mods = state.macros.modules;
      const procs = mods.flatMap((m) => m.procedures.map((p) => ({ ...p, module: m })));
      node.append(sectionHead(sec, 'Automatización', 'Lo que hacen las macros del archivo, qué hojas usan y cómo se reflejan en las secciones de esta página. Las macros no se ejecutan en el navegador: sus resultados ya están en los datos del archivo.'));
      if (!state.macros.available) {
        node.append(el('div', { class: 'card', style: { padding: '18px' } }, el('p', { style: { margin: 0 } }, 'Este archivo no contiene macros. Si tu libro original es .xlsm (o .xls con macros), súbelo en ese formato para ver aquí el detalle.')));
        return;
      }
      if (state.macros.error) node.append(el('div', { class: 'notice notice--error' }, 'Las macros no se pudieron leer por completo: ' + state.macros.error));
      const sheetsTouched = new Set(procs.flatMap((p) => p.sheets));
      node.append(el('div', { class: 'stat-row' },
        el('div', { class: 'card stat' }, el('b', {}, mods.length), el('span', {}, 'módulos')),
        el('div', { class: 'card stat' }, el('b', {}, procs.length), el('span', {}, 'procedimientos')),
        el('div', { class: 'card stat' }, el('b', {}, sheetsTouched.size), el('span', {}, 'hojas usadas')),
        el('div', { class: 'card stat' }, el('b', {}, procs.filter((p) => p.isEvent).length), el('span', {}, 'automáticas (eventos)'))));

      // Controles de las hojas (botones, listas) y su macro
      const controls = [];
      for (const [sheet, info] of Object.entries((state.parts && state.parts.sheets) || {}))
        for (const it of info.items)
          if (it.macro || it.type === 'control') controls.push({ sheet, it });
      if (controls.length) {
        node.append(el('h2', { class: 'block-title' }, 'Botones y controles en las hojas'));
        const t = el('table', { class: 'data controls-table' });
        t.append(el('thead', {}, el('tr', {}, el('th', {}, 'Hoja'), el('th', {}, 'Control'), el('th', {}, 'Texto'), el('th', {}, 'Macro / vínculo'), el('th', {}, 'En esta página'))));
        const tb = el('tbody');
        for (const { sheet, it } of controls) {
          const tgt = macroTarget(it.macro);
          const kind = it.type === 'control' ? { Button: 'Botón', Drop: 'Lista desplegable', Checkbox: 'Casilla', Spin: 'Control de número', List: 'Cuadro de lista', Radio: 'Botón de opción', Scroll: 'Barra de desplazamiento' }[it.control] || it.control : it.type === 'image' ? 'Imagen con macro' : 'Forma con macro';
          tb.append(el('tr', {},
            el('td', {}, sheet),
            el('td', {}, kind),
            el('td', {}, (it.text || it.name || '').split('\n')[0]),
            el('td', {}, it.macro ? el('code', { class: 'inline' }, it.macro.replace(/^\[\d+\]!/, '')) : it.link ? 'Celda ' + it.link + (it.range ? ' (opciones: ' + it.range + ')' : '') : '—'),
            el('td', {}, tgt && tgt.section ? el('a', { href: '#/' + tgt.section.id }, 'Sección ' + tgt.section.title) : '—')));
        }
        t.append(tb);
        node.append(el('div', { class: 'card' }, el('div', { class: 'table-scroll' }, t)));
      }

      node.append(el('h2', { class: 'block-title' }, 'Módulos'));
      const kindLabel = { standard: 'Módulo', sheet: 'Código de hoja', class: 'Clase / libro' };
      mods.forEach((m, mi) => {
        const det = el('details', { class: 'card module' });
        if (mi === 0) det.open = true;
        det.append(el('summary', {}, el('span', { class: 'chev' }, icon('chev', 16)), el('span', { class: 'module__name' }, m.name), el('span', { class: 'pill' }, kindLabel[m.kind] || 'Módulo'), m.owner ? el('span', { class: 'pill pill--brand' }, 'Hoja ' + m.owner) : null, el('span', { class: 'muted', style: { marginLeft: 'auto', fontSize: '12.5px' } }, `${m.procedures.length} procedimiento${m.procedures.length === 1 ? '' : 's'}`)));
        const body = el('div', { class: 'module__body' });
        if (!m.procedures.length) body.append(el('p', { class: 'muted', style: { margin: 0 } }, 'Solo contiene declaraciones.'));
        for (const p of m.procedures) {
          const box = el('div', { class: 'proc' });
          box.append(el('div', { class: 'proc__head' }, el('span', { class: 'proc__name' }, p.name), el('span', { class: 'pill' }, p.kind), p.isEvent ? el('span', { class: 'pill pill--gold' }, 'Se ejecuta automáticamente') : null, el('span', { class: 'muted', style: { fontSize: '12px' } }, `${p.lineCount} líneas`)));
          box.append(el('p', { class: 'proc__desc' }, describeProc(p)));
          if (p.actions.length) box.append(el('div', { class: 'proc__tags' }, ...p.actions.map((x) => el('span', { class: 'pill pill--brand' }, x))));
          const refs = el('div', { class: 'proc__refs' });
          if (p.sheets.length) {
            refs.append(el('span', {}, 'Hojas:'));
            p.sheets.forEach((s) => {
              const sx = sectionForSheet(s) || (s === state.dashSheet ? state.sections[0] : null);
              refs.append(sx ? el('a', { class: 'chip', href: '#/' + sx.id }, s) : el('span', { class: 'pill' }, s));
            });
          }
          if (p.pivots.length) refs.append(el('span', {}, 'Tablas dinámicas: ' + p.pivots.join(', ')));
          if (p.ranges.length) refs.append(el('span', {}, 'Rangos: ' + p.ranges.join(', ')));
          if (p.calls.length) refs.append(el('span', {}, 'Llama a: ' + p.calls.join(', ')));
          if (refs.childNodes.length) box.append(refs);
          const btn = el('button', { class: 'btn btn--sm', style: { marginTop: '10px' } }, 'Ver código');
          let pre = null;
          btn.addEventListener('click', () => {
            if (pre) {
              pre.remove();
              pre = null;
              btn.textContent = 'Ver código';
              return;
            }
            pre = el('pre', { class: 'code', html: highlightVba(p.code) });
            box.append(pre);
            btn.textContent = 'Ocultar código';
          });
          box.append(btn);
          body.append(box);
        }
        det.append(body);
        node.append(det);
      });
    },

    file(node, sec) {
      const wb = state.wb;
      const p = wb.Props || {};
      const f = state.file;
      node.append(sectionHead(sec, 'Archivo', 'Origen de los datos y cómo mantener la página al día.'));
      const info = el('dl', { class: 'info-list' });
      const add = (k, v) => v && info.append(el('dt', {}, k), el('dd', {}, v));
      add('Archivo', f.name);
      add('Tamaño', (f.size / 1024 / 1024).toFixed(2).replace('.', ',') + ' MB');
      add('Origen', f.source === 'repo' ? 'Publicado en la carpeta data/ del sitio' : 'Cargado en este navegador');
      add('Cargado', new Date(f.loadedAt).toLocaleString('es-CO'));
      add('Título', p.Title);
      add('Autor', p.Author);
      add('Última modificación', p.ModifiedDate ? new Date(p.ModifiedDate).toLocaleString('es-CO') : p.LastAuthor ? '' : null);
      add('Modificado por', p.LastAuthor);
      node.append(el('div', { class: 'card', style: { padding: '18px' } }, info));

      node.append(el('h2', { class: 'block-title' }, 'Hojas'));
      const t = el('table', { class: 'data' });
      t.append(el('thead', {}, el('tr', {}, el('th', {}, 'Hoja'), el('th', {}, 'Estado'), el('th', { class: 'num' }, 'Tablas'), el('th', { class: 'num' }, 'Gráficos'), el('th', { class: 'num' }, 'Indicadores'), el('th', {}, 'Sección'))));
      const tb = el('tbody');
      for (const n of wb.SheetNames) {
        const a = state.sheets[n] || { tables: [], kpis: [] };
        const items = state.parts && state.parts.sheets[n] ? state.parts.sheets[n].items : [];
        const sx = n === state.dashSheet ? state.sections[0] : sectionForSheet(n);
        tb.append(el('tr', {}, el('td', {}, n), el('td', {}, { visible: 'Visible', hidden: 'Oculta', veryHidden: 'Muy oculta' }[sheetState(n)] || 'Visible'), el('td', { class: 'num' }, a.tables.length), el('td', { class: 'num' }, items.filter((i) => i.type === 'chart').length), el('td', { class: 'num' }, a.kpis.length), el('td', {}, sx ? el('a', { href: '#/' + sx.id }, sx.title) : '—')));
      }
      t.append(tb);
      node.append(el('div', { class: 'card' }, el('div', { class: 'table-scroll' }, t)));

      const names = ((wb.Workbook && wb.Workbook.Names) || []).filter((n) => !/^_xlnm|^_/.test(n.Name));
      if (names.length) {
        node.append(el('h2', { class: 'block-title' }, 'Nombres definidos'));
        const nt = el('table', { class: 'data' });
        nt.append(el('thead', {}, el('tr', {}, el('th', {}, 'Nombre'), el('th', {}, 'Referencia'))));
        nt.append(el('tbody', {}, ...names.slice(0, 200).map((n) => el('tr', {}, el('td', {}, n.Name), el('td', {}, el('code', { class: 'inline' }, n.Ref))))));
        node.append(el('div', { class: 'card' }, el('div', { class: 'table-scroll' }, nt)));
      }

      node.append(el('h2', { class: 'block-title' }, 'Cómo actualizar la página'));
      node.append(el('div', { class: 'card', style: { padding: '18px 20px' } },
        el('ol', { class: 'steps' },
          el('li', {}, 'Actualiza el Excel como siempre (ejecuta las macros y guarda el archivo en Excel para que los gráficos queden con los valores al día).'),
          el('li', {}, el('b', {}, 'Rápido, solo en tu equipo: '), 'pulsa «Subir Excel» (o arrastra el archivo sobre la página). Se guarda en este navegador para la próxima visita.'),
          el('li', {}, el('b', {}, 'Para todos los que abren la página: '), 'reemplaza el archivo en la carpeta ', el('code', { class: 'inline' }, 'data/'), ' del repositorio con el mismo nombre que indica ', el('code', { class: 'inline' }, 'data/config.json'), '.')),
        el('div', { style: { display: 'flex', gap: '8px', marginTop: '14px', flexWrap: 'wrap' } },
          el('button', { class: 'btn btn--primary', onclick: () => $('#file-input').click() }, icon('upload', 16), 'Subir nuevo Excel'),
          f.source === 'local' ? el('button', { class: 'btn', onclick: forgetLocal }, 'Olvidar archivo de este navegador') : null)));
    },
  };

  function describeProc(p) {
    const parts = [];
    if (p.isEvent) {
      const ev = {
        Workbook_Open: 'al abrir el libro', Auto_Open: 'al abrir el libro', Worksheet_Change: 'cuando cambia una celda de la hoja', Worksheet_Activate: 'al entrar a la hoja',
        Worksheet_SelectionChange: 'al seleccionar celdas', Worksheet_Calculate: 'al recalcular la hoja', Workbook_BeforeClose: 'antes de cerrar el libro', Workbook_SheetChange: 'cuando cambia cualquier hoja',
        Worksheet_PivotTableUpdate: 'al actualizar una tabla dinámica', Workbook_BeforeSave: 'antes de guardar',
      };
      const key = Object.keys(ev).find((k) => k.toLowerCase() === p.name.toLowerCase());
      parts.push('Se ejecuta ' + (key ? ev[key] : 'automáticamente') + '.');
    }
    if (p.navigatesTo.length) parts.push('Lleva a ' + (p.navigatesTo.length === 1 ? 'la hoja «' + p.navigatesTo[0] + '»' : 'las hojas «' + p.navigatesTo.join('», «') + '»') + '.');
    const other = p.sheets.filter((s) => !p.navigatesTo.includes(s));
    if (other.length) parts.push('Trabaja con ' + (other.length === 1 ? 'la hoja «' + other[0] + '»' : 'las hojas «' + other.join('», «') + '»') + '.');
    if (p.messages.length) parts.push('Muestra el mensaje: «' + p.messages[0] + '».');
    if (!parts.length) parts.push(p.actions.length ? 'Realiza: ' + p.actions.join(', ').toLowerCase() + '.' : 'Procedimiento auxiliar.');
    return parts.join(' ');
  }

  function highlightVba(code) {
    const KW = /\b(Sub|Function|End|If|Then|Else|ElseIf|For|Each|Next|To|Step|Do|Loop|While|Wend|Until|With|Dim|As|Set|Private|Public|Const|Call|Exit|Select|Case|Not|And|Or|In|Is|Nothing|True|False|On|Error|Resume|GoTo|Static|Option|Explicit|Long|Integer|String|Double|Boolean|Variant|Object|Range|Worksheet|Workbook|New|ByVal|ByRef|Optional|Property|Get|Let|ReDim|Preserve|Me)\b/g;
    return code
      .split('\n')
      .map((line) => {
        let out = '', inStr = false, i = 0, buf = '';
        const flush = (cls) => {
          if (!buf) return;
          out += cls ? `<span class="${cls}">${esc(buf)}</span>` : esc(buf).replace(KW, '<span class="k">$1</span>').replace(/\b(\d+(\.\d+)?)\b/g, '<span class="n">$1</span>');
          buf = '';
        };
        for (; i < line.length; i++) {
          const ch = line[i];
          if (!inStr && ch === "'") {
            flush();
            buf = line.slice(i);
            flush('c');
            return out;
          }
          if (ch === '"') {
            if (!inStr) {
              flush();
              inStr = true;
              buf = '"';
            } else {
              buf += '"';
              inStr = false;
              flush('s');
            }
            continue;
          }
          buf += ch;
        }
        flush(inStr ? 's' : null);
        return out;
      })
      .join('\n');
  }

  /* =======================================================================
     Carga del archivo
     ======================================================================= */
  function setLoading(text) {
    $('#content').innerHTML = '';
    $('#content').append(el('div', { class: 'loading' }, el('div', { class: 'spinner' }), text));
  }

  function toast(msg) {
    const t = el('div', { class: 'toast', role: 'status' }, msg);
    document.body.append(t);
    setTimeout(() => t.remove(), 3200);
  }

  async function loadBuffer(buffer, meta) {
    setLoading('Leyendo el archivo ' + meta.name + '…');
    await new Promise((r) => setTimeout(r, 30));
    const wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true, cellNF: true, bookVBA: true, WTF: false });
    state.wb = wb;
    try {
      state.parts = await OOXML.readWorkbookParts(buffer);
    } catch (e) {
      console.warn('No se pudieron leer los gráficos', e);
      state.parts = null;
    }
    // ortografía de cuadros de texto y títulos de gráficos
    if (state.parts)
      for (const info of Object.values(state.parts.sheets))
        for (const it of info.items) {
          if (it.text) it.text = Correcciones.fix(it.text);
          if (it.chart && it.chart.title && it.chart.title.text) it.chart.title.text = Correcciones.fix(it.chart.title.text);
          if (it.chart) it.chart.axes.forEach((a) => a.title && (a.title = Correcciones.fix(a.title)));
        }
    if (state.parts) for (const info of Object.values(state.parts.sheets)) markChartCopies(info.items);
    state.sheets = {};
    for (const n of wb.SheetNames) {
      const ws = wb.Sheets[n];
      state.sheets[n] = ws ? A.analyzeSheet(ws, n) : { name: n, tables: [], kpis: [], texts: [] };
    }
    state.macros = VBA.readMacros(wb, buffer);
    state.file = { ...meta, loadedAt: meta.loadedAt || Date.now() };
    for (const [, e] of state.charts) e.chart && e.chart.destroy();
    state.charts.clear();
    state.rendered.clear();
    buildSections();
    $('#content').innerHTML = '';
    renderNav();
    updateHeader();
    $('#app').hidden = false;
    $('#empty').hidden = true;
    showSection(currentId());
  }

  function updateHeader() {
    const f = state.file;
    const chip = $('#file-chip');
    chip.hidden = !f;
    if (!f) return;
    const p = state.wb.Props || {};
    const when = p.ModifiedDate ? new Date(p.ModifiedDate) : f.lastModified ? new Date(f.lastModified) : null;
    chip.innerHTML = '';
    chip.append(icon('file', 14), el('strong', { title: f.name }, f.name), when ? el('span', {}, '· ' + when.toLocaleDateString('es-CO', { day: 'numeric', month: 'short', year: 'numeric' })) : null);
    $('#foot-file').textContent = f.name;
  }

  async function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xlsm|xlsb|xls|xltx|xltm)$/i.test(file.name)) {
      toast('Formato no compatible: sube un archivo de Excel (.xlsx, .xlsm, .xls)');
      return;
    }
    const buffer = await file.arrayBuffer();
    const meta = { name: file.name, size: file.size, lastModified: file.lastModified, source: 'local', loadedAt: Date.now() };
    try {
      await loadBuffer(buffer, meta);
      await store.put({ ...meta, buffer });
      toast('Tablero actualizado con ' + file.name);
    } catch (e) {
      console.error(e);
      showError('No se pudo leer el archivo: ' + e.message);
    }
  }

  async function fetchRepoFile() {
    try {
      const cfgRes = await fetch('data/config.json', { cache: 'no-store' });
      if (cfgRes.ok) Object.assign(state.config, await cfgRes.json());
    } catch (e) {
      /* sin configuración */
    }
    applyConfig();
    if (!state.config.archivo) return null;
    try {
      const res = await fetch('data/' + encodeURIComponent(state.config.archivo).replace(/%2F/g, '/'), { cache: 'no-store' });
      if (!res.ok) return null;
      const buffer = await res.arrayBuffer();
      const lm = res.headers.get('last-modified');
      return { buffer, meta: { name: state.config.archivo.split('/').pop(), size: buffer.byteLength, lastModified: lm ? Date.parse(lm) : null, source: 'repo' } };
    } catch (e) {
      return null;
    }
  }

  async function useRepoFile() {
    if (!state.repoFile) return;
    await store.clear();
    await loadBuffer(state.repoFile.buffer, state.repoFile.meta);
  }

  async function forgetLocal() {
    await store.clear();
    if (state.repoFile) await loadBuffer(state.repoFile.buffer, state.repoFile.meta);
    else location.reload();
  }

  function applyConfig() {
    const c = state.config;
    if (c.correcciones && typeof c.correcciones === 'object') Correcciones.set(c.correcciones);
    $('#brand-org').textContent = c.organizacion || 'Sector palmero';
    $('#brand-title').textContent = c.titulo || 'Tablero de indicadores';
    document.title = c.titulo || 'Tablero de indicadores';
  }

  function showError(msg) {
    $('#content').innerHTML = '';
    $('#content').append(el('div', { class: 'notice notice--error' }, msg));
  }

  /* ---------- eventos globales ---------- */
  function bindUi() {
    const input = $('#file-input');
    input.addEventListener('change', () => {
      handleFile(input.files[0]);
      input.value = '';
    });
    document.querySelectorAll('[data-action="upload"]').forEach((b) => b.addEventListener('click', () => input.click()));
    $('#btn-print').addEventListener('click', () => {
      // mostrar todas las secciones antes de imprimir
      state.sections.forEach((s) => {
        if (!state.rendered.has(s.id)) {
          const prev = currentId();
          showSection(s.id);
          showSection(prev);
        }
      });
      setTimeout(() => window.print(), 300);
    });
    const themeBtn = $('#btn-theme');
    const syncThemeIcon = () => {
      themeBtn.replaceChildren(icon(isDark() ? 'sun' : 'moon', 18));
      themeBtn.setAttribute('aria-label', isDark() ? 'Cambiar a tema claro' : 'Cambiar a tema oscuro');
      themeBtn.title = themeBtn.getAttribute('aria-label');
    };
    themeBtn.addEventListener('click', () => {
      applyTheme(isDark() ? 'light' : 'dark');
      syncThemeIcon();
      rebuildCharts();
    });
    syncThemeIcon();
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
      syncThemeIcon();
      rebuildCharts();
    });
    const palBtn = $('#btn-palette');
    const syncPal = () => {
      const excel = useExcelColors();
      palBtn.title = excel ? 'Usando colores del Excel · cambiar a paleta verde' : 'Usando paleta verde · cambiar a colores del Excel';
      palBtn.setAttribute('aria-label', palBtn.title);
      palBtn.setAttribute('aria-pressed', String(excel));
    };
    palBtn.addEventListener('click', () => {
      prefs.set('colores', useExcelColors() ? 'verde' : 'excel');
      syncPal();
      rebuildCharts();
      toast(useExcelColors() ? 'Colores originales del Excel' : 'Paleta verde');
    });
    syncPal();

    window.addEventListener('hashchange', () => state.wb && showSection(currentId()));

    // arrastrar y soltar en cualquier parte
    const overlay = $('#drop-overlay');
    let depth = 0;
    window.addEventListener('dragenter', (e) => {
      if (!e.dataTransfer || !Array.from(e.dataTransfer.types).includes('Files')) return;
      depth++;
      overlay.hidden = false;
    });
    window.addEventListener('dragleave', () => {
      depth = Math.max(0, depth - 1);
      if (!depth) overlay.hidden = true;
    });
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => {
      e.preventDefault();
      depth = 0;
      overlay.hidden = true;
      const f = e.dataTransfer && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });
    const dz = $('#dropzone');
    dz.addEventListener('click', () => input.click());
    dz.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        input.click();
      }
    });
  }

  async function init() {
    applyTheme(prefs.get('tema', 'auto'));
    if (window.Chart) {
      Chart.defaults.font.family = getComputedStyle(document.documentElement).getPropertyValue('--font');
      Chart.defaults.font.size = 12;
    }
    bindUi();
    const [repo, local] = await Promise.all([fetchRepoFile(), store.get()]);
    state.repoFile = repo;
    try {
      if (local && local.buffer) await loadBuffer(local.buffer, local);
      else if (repo) await loadBuffer(repo.buffer, repo.meta);
      else {
        $('#empty').hidden = false;
        $('#app').hidden = true;
      }
    } catch (e) {
      console.error(e);
      await store.clear();
      $('#empty').hidden = false;
      $('#app').hidden = true;
      toast('No se pudo abrir el último archivo: ' + e.message);
    }
  }

  // expuesto para pruebas y depuración
  window.Tablero = { state, loadBuffer, handleFile };
  document.addEventListener('DOMContentLoaded', init);
})();

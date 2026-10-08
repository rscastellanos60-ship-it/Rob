/*
 * emf.js — Convierte imágenes EMF a SVG.
 * Excel guarda como EMF las tablas copiadas "como imagen"; el navegador no
 * sabe mostrarlas. Este intérprete cubre los registros que usa Excel:
 * rectángulos rellenos (BITBLT), líneas, polígonos y texto.
 */
(function (global) {
  'use strict';

  const NS = 'http://www.w3.org/2000/svg';
  const color = (dv, o) => `rgb(${dv.getUint8(o)},${dv.getUint8(o + 1)},${dv.getUint8(o + 2)})`;
  const STOCK = {
    0: { kind: 'brush', color: '#ffffff' },
    1: { kind: 'brush', color: '#c0c0c0' },
    2: { kind: 'brush', color: '#808080' },
    3: { kind: 'brush', color: '#404040' },
    4: { kind: 'brush', color: '#000000' },
    5: { kind: 'brush', none: true },
    6: { kind: 'pen', color: '#ffffff', width: 1 },
    7: { kind: 'pen', color: '#000000', width: 1 },
    8: { kind: 'pen', none: true },
    10: { kind: 'font', size: 12, weight: 400, face: 'sans-serif' },
    13: { kind: 'font', size: 12, weight: 400, face: 'sans-serif' },
  };

  function render(bytes) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (dv.getUint32(0, true) !== 1) throw new Error('No es un EMF');
    const b = [8, 12, 16, 20].map((o) => dv.getInt32(o, true));
    const W = Math.max(1, b[2] - b[0] + 1), H = Math.max(1, b[3] - b[1] + 1);
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', `${b[0]} ${b[1]} ${W} ${H}`);
    svg.setAttribute('xmlns', NS);
    const objects = {};
    let st = {
      pen: STOCK[7], brush: STOCK[0], font: STOCK[13], text: 'rgb(0,0,0)', align: 0,
      wOrg: [0, 0], vOrg: [0, 0], wExt: [1, 1], vExt: [1, 1], cur: [0, 0],
      xf: [1, 0, 0, 1, 0, 0], // transformación de mundo (eM11, eM12, eM21, eM22, eDx, eDy)
      clip: null,
    };
    const stack = [];
    let path = null;
    // mundo -> página -> dispositivo
    const P = (x, y) => {
      const m = st.xf;
      const px = x * m[0] + y * m[2] + m[4], py = x * m[1] + y * m[3] + m[5];
      return [(px - st.wOrg[0]) * (st.vExt[0] / st.wExt[0]) + st.vOrg[0], (py - st.wOrg[1]) * (st.vExt[1] / st.wExt[1]) + st.vOrg[1]];
    };
    const X = (x, y = 0) => P(x, y)[0];
    const Y = (y, x = 0) => P(x, y)[1];
    const scale = () => Math.hypot(st.xf[2], st.xf[3]) * Math.abs(st.vExt[1] / st.wExt[1]);
    const readXf = (o) => [0, 4, 8, 12, 16, 20].map((k) => dv.getFloat32(o + k, true));
    const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
    const rectPath = (l, t, r, b2) => {
      const pts = [P(l, t), P(r, t), P(r, b2), P(l, b2)];
      return pts.map((p) => p.join(',')).join(' ');
    };
    // Mapa de bits DIB -> imagen PNG embebida
    const dib = (offBmi, cbBmi, offBits, x, y, cx, cy) => {
      if (!cbBmi || typeof document === 'undefined') return;
      const bo = offBmi, w = dv.getInt32(bo + 4, true), h = dv.getInt32(bo + 8, true), bpp = dv.getUint16(bo + 14, true), comp = dv.getUint32(bo + 16, true);
      if (comp !== 0 && comp !== 3) return;
      const aw = Math.abs(w), ah = Math.abs(h);
      if (!aw || !ah || aw * ah > 4e6) return;
      const hdr = dv.getUint32(bo, true);
      const nPal = bpp <= 8 ? dv.getUint32(bo + 32, true) || 1 << bpp : 0;
      const palO = bo + hdr + (comp === 3 && hdr === 40 ? 12 : 0);
      const stride = Math.floor((aw * bpp + 31) / 32) * 4;
      const canvas = document.createElement('canvas');
      canvas.width = aw;
      canvas.height = ah;
      const ctx = canvas.getContext('2d');
      const img = ctx.createImageData(aw, ah);
      for (let row = 0; row < ah; row++) {
        const src = offBits + (h > 0 ? ah - 1 - row : row) * stride;
        for (let col = 0; col < aw; col++) {
          let r = 0, g = 0, b3 = 0;
          if (bpp === 24 || bpp === 32) {
            const q = src + col * (bpp / 8);
            if (q + 2 >= dv.byteLength) continue;
            b3 = dv.getUint8(q); g = dv.getUint8(q + 1); r = dv.getUint8(q + 2);
          } else if (bpp <= 8) {
            const bit = col * bpp, byte = dv.getUint8(src + (bit >> 3));
            const idx = (byte >> (8 - bpp - (bit & 7))) & ((1 << bpp) - 1);
            const po = palO + Math.min(idx, nPal - 1) * 4;
            b3 = dv.getUint8(po); g = dv.getUint8(po + 1); r = dv.getUint8(po + 2);
          } else return;
          const d = (row * aw + col) * 4;
          img.data[d] = r; img.data[d + 1] = g; img.data[d + 2] = b3; img.data[d + 3] = 255;
        }
      }
      ctx.putImageData(img, 0, 0);
      const [x1, y1] = P(x, y), [x2, y2] = P(x + cx, y + cy);
      add('image', { x: Math.min(x1, x2), y: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1), href: canvas.toDataURL('image/png'), preserveAspectRatio: 'none' });
    };
    // recorte: Excel limita cada celda/relleno con regiones de recorte
    const defs = document.createElementNS(NS, 'defs');
    svg.appendChild(defs);
    const clipIds = new Map();
    const clipRef = () => {
      const c = st.clip;
      if (!c) return null;
      const key = c.map((v) => Math.round(v * 10) / 10).join(',');
      if (!clipIds.has(key)) {
        const id = 'c' + Math.random().toString(36).slice(2, 8) + clipIds.size;
        const cp = document.createElementNS(NS, 'clipPath');
        cp.setAttribute('id', id);
        cp.setAttribute('clipPathUnits', 'userSpaceOnUse');
        const r = document.createElementNS(NS, 'rect');
        r.setAttribute('x', c[0]);
        r.setAttribute('y', c[1]);
        r.setAttribute('width', Math.max(0, c[2] - c[0]));
        r.setAttribute('height', Math.max(0, c[3] - c[1]));
        cp.appendChild(r);
        defs.appendChild(cp);
        clipIds.set(key, id);
      }
      return clipIds.get(key);
    };
    const intersect = (a, b2) => (a ? [Math.max(a[0], b2[0]), Math.max(a[1], b2[1]), Math.min(a[2], b2[2]), Math.min(a[3], b2[3])] : b2);
    const add = (tag, attrs, text) => {
      const n = document.createElementNS(NS, tag);
      for (const k in attrs) n.setAttribute(k, attrs[k]);
      const cid = clipRef();
      if (cid) n.setAttribute('clip-path', `url(#${cid})`);
      if (text != null) n.textContent = text;
      svg.appendChild(n);
    };
    // grosor del lápiz: los geométricos están en unidades lógicas (se escalan); los cosméticos miden 1 px
    const stroke = () => {
      if (st.pen.none) return { stroke: 'none' };
      const w = st.pen.geometric ? (st.pen.width || 1) * Math.hypot(st.xf[0], st.xf[1]) * Math.abs(st.vExt[0] / st.wExt[0]) : Math.max(1, (st.pen.width || 1) * Math.hypot(st.xf[0], st.xf[1]));
      return { stroke: st.pen.color, 'stroke-width': +Math.max(0.5, w).toFixed(2) };
    };
    const fill = () => (st.brush.none ? 'none' : st.brush.color);
    const points16 = (o, count) => {
      const pts = [];
      for (let i = 0; i < count; i++) pts.push(P(dv.getInt16(o + i * 4, true), dv.getInt16(o + i * 4 + 2, true)).join(','));
      return pts.join(' ');
    };

    let off = 0, guard = 0;
    while (off + 8 <= dv.byteLength && guard++ < 200000) {
      const type = dv.getUint32(off, true), size = dv.getUint32(off + 4, true);
      if (size < 8 || off + size > dv.byteLength) break;
      const o = off;
      switch (type) {
        case 9: st.wExt = [dv.getInt32(o + 8, true) || 1, dv.getInt32(o + 12, true) || 1]; break; // SETWINDOWEXTEX
        case 10: st.wOrg = [dv.getInt32(o + 8, true), dv.getInt32(o + 12, true)]; break; // SETWINDOWORGEX
        case 11: st.vExt = [dv.getInt32(o + 8, true) || 1, dv.getInt32(o + 12, true) || 1]; break; // SETVIEWPORTEXTEX
        case 12: st.vOrg = [dv.getInt32(o + 8, true), dv.getInt32(o + 12, true)]; break; // SETVIEWPORTORGEX
        case 22: st.align = dv.getUint32(o + 8, true); break; // SETTEXTALIGN
        case 24: st.text = color(dv, o + 8); break; // SETTEXTCOLOR
        case 27: // MOVETOEX
          st.cur = [dv.getInt32(o + 8, true), dv.getInt32(o + 12, true)];
          if (path) path.push('M' + P(...st.cur).join(','));
          break;
        case 35: st.xf = readXf(o + 8); break; // SETWORLDTRANSFORM
        case 36: { // MODIFYWORLDTRANSFORM
          const m = readXf(o + 8), mode = dv.getUint32(o + 32, true);
          st.xf = mode === 1 ? [1, 0, 0, 1, 0, 0] : mode === 2 ? mul(m, st.xf) : mode === 3 ? mul(st.xf, m) : m;
          break;
        }
        case 30: { // INTERSECTCLIPRECT (unidades lógicas)
          const [x1, y1] = P(dv.getInt32(o + 8, true), dv.getInt32(o + 12, true));
          const [x2, y2] = P(dv.getInt32(o + 16, true), dv.getInt32(o + 20, true));
          st.clip = intersect(st.clip, [Math.min(x1, x2), Math.min(y1, y2), Math.max(x1, x2), Math.max(y1, y2)]);
          break;
        }
        case 75: { // EXTSELECTCLIPRGN (unidades de dispositivo)
          const cb = dv.getUint32(o + 8, true), mode = dv.getUint32(o + 12, true);
          if (!cb) {
            if (mode === 5) st.clip = null;
            break;
          }
          const rb = [32, 36, 40, 44].map((k) => dv.getInt32(o + 16 + k, true));
          st.clip = mode === 1 ? intersect(st.clip, rb) : mode === 5 ? rb : st.clip;
          break;
        }
        case 59: path = []; break; // BEGINPATH
        case 60: break; // ENDPATH
        case 61: if (path) path.push('Z'); break; // CLOSEFIGURE
        case 62: case 63: case 64: { // FILLPATH / STROKEANDFILLPATH / STROKEPATH
          if (path && path.length) {
            const attrs = { d: path.join(' '), fill: type === 64 ? 'none' : fill() };
            Object.assign(attrs, type === 62 ? { stroke: 'none' } : stroke());
            add('path', attrs);
          }
          path = null;
          break;
        }
        case 88: case 89: { // POLYBEZIERTO16 / POLYLINETO16
          const count = dv.getUint32(o + 24, true);
          const pts = [];
          for (let i = 0; i < count; i++) pts.push(P(dv.getInt16(o + 28 + i * 4, true), dv.getInt16(o + 30 + i * 4, true)).join(','));
          if (count) st.cur = [dv.getInt16(o + 28 + (count - 1) * 4, true), dv.getInt16(o + 30 + (count - 1) * 4, true)];
          if (path) path.push((type === 88 ? 'C' : 'L') + pts.join(' '));
          else if (!st.pen.none) add('polyline', { points: P(...st.cur).join(',') + ' ' + pts.join(' '), fill: 'none', ...stroke() });
          break;
        }
        case 81: { // STRETCHDIBITS
          const x = dv.getInt32(o + 24, true), y = dv.getInt32(o + 28, true);
          if (dv.getUint32(o + 68, true) === 0x00cc0020) dib(o + dv.getUint32(o + 48, true), dv.getUint32(o + 52, true), o + dv.getUint32(o + 56, true), x, y, dv.getInt32(o + 72, true), dv.getInt32(o + 76, true));
          break;
        }
        case 33: stack.push({ ...st }); break; // SAVEDC
        case 34: { // RESTOREDC
          const n = dv.getInt32(o + 8, true);
          let s = null;
          for (let i = 0; i < Math.abs(n || 1) && stack.length; i++) s = stack.pop();
          if (s) st = s;
          break;
        }
        case 37: { // SELECTOBJECT
          const ih = dv.getUint32(o + 8, true);
          const obj = ih & 0x80000000 ? STOCK[ih & 0x7fffffff] : objects[ih];
          if (obj) st[obj.kind] = obj;
          break;
        }
        case 38: { // CREATEPEN
          const style = dv.getUint32(o + 12, true);
          objects[dv.getUint32(o + 8, true)] = { kind: 'pen', none: (style & 0xf) === 5, width: dv.getInt32(o + 16, true), color: color(dv, o + 24) };
          break;
        }
        case 95: { // EXTCREATEPEN
          const style = dv.getUint32(o + 28, true);
          objects[dv.getUint32(o + 8, true)] = { kind: 'pen', none: (style & 0xf) === 5, geometric: !!(style & 0x10000), width: dv.getUint32(o + 32, true), color: color(dv, o + 40) };
          break;
        }
        case 39: { // CREATEBRUSHINDIRECT
          const style = dv.getUint32(o + 12, true);
          objects[dv.getUint32(o + 8, true)] = { kind: 'brush', none: style === 1, color: color(dv, o + 16) };
          break;
        }
        case 82: { // EXTCREATEFONTINDIRECTW
          const h = dv.getInt32(o + 12, true);
          let face = '';
          for (let i = 0; i < 32; i++) {
            const c = dv.getUint16(o + 40 + i * 2, true);
            if (!c) break;
            face += String.fromCharCode(c);
          }
          objects[dv.getUint32(o + 8, true)] = {
            kind: 'font', size: Math.abs(h) * (h > 0 ? 0.85 : 1) || 12, weight: dv.getInt32(o + 28, true) || 400,
            italic: !!dv.getUint8(o + 32), underline: !!dv.getUint8(o + 33), face: face || 'sans-serif',
          };
          break;
        }
        case 40: delete objects[dv.getUint32(o + 8, true)]; break; // DELETEOBJECT
        case 54: { // LINETO
          const x = dv.getInt32(o + 8, true), y = dv.getInt32(o + 12, true);
          if (path) path.push('L' + P(x, y).join(','));
          else if (!st.pen.none) {
            const [x1, y1] = P(...st.cur), [x2, y2] = P(x, y);
            add('line', { x1, y1, x2, y2, ...stroke() });
          }
          st.cur = [x, y];
          break;
        }
        case 43: { // RECTANGLE
          const l = dv.getInt32(o + 8, true), t = dv.getInt32(o + 12, true), r = dv.getInt32(o + 16, true), bt = dv.getInt32(o + 20, true);
          add('polygon', { points: rectPath(l, t, r, bt), fill: fill(), ...stroke() });
          break;
        }
        case 76: { // BITBLT: sin mapa de bits = relleno con el pincel actual
          const x = dv.getInt32(o + 24, true), y = dv.getInt32(o + 28, true), cx = dv.getInt32(o + 32, true), cy = dv.getInt32(o + 36, true);
          const cbBmi = dv.getUint32(o + 88, true), rop = dv.getUint32(o + 40, true);
          // solo operaciones que pintan: copiar mapa de bits, rellenar con pincel, blanco o negro
          if (cbBmi) {
            if (rop === 0x00cc0020) dib(o + dv.getUint32(o + 84, true), cbBmi, o + dv.getUint32(o + 92, true), x, y, cx, cy);
          } else if (rop === 0x00f00021 && !st.brush.none) add('polygon', { points: rectPath(x, y, x + cx, y + cy), fill: st.brush.color, 'shape-rendering': 'crispEdges' });
          else if (rop === 0x00ff0062 || rop === 0x00000042) add('polygon', { points: rectPath(x, y, x + cx, y + cy), fill: rop === 0x00000042 ? '#000' : '#fff', 'shape-rendering': 'crispEdges' });
          break;
        }
        case 86: case 87: { // POLYGON16 / POLYLINE16
          const count = dv.getUint32(o + 24, true);
          if (path) path.push('M' + points16(o + 28, count).replace(/ /g, ' L') + (type === 86 ? ' Z' : ''));
          else add(type === 86 ? 'polygon' : 'polyline', { points: points16(o + 28, count), fill: type === 86 ? fill() : 'none', ...stroke() });
          break;
        }
        case 84: { // EXTTEXTOUTW
          const rx = dv.getInt32(o + 36, true), ry = dv.getInt32(o + 40, true);
          const n = dv.getUint32(o + 44, true), so = dv.getUint32(o + 48, true);
          let s = '';
          for (let i = 0; i < n && o + so + i * 2 + 1 < dv.byteLength; i++) s += String.fromCharCode(dv.getUint16(o + so + i * 2, true));
          if (!s.trim()) break;
          const a = st.align;
          const [px, py] = a & 1 ? st.cur : [rx, ry];
          const f = st.font;
          const [tx, ty] = P(px, py);
          add('text', {
            x: tx, y: ty,
            'text-anchor': (a & 6) === 6 ? 'middle' : a & 2 ? 'end' : 'start',
            'dominant-baseline': (a & 24) === 24 ? 'alphabetic' : a & 8 ? 'text-after-edge' : 'text-before-edge',
 'font-family': `${f.face}, Arial, sans-serif`, 'font-size': +(f.size * scale()).toFixed(2), 'font-weight': f.weight >= 600 ? 700 : 400,
            'font-style': f.italic ? 'italic' : 'normal', fill: st.text, 'xml:space': 'preserve',
          }, s);
          break;
        }
        case 14: off = dv.byteLength; break; // EOF
        default: break;
      }
      off += size;
    }
    return { svg, width: W, height: H };
  }

  global.EMF = { render };
})(window);

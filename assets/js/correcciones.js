/*
 * correcciones.js — Ortografía de los textos que vienen del Excel.
 *
 * La página no modifica el archivo: corrige lo que muestra. Para agregar más
 * correcciones sin tocar el código, usa "correcciones" en data/config.json:
 *   "correcciones": { "texto como está en Excel": "texto corregido" }
 */
(function (global) {
  'use strict';

  // Frases completas (se comparan sin distinguir mayúsculas)
  const FRASES = {
    'Dsitribución Ocupación': 'Distribución por Ocupación',
    'Retiros Sin justa Causa': 'Retiros Sin Justa Causa',
    'Desempeño 2 trimestre': 'Desempeño 2.º trimestre',
    'Prom': 'Promedio',
    'OKR´s': 'OKR',
    'Las bases de estos indicadores se alimentan total compañía y directamente.':
      'Las bases de estos indicadores se alimentan directamente con la información del total de la compañía.',
  };

  // Palabras (se respeta si estaban en MAYÚSCULAS o con inicial mayúscula)
  const PALABRAS = {
    dsitribución: 'distribución', distribucion: 'distribución', genero: 'género', generos: 'géneros',
    antiguedad: 'antigüedad', regimen: 'régimen', perído: 'período', perídos: 'períodos',
    critica: 'crítica', critico: 'crítico', termino: 'término', basico: 'básico', basica: 'básica',
    terminacion: 'terminación', pasantia: 'pasantía', gestion: 'gestión', vancates: 'vacantes',
    parametros: 'parámetros', parametro: 'parámetro', seleccion: 'selección', rotacion: 'rotación',
    formacion: 'formación', contratacion: 'contratación', ocupacion: 'ocupación', pension: 'pensión',
    generacion: 'generación', induccion: 'inducción', evaluacion: 'evaluación', informacion: 'información',
    poblacion: 'población', direccion: 'dirección', area: 'área', areas: 'áreas', indice: 'índice',
    numero: 'número', dias: 'días', actalización: 'actualización', actualizacion: 'actualización',
    categoria: 'categoría', compañia: 'compañía', compania: 'compañía', medica: 'médica', medico: 'médico',
    domestica: 'doméstica', codigo: 'código', analisis: 'análisis', tecnico: 'técnico', tecnica: 'técnica',
    administracion: 'administración', comunicacion: 'comunicación', participacion: 'participación',
    clasificacion: 'clasificación', capacitacion: 'capacitación', validacion: 'validación',
    articulacion: 'articulación', retroalimentacion: 'retroalimentación', remuneracion: 'remuneración',
    votacion: 'votación',
    ejecucion: 'ejecución', distribuicion: 'distribución', proposito: 'propósito', vacacion: 'vacación',
    '2er': '2.º', '4er': '4.º',
  };

  const MESES = { jan: 'ene', feb: 'feb', mar: 'mar', apr: 'abr', may: 'may', jun: 'jun', jul: 'jul', aug: 'ago', sep: 'sep', oct: 'oct', nov: 'nov', dec: 'dic' };
  const MESES_LARGOS = { january: 'enero', february: 'febrero', march: 'marzo', april: 'abril', may: 'mayo', june: 'junio', july: 'julio', august: 'agosto', september: 'septiembre', october: 'octubre', november: 'noviembre', december: 'diciembre' };

  let frases = new Map();
  let palabrasRe = null;
  let palabras = {};
  const memo = new Map();

  function build(extra) {
    frases = new Map();
    for (const [k, v] of Object.entries({ ...FRASES, ...(extra || {}) })) frases.set(k.trim().toLowerCase(), v);
    palabras = { ...PALABRAS };
    const keys = Object.keys(palabras).sort((a, b) => b.length - a.length).map((k) => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    palabrasRe = new RegExp(`(?<![\\p{L}\\p{N}])(${keys.join('|')})(?![\\p{L}\\p{N}])`, 'giu');
    memo.clear();
  }

  // Mantiene MAYÚSCULAS / Inicial como estaban
  function sameCase(orig, fixed) {
    if (orig === orig.toUpperCase() && orig !== orig.toLowerCase()) return fixed.toUpperCase();
    if (orig[0] === orig[0].toUpperCase()) return fixed[0].toUpperCase() + fixed.slice(1);
    return fixed;
  }

  // Meses en inglés de los formatos de fecha de Excel ("Jan-26" -> "ene-26")
  function spanishMonths(text, anywhere) {
    let s = String(text);
    // "May" es igual en ambas listas: se trata como abreviatura
    s = s.replace(/(?<![\p{L}])(january|february|march|april|june|july|august|september|october|november|december)(?![\p{L}])/giu, (m) => MESES_LARGOS[m.toLowerCase()]);
    const re = anywhere
      ? /(?<![\p{L}])(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(?![\p{L}])/giu
      : /(?<![\p{L}])(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(?=[-/ .]?\d{2,4}(?!\d))/giu;
    return s.replace(re, (m) => MESES[m.toLowerCase()]);
  }

  function fix(text) {
    if (text == null) return text;
    const s = String(text);
    if (!s.trim() || !/[\p{L}/]/u.test(s)) return s;
    if (memo.has(s)) return memo.get(s);
    let out = frases.get(s.trim().toLowerCase());
    if (out == null) {
      out = s
        .replace(palabrasRe, (m) => sameCase(m, palabras[m.toLowerCase()]))
        .replace(/(?<![\d.])([2-9]|\d{2,}) (per[ií]odo)(?![\p{L}])/giu, (m, n, w) => `${n} ${sameCase(w, 'períodos')}`)
        .replace(/\s*\/\/\s*/g, ' / ')
        .replace(/ {2,}/g, ' ');
      out = spanishMonths(out, false);
    }
    if (memo.size > 50000) memo.clear();
    memo.set(s, out);
    return out;
  }

  build();
  global.Correcciones = { fix, spanishMonths, set: build };
})(window);

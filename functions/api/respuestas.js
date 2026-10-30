// Recibe las respuestas de encuesta-semillas.html y las guarda en Cloudflare D1 (binding "DB").
// POST /api/respuestas  →  { ok: true, radicado: "S26-XXXXXXXX" }

const ANIO = 2026;
const MAX_BYTES = 1_000_000;
const MAX_MATERIALES = 300;
const MAX_VENTAS = 1000;

const CATS = {
  prod: { label: `Semilla producida en ${ANIO}`, campos: ['cant', 'dens', 'desc', 'anio'] },
  pre: { label: 'Semilla en previvero para venta o entrega', campos: ['cant', 'dens', 'desc', 'edad'] },
  viv: { label: 'Semilla en vivero para venta o entrega', campos: ['cant', 'dens', 'desc', 'edad'] },
  imp: { label: 'Semilla importada', campos: ['cant', 'dens', 'desc', 'anio'] },
  sie: { label: 'Semilla sembrada por la propia empresa', campos: ['cant', 'mun', 'anio', 'ha'] },
};
const TIPOS = ['Intermediario', 'Cultivador de palma de aceite', 'Intermediario y cultivador'];
const DEPARTAMENTOS = ['Amazonas', 'Antioquia', 'Arauca', 'Atlántico', 'Bogotá D.C.', 'Bolívar', 'Boyacá',
  'Caldas', 'Caquetá', 'Casanare', 'Cauca', 'Cesar', 'Chocó', 'Córdoba', 'Cundinamarca', 'Guainía',
  'Guaviare', 'Huila', 'La Guajira', 'Magdalena', 'Meta', 'Nariño', 'Norte de Santander', 'Putumayo',
  'Quindío', 'Risaralda', 'San Andrés, Providencia y Santa Catalina', 'Santander', 'Sucre', 'Tolima',
  'Valle del Cauca', 'Vaupés', 'Vichada'];

// Same rules as parseNumber() in encuesta-semillas.html: accepts "120.000" and "12,5".
function parseNumber(raw, kind) {
  let s = String(raw ?? '').trim().replace(/\s+/g, '');
  if (!s) return null;
  if (kind === 'int') {
    if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
    return /^\d+$/.test(s) ? Number(s) : NaN;
  }
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '');
  s = s.replace(',', '.');
  return /^\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
}

function checker() {
  const errores = [];
  const txt = (v, donde, { req = true, max = 300 } = {}) => {
    const s = String(v ?? '').trim();
    if (!s) { if (req) errores.push(`${donde}: obligatorio`); return null; }
    if (s.length > max) { errores.push(`${donde}: demasiado largo`); return s.slice(0, max); }
    return s;
  };
  const num = (v, donde, kind, { min, max, gt } = {}) => {
    const n = parseNumber(v, kind);
    if (n == null) { errores.push(`${donde}: obligatorio`); return null; }
    if (Number.isNaN(n)) { errores.push(`${donde}: número no válido`); return null; }
    if ((gt != null && !(n > gt)) || (min != null && n < min) || (max != null && n > max)) {
      errores.push(`${donde}: fuera de rango`);
      return null;
    }
    return n;
  };
  const oneOf = (v, lista, donde) => {
    if (!lista.includes(v)) { errores.push(`${donde}: opción no válida`); return null; }
    return v;
  };
  return { errores, txt, num, oneOf };
}

function validar(d) {
  const { errores, txt, num, oneOf } = checker();
  const e = d.empresa || {};
  const nit = txt(e.emp_nit, 'NIT', { max: 20 });
  if (nit && !/^\d{5,12}-?\d$/.test(nit.replace(/[.\s]/g, ''))) errores.push('NIT: formato no válido');
  const correo = txt(e.emp_correo, 'Correo', { max: 200 });
  if (correo && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) errores.push('Correo: formato no válido');
  const actividad = oneOf(d.actividad, ['Sí', 'No'], 'Actividad');

  const envio = {
    nit: nit && nit.replace(/[.\s]/g, ''),
    empresa: txt(e.emp_nombre, 'Empresa'),
    rep: txt(e.emp_rep, 'Representante legal'),
    rep_id: txt(e.emp_rep_id, 'Cédula del representante', { max: 30 }),
    tel: txt(e.emp_tel, 'Teléfono', { max: 30 }),
    cel: txt(e.emp_cel, 'Celular', { max: 30 }),
    fax: txt(e.emp_fax, 'Fax', { req: false, max: 30 }),
    direccion: txt(e.emp_dir, 'Dirección'),
    ciudad: txt(e.emp_ciudad, 'Ciudad'),
    dil: txt(e.emp_dil, 'Diligenciado por'),
    cargo: txt(e.emp_cargo, 'Cargo'),
    correo,
    fecha: txt(e.emp_fecha, 'Fecha', { max: 10 }),
    actividad,
    obs: txt(d.obs, 'Observaciones', { req: false, max: 2000 }),
    declaracion: d.declaracion ? 1 : 0,
  };
  if (envio.fecha && !/^\d{4}-\d{2}-\d{2}$/.test(envio.fecha)) errores.push('Fecha: formato no válido');

  const activo = actividad === 'Sí';
  if (activo && !envio.declaracion) errores.push('Declaración: debe aceptarse');
  const mats = activo && Array.isArray(d.materiales) ? d.materiales : [];
  const vens = activo && Array.isArray(d.ventas) ? d.ventas : [];
  if (mats.length > MAX_MATERIALES) errores.push(`Materiales: máximo ${MAX_MATERIALES}`);
  if (vens.length > MAX_VENTAS) errores.push(`Ventas: máximo ${MAX_VENTAS}`);

  const materiales = mats.slice(0, MAX_MATERIALES).map((m, i) => {
    const donde = `Material ${i + 1}`;
    const cat = CATS[m.cat];
    if (!cat) { errores.push(`${donde}: tipo de registro no válido`); return null; }
    const usa = (c) => cat.campos.includes(c);
    return {
      n: i + 1,
      categoria: cat.label,
      material: txt(m.mat, `${donde}, material`),
      cantidad: num(m.cant, `${donde}, cantidad`, 'int', { min: 0 }),
      densidad: usa('dens') ? num(m.dens, `${donde}, densidad`, 'dec', { gt: 0 }) : null,
      descarte: usa('desc') ? num(m.desc, `${donde}, descarte`, 'dec', { min: 0, max: 100 }) : null,
      anio: usa('anio') ? num(m.anio, `${donde}, año`, 'int', { min: 2000, max: ANIO }) : null,
      edad: usa('edad') ? num(m.edad, `${donde}, edad`, 'dec', { min: 0 }) : null,
      municipio: usa('mun') ? txt(m.mun, `${donde}, municipio`) : null,
      hectareas: usa('ha') ? num(m.ha, `${donde}, hectáreas`, 'dec', { min: 0 }) : null,
      obs: txt(m.obs, `${donde}, observaciones`, { req: false, max: 1000 }),
    };
  });

  const ventas = vens.slice(0, MAX_VENTAS).map((v, i) => {
    const donde = `Venta ${i + 1}`;
    return {
      n: i + 1,
      anio: num(v.anio, `${donde}, año`, 'int', { min: 2000, max: ANIO }),
      comprador: txt(v.razon, `${donde}, comprador`),
      comprador_id: txt(v.id, `${donde}, identificación`, { max: 30 }),
      telefono: txt(v.tel, `${donde}, teléfono`, { max: 30 }),
      material: txt(v.mat, `${donde}, material`),
      cantidad: num(v.cant, `${donde}, cantidad`, 'int', { gt: 0 }),
      departamento: oneOf(v.dep, DEPARTAMENTOS, `${donde}, departamento`),
      municipio: txt(v.mun, `${donde}, municipio`),
      tipo: oneOf(v.tipo, TIPOS, `${donde}, tipo de comprador`),
      obs: txt(v.obs, `${donde}, observaciones`, { req: false, max: 1000 }),
    };
  });

  return { errores, envio, materiales, ventas };
}

function radicado() {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const bytes = crypto.getRandomValues(new Uint8Array(10));
  return `S${String(ANIO).slice(2)}-` + Array.from(bytes, (b) => abc[b % abc.length]).join('');
}

const json = (body, status = 200) => new Response(JSON.stringify(body), {
  status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' },
});

export async function onRequestPost({ request, env }) {
  if (!env.DB) return json({ ok: false, error: 'La base de datos no está configurada (binding DB).' }, 500);
  const raw = await request.text();
  if (raw.length > MAX_BYTES) return json({ ok: false, error: 'El envío es demasiado grande.' }, 413);
  let d;
  try { d = JSON.parse(raw); } catch { return json({ ok: false, error: 'Datos no válidos.' }, 400); }
  // Campo trampa: las personas no lo ven; los robots que llenan todo, sí.
  if (d.sitio_web) return json({ ok: true, radicado: radicado() });

  const { errores, envio, materiales, ventas } = validar(d);
  if (errores.length) return json({ ok: false, error: 'Hay datos por corregir.', detalles: errores.slice(0, 20) }, 400);

  const id = radicado();
  const cols = Object.keys(envio);
  const stmts = [
    env.DB.prepare(`INSERT INTO envios (id, anio, ${cols.join(', ')}) VALUES (?, ?, ${cols.map(() => '?').join(', ')})`)
      .bind(id, ANIO, ...cols.map((c) => envio[c])),
  ];
  // Una sola sentencia por tabla (json_each), para no superar el límite de consultas por invocación.
  const insertJson = (tabla, rows) => {
    const c = Object.keys(rows[0]);
    stmts.push(env.DB.prepare(
      `INSERT INTO ${tabla} (envio_id, ${c.join(', ')}) SELECT ?1, ${c.map((k) => `json_extract(value, '$.${k}')`).join(', ')} FROM json_each(?2)`,
    ).bind(id, JSON.stringify(rows)));
  };
  if (materiales.length) insertJson('materiales', materiales);
  if (ventas.length) insertJson('ventas', ventas);

  try {
    await env.DB.batch(stmts);
  } catch (err) {
    console.error('D1', err);
    return json({ ok: false, error: 'No se pudo guardar. Intente de nuevo en unos minutos.' }, 500);
  }
  return json({ ok: true, radicado: id });
}

export function onRequest() {
  return json({ ok: false, error: 'Método no permitido.' }, 405);
}

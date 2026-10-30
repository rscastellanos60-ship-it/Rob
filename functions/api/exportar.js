// Entrega todas las respuestas guardadas en D1 para admin-semillas.html.
// GET /api/exportar con el encabezado "Authorization: Bearer <ADMIN_KEY>".
// ADMIN_KEY se define como variable secreta del proyecto en Cloudflare Pages.

const json = (body, status = 200) => new Response(JSON.stringify(body), {
  status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' },
});

async function claveValida(dada, esperada) {
  const enc = new TextEncoder();
  // Compare digests so both buffers have the same length for timingSafeEqual.
  const [a, b] = await Promise.all([dada, esperada].map((s) => crypto.subtle.digest('SHA-256', enc.encode(s))));
  return crypto.subtle.timingSafeEqual(a, b);
}

export async function onRequestGet({ request, env }) {
  if (!env.DB) return json({ ok: false, error: 'La base de datos no está configurada (binding DB).' }, 500);
  if (!env.ADMIN_KEY || env.ADMIN_KEY.length < 12) {
    return json({ ok: false, error: 'Falta definir ADMIN_KEY (mínimo 12 caracteres) en Cloudflare.' }, 500);
  }
  const auth = request.headers.get('authorization') || '';
  const dada = auth.startsWith('Bearer ') ? auth.slice(7) : '';
  if (!dada || !(await claveValida(dada, env.ADMIN_KEY))) return json({ ok: false, error: 'Clave incorrecta.' }, 401);

  const [envios, materiales, ventas] = await env.DB.batch([
    env.DB.prepare('SELECT * FROM envios ORDER BY recibido, id'),
    env.DB.prepare(`SELECT m.*, e.nit, e.empresa, e.recibido FROM materiales m JOIN envios e ON e.id = m.envio_id
                    ORDER BY e.recibido, m.envio_id, m.n`),
    env.DB.prepare(`SELECT v.*, e.nit, e.empresa, e.recibido FROM ventas v JOIN envios e ON e.id = v.envio_id
                    ORDER BY e.recibido, v.envio_id, v.n`),
  ]);
  return json({ ok: true, envios: envios.results, materiales: materiales.results, ventas: ventas.results });
}

export function onRequest() {
  return json({ ok: false, error: 'Método no permitido.' }, 405);
}

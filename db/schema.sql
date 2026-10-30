-- Base de datos de la Encuesta de Semillas (Cloudflare D1).
-- Pegue este archivo en la consola de D1 una sola vez, al crear la base de datos.

CREATE TABLE IF NOT EXISTS envios (
  id          TEXT PRIMARY KEY,               -- radicado que ve el encuestado
  recibido    TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now')), -- UTC
  anio        INTEGER NOT NULL,
  nit         TEXT NOT NULL,
  empresa     TEXT NOT NULL,
  rep         TEXT,
  rep_id      TEXT,
  tel         TEXT,
  cel         TEXT,
  fax         TEXT,
  direccion   TEXT,
  ciudad      TEXT,
  dil         TEXT,
  cargo       TEXT,
  correo      TEXT,
  fecha       TEXT,
  actividad   TEXT NOT NULL,
  obs         TEXT,
  declaracion INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS envios_nit ON envios (nit, recibido);

CREATE TABLE IF NOT EXISTS materiales (
  envio_id  TEXT NOT NULL REFERENCES envios (id),
  n         INTEGER NOT NULL,
  categoria TEXT NOT NULL,
  material  TEXT NOT NULL,
  cantidad  INTEGER,
  densidad  REAL,
  descarte  REAL,
  anio      INTEGER,
  edad      REAL,
  municipio TEXT,
  hectareas REAL,
  obs       TEXT,
  PRIMARY KEY (envio_id, n)
);

CREATE TABLE IF NOT EXISTS ventas (
  envio_id     TEXT NOT NULL REFERENCES envios (id),
  n            INTEGER NOT NULL,
  anio         INTEGER,
  comprador    TEXT NOT NULL,
  comprador_id TEXT,
  telefono     TEXT,
  material     TEXT,
  cantidad     INTEGER,
  departamento TEXT,
  municipio    TEXT,
  tipo         TEXT,
  obs          TEXT,
  PRIMARY KEY (envio_id, n)
);

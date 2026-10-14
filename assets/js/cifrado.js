/*
 * cifrado.js — Protege el Excel publicado con una clave.
 *
 * El archivo se guarda cifrado con AES-GCM (256 bits); la llave se deriva de
 * la clave con PBKDF2-SHA256. Sin la clave el archivo publicado no se puede
 * leer. Todo ocurre en el navegador: la clave nunca se envía a ningún lado.
 *
 * Formato: "TBX1" | iteraciones (uint32) | sal (16 bytes) | iv (12 bytes) | datos cifrados
 * Datos descifrados: largo del nombre (uint32) | nombre del archivo (UTF-8) | contenido
 */
(function (global) {
  'use strict';

  const MAGIC = [0x54, 0x42, 0x58, 0x31]; // "TBX1"
  const ITERATIONS = 600000;
  const subtle = () => {
    const c = global.crypto && global.crypto.subtle;
    if (!c) throw new Error('Este navegador no permite descifrar el archivo (se necesita HTTPS).');
    return c;
  };

  async function deriveKey(password, salt, iterations) {
    const base = await subtle().importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
    return subtle().deriveKey({ name: 'PBKDF2', salt, iterations, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  }

  function isEncrypted(bytes) {
    const b = new Uint8Array(bytes);
    return b.length > 36 && MAGIC.every((m, i) => b[i] === m);
  }

  async function encrypt(fileBytes, fileName, password) {
    const salt = global.crypto.getRandomValues(new Uint8Array(16));
    const iv = global.crypto.getRandomValues(new Uint8Array(12));
    const key = await deriveKey(password, salt, ITERATIONS);
    const name = new TextEncoder().encode(fileName || 'tablero.xlsx');
    const plain = new Uint8Array(4 + name.length + fileBytes.byteLength);
    new DataView(plain.buffer).setUint32(0, name.length);
    plain.set(name, 4);
    plain.set(new Uint8Array(fileBytes), 4 + name.length);
    const ct = new Uint8Array(await subtle().encrypt({ name: 'AES-GCM', iv }, key, plain));
    const out = new Uint8Array(4 + 4 + 16 + 12 + ct.length);
    out.set(MAGIC, 0);
    new DataView(out.buffer).setUint32(4, ITERATIONS);
    out.set(salt, 8);
    out.set(iv, 24);
    out.set(ct, 36);
    return out;
  }

  // Devuelve { name, buffer } o lanza un error "Clave incorrecta"
  async function decrypt(encBytes, password) {
    const b = new Uint8Array(encBytes);
    if (!isEncrypted(b)) throw new Error('El archivo publicado no tiene el formato esperado.');
    const iterations = new DataView(b.buffer, b.byteOffset, b.byteLength).getUint32(4);
    const key = await deriveKey(password, b.slice(8, 24), iterations);
    let plain;
    try {
      plain = new Uint8Array(await subtle().decrypt({ name: 'AES-GCM', iv: b.slice(24, 36) }, key, b.slice(36)));
    } catch (e) {
      const err = new Error('Clave incorrecta');
      err.wrongPassword = true;
      throw err;
    }
    const nameLen = new DataView(plain.buffer).getUint32(0);
    const name = new TextDecoder().decode(plain.slice(4, 4 + nameLen));
    return { name, buffer: plain.slice(4 + nameLen).buffer };
  }

  global.Cifrado = { encrypt, decrypt, isEncrypted };
})(typeof window !== 'undefined' ? window : globalThis);

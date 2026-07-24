/**
 * Persist the scored coaching bank.
 * Prefers Netlify Blobs when available; falls back to data/results.json.
 */

import { promises as fs } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const LOCAL_PATH = path.join(__dirname, '..', 'data', 'results.json');

async function getBlobStore(storeName) {
  try {
    const { getStore } = await import('@netlify/blobs');
    return getStore(storeName);
  } catch {
    return null;
  }
}

export async function loadBank(storeName) {
  const store = await getBlobStore(storeName);
  if (store) {
    try {
      const raw = await store.get('bank.json', { type: 'json' });
      if (raw) return raw;
    } catch {
      // empty store
    }
  }

  try {
    const text = await fs.readFile(LOCAL_PATH, 'utf8');
    return JSON.parse(text);
  } catch {
    return emptyBank();
  }
}

export async function saveBank(storeName, bank) {
  const payload = {
    ...bank,
    updatedAt: new Date().toISOString(),
  };

  const store = await getBlobStore(storeName);
  if (store) {
    try {
      await store.setJSON('bank.json', payload);
      return { backend: 'blobs', payload };
    } catch (err) {
      // fall through to local file for netlify dev / unit contexts
      console.warn('Blob save failed, writing local file:', err.message);
    }
  }

  await fs.mkdir(path.dirname(LOCAL_PATH), { recursive: true });
  await fs.writeFile(LOCAL_PATH, JSON.stringify(payload, null, 2));
  return { backend: 'file', payload };
}

export function emptyBank() {
  return {
    updatedAt: null,
    lastSync: null,
    sessions: [],
    coaches: [],
    meta: {},
  };
}

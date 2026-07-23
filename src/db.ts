import { openDB, type DBSchema, type IDBPDatabase } from 'idb';
import type { WorkbookRef } from './graph';
import type { PendingClose, RegisterAsset, ScanRecord, Session } from './types';

interface StocktakeDB extends DBSchema {
  register: {
    key: string; // assetTag
    value: RegisterAsset;
  };
  sessions: {
    key: string; // sessionId
    value: Session;
    indexes: { byStatus: string };
  };
  scans: {
    key: string; // scanId
    value: ScanRecord;
    indexes: { bySession: string; bySynced: number; bySyncedTime: [number, string] };
  };
  meta: {
    key: string;
    value: { key: string; value: unknown };
  };
}

let dbPromise: Promise<IDBPDatabase<StocktakeDB>> | null = null;

export function getDB(): Promise<IDBPDatabase<StocktakeDB>> {
  if (!dbPromise) {
    dbPromise = openDB<StocktakeDB>('asset-stocktake', 2, {
      upgrade(db, oldVersion) {
        // v1 keyed the register on `assetTag` and stored a ScanRecord shape that
        // no longer exists. Nothing readable survives the change, and v1 could
        // never load a register in the first place, so rebuild from empty.
        if (oldVersion > 0) {
          for (const name of ['register', 'sessions', 'scans', 'meta'] as const) {
            if (db.objectStoreNames.contains(name)) db.deleteObjectStore(name);
          }
        }
        db.createObjectStore('register', { keyPath: 'assetId' });
        const sessions = db.createObjectStore('sessions', { keyPath: 'sessionId' });
        sessions.createIndex('byStatus', 'status');
        const scans = db.createObjectStore('scans', { keyPath: 'scanId' });
        scans.createIndex('bySession', 'sessionId');
        scans.createIndex('bySynced', 'synced');
        // Compound so a batch drains oldest-first; the plain bySynced index is
        // ordered by random UUID, which would scramble the audit log.
        scans.createIndex('bySyncedTime', ['synced', 'timestamp']);
        db.createObjectStore('meta', { keyPath: 'key' });
      },
    });
  }
  return dbPromise;
}

export class StorageFullError extends Error {
  constructor() {
    super('Device storage is full. Scans can no longer be saved.');
    this.name = 'StorageFullError';
  }
}

/**
 * Every write funnels through here so quota exhaustion is always converted into
 * a typed, blocking error — persistence must never fail silently.
 */
async function guarded<T>(op: () => Promise<T>): Promise<T> {
  try {
    return await op();
  } catch (e) {
    if (e instanceof DOMException && (e.name === 'QuotaExceededError' || e.name === 'AbortError')) {
      throw new StorageFullError();
    }
    throw e;
  }
}

// ---- register cache ----

export async function replaceRegister(assets: RegisterAsset[]): Promise<void> {
  const db = await getDB();
  await guarded(async () => {
    const tx = db.transaction(['register', 'meta'], 'readwrite');
    await tx.objectStore('register').clear();
    for (const a of assets) tx.objectStore('register').put(a);
    tx.objectStore('meta').put({ key: 'registerFetchedAt', value: new Date().toISOString() });
    await tx.done;
  });
}

export async function loadRegister(): Promise<RegisterAsset[]> {
  const db = await getDB();
  return db.getAll('register');
}

export async function registerFetchedAt(): Promise<string | null> {
  const db = await getDB();
  const row = await db.get('meta', 'registerFetchedAt');
  return (row?.value as string) ?? null;
}

// ---- chosen workbook ----

/**
 * Picked once per device and remembered, so a scanner isn't asked to find the
 * file again every shift. Stored locally rather than derived, because the app
 * must know its target without a network round trip.
 */
export async function getWorkbook(): Promise<WorkbookRef | null> {
  const db = await getDB();
  const row = await db.get('meta', 'workbook');
  return (row?.value as WorkbookRef) ?? null;
}

export async function setWorkbook(wb: WorkbookRef): Promise<void> {
  const db = await getDB();
  await guarded(() => db.put('meta', { key: 'workbook', value: wb }).then(() => undefined));
}

// ---- sessions ----

export async function putSession(s: Session): Promise<void> {
  const db = await getDB();
  await guarded(() => db.put('sessions', s).then(() => undefined));
}

export async function findResumableSession(): Promise<Session | null> {
  const db = await getDB();
  const open = await db.getAllFromIndex('sessions', 'byStatus', 'open');
  const closing = await db.getAllFromIndex('sessions', 'byStatus', 'closing');
  return open[0] ?? closing[0] ?? null;
}

// ---- scans ----

export async function persistScan(scan: ScanRecord): Promise<void> {
  const db = await getDB();
  await guarded(() => db.put('scans', scan).then(() => undefined));
}

export async function scansForSession(sessionId: string): Promise<ScanRecord[]> {
  const db = await getDB();
  return db.getAllFromIndex('scans', 'bySession', sessionId);
}

export async function unsyncedScans(limit: number): Promise<ScanRecord[]> {
  const db = await getDB();
  // Oldest first, so the ScanLog reads in capture order rather than UUID order.
  const range = IDBKeyRange.bound([0, ''], [0, '￿']);
  return db.getAllFromIndex('scans', 'bySyncedTime', range, limit);
}

export async function markSynced(scanIds: string[]): Promise<void> {
  const db = await getDB();
  await guarded(async () => {
    const tx = db.transaction('scans', 'readwrite');
    for (const id of scanIds) {
      const row = await tx.store.get(id);
      if (row) tx.store.put({ ...row, synced: 1 });
    }
    await tx.done;
  });
}

export async function unsyncedCount(): Promise<number> {
  const db = await getDB();
  return db.countFromIndex('scans', 'bySynced', 0);
}

// ---- pending close (re-runnable reconciliation) ----

export async function savePendingClose(pc: PendingClose): Promise<void> {
  const db = await getDB();
  await guarded(() => db.put('meta', { key: `pendingClose:${pc.sessionId}`, value: pc }).then(() => undefined));
}

export async function getPendingClose(sessionId: string): Promise<PendingClose | null> {
  const db = await getDB();
  const row = await db.get('meta', `pendingClose:${sessionId}`);
  return (row?.value as PendingClose) ?? null;
}

export async function clearPendingClose(sessionId: string): Promise<void> {
  const db = await getDB();
  await db.delete('meta', `pendingClose:${sessionId}`);
}

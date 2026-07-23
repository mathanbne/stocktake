import { submitClose, submitScans } from './api';
import {
  clearPendingClose,
  findResumableSession,
  getPendingClose,
  markSynced,
  putSession,
  unsyncedCount,
  unsyncedScans,
} from './db';
import { config } from './config';
import type { Session } from './types';

export interface SyncStatus {
  pending: number;
  lastError: string | null;
  flushing: boolean;
}

type Listener = (s: SyncStatus) => void;

const status: SyncStatus = { pending: 0, lastError: null, flushing: false };
const listeners = new Set<Listener>();
let timer: number | null = null;
let closingSession: Session | null = null;
let onClosed: ((s: Session) => void) | null = null;

function notify(): void {
  for (const l of listeners) l({ ...status });
}

export function subscribeSync(l: Listener): () => void {
  listeners.add(l);
  l({ ...status });
  return () => listeners.delete(l);
}

export async function refreshPending(): Promise<void> {
  status.pending = await unsyncedCount();
  notify();
}

/**
 * Flush is safe to call at any time, from anywhere. Server-side dedupe on
 * scanId means an ambiguous failure (request sent, response lost) is resolved
 * by simply sending the same batch again.
 */
export async function flush(): Promise<void> {
  if (status.flushing || !navigator.onLine) return;
  status.flushing = true;
  notify();
  try {
    // Loop until drained so a long offline backlog clears in one online window.
    for (;;) {
      const batch = await unsyncedScans(config.syncBatchSize);
      if (batch.length === 0) break;
      await submitScans(batch);
      await markSynced(batch.map((s) => s.scanId));
      await refreshPending();
    }
    // Close only after every scan row is safely in the log — the flow derives
    // nothing, but the audit trail must be complete before statuses change.
    if (closingSession) {
      const pc = await getPendingClose(closingSession.sessionId);
      if (pc) {
        await submitClose(pc);
        await clearPendingClose(closingSession.sessionId);
      }
      const closed: Session = { ...closingSession, status: 'closed', closedAt: new Date().toISOString() };
      await putSession(closed);
      const cb = onClosed;
      closingSession = null;
      cb?.(closed);
    }
    status.lastError = null;
  } catch (e) {
    status.lastError = e instanceof Error ? e.message : 'Sync failed';
  } finally {
    status.flushing = false;
    notify();
  }
}

/**
 * A session left in 'closing' by a crash or a force-quit must finish on its own.
 * `closingSession` is module state, so it does not survive a reload — this
 * restores it at boot, making the close genuinely "replays on the next tick"
 * rather than "replays once the user happens to tap Resume".
 */
async function rearmPendingClose(): Promise<void> {
  if (closingSession) return;
  const session = await findResumableSession();
  if (session?.status === 'closing') closingSession = session;
}

export function startSyncLoop(): void {
  if (timer !== null) return;
  void refreshPending();
  void rearmPendingClose().then(() => flush());
  timer = window.setInterval(() => void flush(), config.syncIntervalMs);
  window.addEventListener('online', () => void flush());
}

/** Registers a session whose close should be pushed once the queue drains. */
export function requestClose(session: Session, cb: (s: Session) => void): void {
  closingSession = session;
  onClosed = cb;
  void flush();
}

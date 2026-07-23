import { RegisterIndex, normalizeTag } from './capture';
import { persistScan, putSession, savePendingClose, scansForSession } from './db';
import type { PendingClose, RegisterAsset, ScanRecord, Session } from './types';

/**
 * The set of Asset Ids that count as "seen". Any register match — sighted,
 * wrong location, duplicate, manual — clears an asset from the missing list.
 * unknown_tag never does: it resolved to no asset at all.
 */
export function seenAssetIds(scans: ScanRecord[]): Set<string> {
  const seen = new Set<string>();
  for (const s of scans) {
    if (s.result !== 'unknown_tag' && s.assetId !== '') seen.add(normalizeTag(s.assetId));
  }
  return seen;
}

export async function previewMissing(session: Session, index: RegisterIndex): Promise<RegisterAsset[]> {
  const scans = await scansForSession(session.sessionId);
  const seen = seenAssetIds(scans);
  return index
    .inScope(session)
    .filter((a) => !seen.has(normalizeTag(a.assetId)))
    .sort((a, b) => a.expectedLocation.localeCompare(b.expectedLocation));
}

/**
 * Confirming a close is a local, durable act:
 * 1. append a "missing" audit row per unseen asset (idempotent via scanId),
 * 2. persist the exact close payload,
 * 3. flip the session to 'closing'.
 * The sync engine then delivers it — today, or after the next reboot. Re-running
 * after a partial failure replays the identical payload; the server's dedupe and
 * key-based updates make the replay harmless.
 */
export async function confirmClose(
  session: Session,
  missing: RegisterAsset[],
): Promise<{ session: Session; pendingClose: PendingClose }> {
  const now = new Date().toISOString();
  const scans = await scansForSession(session.sessionId);
  const existingMissing = new Set(
    scans.filter((s) => s.result === 'missing').map((s) => normalizeTag(s.assetId)),
  );

  for (const asset of missing) {
    if (existingMissing.has(normalizeTag(asset.assetId))) continue;
    await persistScan({
      scanId: crypto.randomUUID(),
      sessionId: session.sessionId,
      scannedValue: '',
      assetId: asset.assetId,
      result: 'missing',
      scannedBy: session.scannedBy,
      timestamp: now,
      actualLocation: '',
      notes: `Not sighted; expected: ${asset.expectedLocation}`,
      synced: 0,
    });
  }

  const sighted = scans.filter(
    (s) =>
      s.assetId !== '' &&
      (s.result === 'sighted' || s.result === 'sighted_wrong_location' || s.result === 'manual_entry'),
  );
  // Last scan wins if a tag was sighted then re-entered manually with notes.
  const latestById = new Map<string, ScanRecord>();
  for (const s of sighted) latestById.set(normalizeTag(s.assetId), s);

  const pendingClose: PendingClose = {
    sessionId: session.sessionId,
    missingAssetIds: missing.map((a) => a.assetId),
    sightedUpdates: [...latestById.values()].map((s) => ({
      assetId: s.assetId,
      verifiedDate: s.timestamp,
      verifiedBy: s.scannedBy,
    })),
  };
  await savePendingClose(pendingClose);

  const closing: Session = { ...session, status: 'closing' };
  await putSession(closing);
  return { session: closing, pendingClose };
}

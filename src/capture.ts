import { persistScan } from './db';
import type { RegisterAsset, ScanRecord, ScanResult, Session } from './types';

/**
 * Barcodes and spreadsheet cells disagree about punctuation: the register holds
 * `#9339100142` (and Excel may prefix a text cell with an apostrophe) while the
 * printed label almost certainly encodes `9339100142`. Normalising *both* sides
 * the same way means the app matches whichever form it meets.
 */
export function normalizeTag(raw: string): string {
  return raw
    .replace(/\s+/g, '')
    .replace(/^['#]+/, '')
    .toUpperCase();
}

/**
 * Which register columns a scan may legitimately match, most authoritative
 * first. Barcode wins where it is populated; Asset Id is the practical fallback
 * because it is the only column guaranteed non-empty; Old Asset Id catches
 * assets still wearing a pre-relabel sticker.
 */
const ALIAS_FIELDS: (keyof RegisterAsset)[] = ['barcode', 'assetId', 'serial', 'oldAssetId'];

/**
 * The register lives in an in-memory Map during a session so classification is
 * synchronous — the on-screen result never waits on IndexedDB, let alone the
 * network.
 */
export class RegisterIndex {
  private byAlias = new Map<string, RegisterAsset>();
  private byId = new Map<string, RegisterAsset>();

  constructor(assets: RegisterAsset[]) {
    for (const a of assets) this.byId.set(normalizeTag(a.assetId), a);
    // Field-by-field rather than asset-by-asset, so a lower-priority alias on
    // one asset can never shadow a higher-priority alias on another.
    for (const field of ALIAS_FIELDS) {
      for (const a of assets) {
        const key = normalizeTag(String(a[field] ?? ''));
        if (key === '' || this.byAlias.has(key)) continue;
        this.byAlias.set(key, a);
      }
    }
  }

  lookup(scannedValue: string): RegisterAsset | undefined {
    return this.byAlias.get(normalizeTag(scannedValue));
  }

  byAssetId(assetId: string): RegisterAsset | undefined {
    return this.byId.get(normalizeTag(assetId));
  }

  inScope(session: Session): RegisterAsset[] {
    const needle = session.scopeValue.trim().toLowerCase();
    if (needle === '') return [];
    return [...this.byId.values()].filter((a) =>
      a.expectedLocation.toLowerCase().includes(needle),
    );
  }

  get size(): number {
    return this.byId.size;
  }
}

export interface CaptureOutcome {
  scan: ScanRecord;
  asset?: RegisterAsset;
}

export function classify(
  scannedValue: string,
  index: RegisterIndex,
  alreadyScanned: Set<string>,
  manual: boolean,
): { result: ScanResult; asset?: RegisterAsset; notes: string } {
  const asset = index.lookup(scannedValue);
  if (!asset) return { result: 'unknown_tag', notes: 'No matching register entry' };
  // Dedupe on the resolved Asset Id, not the scanned string — the same asset
  // read once by barcode and once by serial is still one sighting.
  if (alreadyScanned.has(normalizeTag(asset.assetId))) {
    return { result: 'duplicate', asset, notes: 'Already scanned this session' };
  }
  if (manual) {
    return { result: 'manual_entry', asset, notes: 'Typed by hand, not scanned' };
  }
  return { result: 'sighted', asset, notes: '' };
}

/**
 * Order of operations is deliberate: classify (microseconds, in-memory), then
 * persist to IndexedDB, and only report success to the UI after the write
 * resolves. A crash mid-write loses at most this one scan.
 */
export async function captureScan(opts: {
  scannedValue: string;
  session: Session;
  index: RegisterIndex;
  alreadyScanned: Set<string>;
  actualLocation: string;
  manual: boolean;
}): Promise<CaptureOutcome> {
  const { result, asset, notes } = classify(
    opts.scannedValue,
    opts.index,
    opts.alreadyScanned,
    opts.manual,
  );

  // Wrong-location check applies to scanned *and* manual entries; a manual
  // entry in the wrong spot keeps result=manual_entry but records the mismatch.
  let finalResult = result;
  let finalNotes = notes;
  if (asset && (result === 'sighted' || result === 'manual_entry')) {
    const here = opts.actualLocation.trim().toLowerCase();
    const expected = asset.expectedLocation.toLowerCase();
    const mismatch = here !== '' && !expected.includes(here) && !here.includes(expected);
    if (mismatch) {
      if (result === 'sighted') finalResult = 'sighted_wrong_location';
      finalNotes = `Expected: ${asset.expectedLocation}${notes ? ' | ' + notes : ''}`;
    }
  }

  const scan: ScanRecord = {
    scanId: crypto.randomUUID(),
    sessionId: opts.session.sessionId,
    scannedValue: opts.scannedValue.trim(),
    assetId: asset?.assetId ?? '',
    result: finalResult,
    scannedBy: opts.session.scannedBy,
    timestamp: new Date().toISOString(),
    actualLocation: opts.actualLocation,
    notes: finalNotes,
    synced: 0,
  };

  await persistScan(scan);
  return { scan, asset };
}

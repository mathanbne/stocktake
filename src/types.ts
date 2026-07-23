/**
 * Mirrors the real `AssetRegister` table in stock.xlsx. `expectedLocation` is the
 * one derived field — Excel has no such column; it is composed from Facility
 * Group + Facility so a single substring match covers both scope granularities.
 */
export interface RegisterAsset {
  assetId: string; // Asset Id — stable key for every Excel write
  oldAssetId: string;
  barcode: string;
  serial: string;
  assetName: string;
  department: string;
  make: string;
  model: string;
  facility: string; // e.g. 01G04
  facilityGroup: string; // e.g. A BLOCK
  expectedLocation: string; // derived: "A BLOCK / 01G04"
  condition: string;
  status: string; // the workbook's own Status column — read-only, never written
  stocktakeStatus: string; // Cited | Missing — written by this app
  lastVerifiedDate: string;
  lastVerifiedBy: string;
}

/** Human-readable label for a register row, used in flashes and previews. */
export function assetLabel(a: RegisterAsset): string {
  return [a.assetName, [a.make, a.model].filter(Boolean).join(' ')]
    .filter((s) => s.trim() !== '')
    .join(' · ');
}

/**
 * "missing" exists in this union only because reconciliation writes it to the
 * audit trail. Nothing in the scanning path is allowed to produce it — it is
 * derived exclusively when a session closes.
 */
export type ScanResult =
  | 'sighted'
  | 'sighted_wrong_location'
  | 'unknown_tag'
  | 'duplicate'
  | 'manual_entry'
  | 'missing';

export interface ScanRecord {
  scanId: string; // client-generated UUID — the idempotency key for sync
  sessionId: string;
  /** Exactly what came off the scanner or keyboard, preserved for the audit trail. */
  scannedValue: string;
  /** Resolved Asset Id, or '' when nothing in the register matched. */
  assetId: string;
  result: ScanResult;
  scannedBy: string;
  timestamp: string; // ISO 8601
  actualLocation: string;
  notes: string;
  synced: 0 | 1; // number, not boolean: IndexedDB indexes can't key booleans
}

export type ScopeType = 'site' | 'floor' | 'room';

export type SessionStatus = 'open' | 'closing' | 'closed';

export interface Session {
  sessionId: string;
  scopeType: ScopeType;
  scopeValue: string; // matched against expectedLocation (case-insensitive substring)
  scannedBy: string;
  startedAt: string;
  status: SessionStatus;
  closedAt?: string;
}

/** Payload persisted at confirm-time so a failed close can be re-run verbatim. */
export interface PendingClose {
  sessionId: string;
  missingAssetIds: string[];
  sightedUpdates: { assetId: string; verifiedDate: string; verifiedBy: string }[];
}

/** A row of the `AssetRegister` table exactly as Graph returns it. */
export interface RegisterRowWire {
  'Asset Id': string;
  'Old Asset Id': string;
  'Asset Name': string;
  'Department': string;
  'Barcode': string;
  'Facility Group': string;
  'Facility': string;
  'Make': string;
  'Model': string;
  'Serial Number': string;
  'Condition': string;
  'Status': string;
  'Stocktake Status': string;
  'Last Verified Date': string;
  'Last Verified By': string;
}

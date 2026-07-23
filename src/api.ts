import { config } from './config';
import { getAccessCode } from './db';
import type { PendingClose, RegisterAsset, RegisterRowWire, ScanRecord } from './types';

/** Thrown when the server rejects the device access code — actionable, unlike a 500. */
export class AccessDeniedError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'AccessDeniedError';
  }
}

/**
 * The server's app registration has lost access — an expired client secret or
 * withdrawn admin consent. Retrying will not help; an admin has to act.
 */
export class ReauthRequiredError extends Error {
  constructor() {
    super("The Excel connection has lost access. Scans are safe — ask IT to check the app registration.");
    this.name = 'ReauthRequiredError';
  }
}

async function callApi<T>(body: unknown): Promise<T> {
  const res = await fetch(config.apiUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-stocktake-code': await getAccessCode(),
    },
    body: JSON.stringify(body),
  });

  // Read as text first. A crashed function returns Vercel's HTML error page, not
  // JSON, and blindly calling .json() would swallow the only clue we have.
  const raw = await res.text();
  let parsed: { error?: string } = {};
  try {
    parsed = raw ? JSON.parse(raw) : {};
  } catch {
    parsed = {};
  }

  if (!res.ok) {
    if (res.status === 401 && parsed.error === 'reauth_required') throw new ReauthRequiredError();
    if (res.status === 401 || res.status === 403) {
      throw new AccessDeniedError('Access code rejected. Check the code and try again.');
    }
    const detail =
      parsed.error ??
      (raw.trim() === ''
        ? 'empty response — the function crashed. Check the Vercel function logs.'
        : `${raw.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 160)}`);
    throw new Error(`${res.status}: ${detail}`);
  }
  return parsed as T;
}

function str(v: unknown): string {
  return v === null || v === undefined ? '' : String(v).trim();
}

/** Pure so it can be exercised without a network or a browser. */
export function mapRegisterRows(rows: RegisterRowWire[]): RegisterAsset[] {
  return rows
    .map((r) => {
      const facilityGroup = str(r['Facility Group']);
      const facility = str(r['Facility']);
      return {
        assetId: str(r['Asset Id']),
        oldAssetId: str(r['Old Asset Id']),
        barcode: str(r['Barcode']),
        serial: str(r['Serial Number']),
        assetName: str(r['Asset Name']),
        department: str(r['Department']),
        make: str(r['Make']),
        model: str(r['Model']),
        facility,
        facilityGroup,
        // Derived, not a workbook column. One string that a substring scope
        // match can hit at either granularity: "A BLOCK" or "01G04".
        expectedLocation: [facilityGroup, facility].filter((s) => s !== '').join(' / '),
        condition: str(r['Condition']),
        status: str(r['Status']),
        stocktakeStatus: str(r['Stocktake Status']),
        lastVerifiedDate: str(r['Last Verified Date']),
        lastVerifiedBy: str(r['Last Verified By']),
      };
    })
    .filter((r) => r.assetId !== '');
}

export async function fetchRegister(): Promise<RegisterAsset[]> {
  const { rows } = await callApi<{ rows: RegisterRowWire[] }>({ action: 'getRegister' });
  return mapRegisterRows(rows);
}

export async function submitScans(scans: ScanRecord[]): Promise<void> {
  await callApi({
    action: 'submitScans',
    scans: scans.map((s) => ({
      sessionId: s.sessionId,
      scanId: s.scanId,
      assetId: s.assetId,
      scannedValue: s.scannedValue,
      result: s.result,
      scannedBy: s.scannedBy,
      timestamp: s.timestamp,
      actualLocation: s.actualLocation,
      notes: s.notes,
    })),
  });
}

export async function submitClose(pc: PendingClose): Promise<void> {
  await callApi({ action: 'closeSession', ...pc });
}

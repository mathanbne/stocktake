import { NotSignedInError } from './auth';
import { getWorkbook } from './db';
import { appendScans, applyClose, getRegisterRows } from './graph';
import type { PendingClose, RegisterAsset, RegisterRowWire, ScanRecord } from './types';

/** No workbook has been chosen on this device yet. */
export class NoWorkbookError extends Error {
  constructor() {
    super('Choose the stocktake workbook before syncing.');
    this.name = 'NoWorkbookError';
  }
}

async function workbook() {
  const wb = await getWorkbook();
  if (!wb) throw new NoWorkbookError();
  return wb;
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
  const rows = await getRegisterRows(await workbook());
  return mapRegisterRows(rows as unknown as RegisterRowWire[]);
}

export async function submitScans(scans: ScanRecord[]): Promise<void> {
  await appendScans(await workbook(), scans);
}

export async function submitClose(pc: PendingClose): Promise<void> {
  await applyClose(await workbook(), pc);
}

export { NotSignedInError };

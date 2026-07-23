import { getToken } from './auth';
import { normalizeTag } from './capture';
import type { PendingClose, ScanRecord } from './types';

const GRAPH = 'https://graph.microsoft.com/v1.0';

export const REGISTER_TABLE = 'AssetRegister';
export const SCANLOG_TABLE = 'ScanLog';

/** Written into the register's Stocktake Status column. */
const CITED = 'Cited';
const MISSING = 'Missing';

/** Which workbook this device is pointed at. Chosen once, then remembered. */
export interface WorkbookRef {
  driveId: string;
  itemId: string;
  name: string;
}

export class WorkbookError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'WorkbookError';
  }
}

function base(wb: WorkbookRef): string {
  return `/drives/${wb.driveId}/items/${wb.itemId}`;
}

async function graph<T>(
  path: string,
  init: { method?: string; body?: unknown; sessionId?: string } = {},
): Promise<T> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${await getToken()}`,
    'Content-Type': 'application/json',
  };
  if (init.sessionId) headers['workbook-session-id'] = init.sessionId;

  const res = await fetch(`${GRAPH}${path}`, {
    method: init.method ?? 'GET',
    headers,
    body: init.body === undefined ? undefined : JSON.stringify(init.body),
  });

  if (!res.ok) {
    const text = await res.text();
    let detail = text;
    try {
      detail = (JSON.parse(text) as { error?: { message?: string } }).error?.message ?? text;
    } catch {
      /* not JSON — use the raw text */
    }
    throw new Error(`${res.status}: ${detail.slice(0, 200)}`);
  }
  return res.status === 204 ? (undefined as T) : ((await res.json()) as T);
}

// ---- picking a workbook ----

interface DriveItem {
  id: string;
  name: string;
  parentReference?: { driveId?: string; path?: string };
  lastModifiedDateTime?: string;
}

export interface WorkbookChoice extends WorkbookRef {
  location: string;
  modified: string;
}

/**
 * Microsoft Search rather than a drive listing, so results span OneDrive *and*
 * any SharePoint or Teams library the signed-in user can reach — which is where
 * a shared asset register usually lives.
 */
export async function findWorkbooks(term: string): Promise<WorkbookChoice[]> {
  const query = term.trim() === '' ? '.xlsx' : term.trim();
  const res = await graph<{
    value: { hitsContainers: { hits?: { resource: DriveItem }[] }[] }[];
  }>('/search/query', {
    method: 'POST',
    body: {
      requests: [
        { entityTypes: ['driveItem'], query: { queryString: query }, from: 0, size: 25 },
      ],
    },
  });

  const hits = res.value[0]?.hitsContainers[0]?.hits ?? [];
  return hits
    .map((h) => h.resource)
    .filter((r) => r.name?.toLowerCase().endsWith('.xlsx') && r.parentReference?.driveId)
    .map((r) => ({
      driveId: r.parentReference!.driveId!,
      itemId: r.id,
      name: r.name,
      location: (r.parentReference?.path ?? '').replace(/^\/drive\/root:?/, '') || 'OneDrive',
      modified: r.lastModifiedDateTime ?? '',
    }));
}

/** Confirms the picked file really is a stocktake workbook before it's accepted. */
export async function validateWorkbook(wb: WorkbookRef): Promise<void> {
  let tables: { value: { name: string }[] };
  try {
    tables = await graph(`${base(wb)}/workbook/tables?$select=name`);
  } catch (e) {
    throw new WorkbookError(
      `Could not open that file as a workbook. ${e instanceof Error ? e.message : ''}`,
    );
  }
  const names = tables.value.map((t) => t.name);
  const missing = [REGISTER_TABLE, SCANLOG_TABLE].filter((n) => !names.includes(n));
  if (missing.length > 0) {
    throw new WorkbookError(
      `That workbook has no ${missing.join(' or ')} table. Found: ${names.join(', ') || 'no tables'}.`,
    );
  }
}

// ---- reading tables ----

interface TableRows {
  headers: string[];
  index: Map<string, number>;
  rows: unknown[][];
  dataBodyAddress: string | null;
}

function headerMap(headers: string[]): Map<string, number> {
  const m = new Map<string, number>();
  headers.forEach((h, i) => m.set(h.trim().toLowerCase(), i));
  return m;
}

function columnIndex(map: Map<string, number>, header: string): number {
  const i = map.get(header.toLowerCase());
  if (i === undefined) throw new WorkbookError(`The workbook is missing the "${header}" column.`);
  return i;
}

async function readTable(wb: WorkbookRef, table: string, sessionId?: string): Promise<TableRows> {
  const t = `${base(wb)}/workbook/tables('${table}')`;

  const header = await graph<{ values: unknown[][] }>(`${t}/headerRowRange`, { sessionId });
  const headers = (header.values[0] ?? []).map((h) => String(h ?? '').trim());

  let dataBodyAddress: string | null = null;
  try {
    const body = await graph<{ address: string }>(`${t}/dataBodyRange`, { sessionId });
    dataBodyAddress = body.address ?? null;
  } catch {
    dataBodyAddress = null; // header-only table
  }

  // Graph pages large tables; a register of any real size needs this.
  const rows: unknown[][] = [];
  let path: string | null = `${t}/rows`;
  while (path) {
    const page: { value: { values: unknown[][] }[]; '@odata.nextLink'?: string } = await graph(path, {
      sessionId,
    });
    for (const r of page.value) rows.push(r.values[0] ?? []);
    const next = page['@odata.nextLink'];
    path = next ? next.replace(GRAPH, '') : null;
  }

  return { headers, index: headerMap(headers), rows, dataBodyAddress };
}

export async function getRegisterRows(wb: WorkbookRef): Promise<Record<string, string>[]> {
  const { headers, rows } = await readTable(wb, REGISTER_TABLE);
  return rows.map((r) => {
    const o: Record<string, string> = {};
    headers.forEach((h, i) => {
      o[h] = r[i] === null || r[i] === undefined ? '' : String(r[i]);
    });
    return o;
  });
}

// ---- writing ----

/** 0-based column index → Excel letters (0 → A, 26 → AA). */
function columnLetter(index: number): string {
  let n = index;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** `Register!A2:O41` → 2. */
function firstDataRow(address: string): number {
  const m = /![A-Z]+(\d+):/.exec(address);
  if (!m) throw new WorkbookError(`Could not read the table's address: ${address}`);
  return Number(m[1]);
}

function sheetNameFrom(address: string): string {
  const i = address.indexOf('!');
  return i === -1 ? address : address.slice(0, i).replace(/^'|'$/g, '');
}

async function openSession(wb: WorkbookRef): Promise<string | undefined> {
  try {
    const r = await graph<{ id: string }>(`${base(wb)}/workbook/createSession`, {
      method: 'POST',
      body: { persistChanges: true },
    });
    return r.id;
  } catch {
    return undefined; // sessionless still works, just slower
  }
}

async function closeSessionHandle(wb: WorkbookRef, sessionId?: string): Promise<void> {
  if (!sessionId) return;
  try {
    await graph(`${base(wb)}/workbook/closeSession`, { method: 'POST', sessionId });
  } catch {
    // Sessions expire by themselves; a failed close isn't worth failing on.
  }
}

/**
 * Idempotent append. Every scan carries a client-generated Scan ID, so a lost
 * response costs nothing: the identical batch arrives again and rows already
 * present are skipped.
 */
export async function appendScans(wb: WorkbookRef, scans: ScanRecord[]): Promise<void> {
  if (scans.length === 0) return;
  const sessionId = await openSession(wb);
  try {
    const { headers, index, rows } = await readTable(wb, SCANLOG_TABLE, sessionId);
    const scanIdCol = columnIndex(index, 'Scan ID');
    const existing = new Set(rows.map((r) => String(r[scanIdCol] ?? '')));

    const fresh = scans.filter((s) => !existing.has(s.scanId));
    if (fresh.length === 0) return;

    const field: Record<string, (s: ScanRecord) => string> = {
      'session id': (s) => s.sessionId,
      'scan id': (s) => s.scanId,
      'asset id': (s) => s.assetId,
      'scanned value': (s) => s.scannedValue,
      result: (s) => s.result,
      'scanned by': (s) => s.scannedBy,
      timestamp: (s) => s.timestamp,
      'actual location': (s) => s.actualLocation,
      notes: (s) => s.notes,
    };

    await graph(`${base(wb)}/workbook/tables('${SCANLOG_TABLE}')/rows/add`, {
      method: 'POST',
      sessionId,
      body: {
        values: fresh.map((s) =>
          headers.map((h) => {
            const get = field[h.trim().toLowerCase()];
            return get ? get(s) : '';
          }),
        ),
      },
    });
  } finally {
    await closeSessionHandle(wb, sessionId);
  }
}

/**
 * Graph has no "update row by key", so the register is read once to build an
 * Asset Id → spreadsheet row map, then each asset's three stocktake cells are
 * PATCHed as one contiguous range. Range writes are idempotent, which is what
 * lets a half-finished close be replayed verbatim.
 */
export async function applyClose(wb: WorkbookRef, pc: PendingClose): Promise<void> {
  const sessionId = await openSession(wb);
  try {
    const { index, rows, dataBodyAddress } = await readTable(wb, REGISTER_TABLE, sessionId);
    if (!dataBodyAddress) throw new WorkbookError('The AssetRegister table has no rows.');

    const assetIdCol = columnIndex(index, 'Asset Id');
    const statusCol = columnIndex(index, 'Stocktake Status');
    const dateCol = columnIndex(index, 'Last Verified Date');
    const byCol = columnIndex(index, 'Last Verified By');

    const first = Math.min(statusCol, dateCol, byCol);
    const last = Math.max(statusCol, dateCol, byCol);
    if (last - first !== 2) {
      throw new WorkbookError(
        'Stocktake Status, Last Verified Date and Last Verified By must be three adjacent columns.',
      );
    }

    const sheet = sheetNameFrom(dataBodyAddress);
    const startRow = firstDataRow(dataBodyAddress);
    const rowOf = new Map<string, number>();
    rows.forEach((r, i) => {
      const key = normalizeTag(String(r[assetIdCol] ?? ''));
      if (key !== '') rowOf.set(key, startRow + i);
    });

    const cells = (status: string, date: string, by: string): string[] => {
      const out = ['', '', ''];
      out[statusCol - first] = status;
      out[dateCol - first] = date;
      out[byCol - first] = by;
      return out;
    };

    const requests: { id: string; method: string; url: string; body: unknown }[] = [];
    const push = (assetId: string, values: string[]) => {
      const row = rowOf.get(normalizeTag(assetId));
      if (row === undefined) return; // no longer in the register
      const address = `${columnLetter(first)}${row}:${columnLetter(last)}${row}`;
      requests.push({
        id: `r${requests.length}`,
        method: 'PATCH',
        url: `${base(wb)}/workbook/worksheets('${encodeURIComponent(sheet)}')/range(address='${address}')`,
        body: { values: [values] },
      });
    };

    for (const u of pc.sightedUpdates) push(u.assetId, cells(CITED, u.verifiedDate, u.verifiedBy));
    for (const id of pc.missingAssetIds) push(id, cells(MISSING, '', ''));

    // Graph caps $batch at 20 sub-requests.
    for (let i = 0; i < requests.length; i += 20) {
      const chunk = requests.slice(i, i + 20);
      const out = await graph<{ responses: { status: number; body?: unknown }[] }>('/$batch', {
        method: 'POST',
        sessionId,
        body: {
          requests: chunk.map((r) => ({
            ...r,
            headers: {
              'Content-Type': 'application/json',
              ...(sessionId ? { 'workbook-session-id': sessionId } : {}),
            },
          })),
        },
      });
      const failed = out.responses.filter((r) => r.status >= 400);
      if (failed.length > 0) {
        throw new Error(
          `${failed.length} of ${chunk.length} register updates failed: ${JSON.stringify(failed[0].body).slice(0, 200)}`,
        );
      }
    }
  } finally {
    await closeSessionHandle(wb, sessionId);
  }
}

import type { VercelRequest, VercelResponse } from '@vercel/node';
import {
  ConfigError,
  ReauthRequired,
  closeWorkbookSession,
  columnIndex,
  columnLetter,
  firstDataRow,
  getAccessToken,
  graph,
  graphBatch,
  headerMap,
  openWorkbookSession,
  workbookBase,
  type BatchRequest,
  type GraphCtx,
  // Explicit .js extension: package.json sets "type": "module", so Node's ESM
  // loader resolves this at runtime and an extensionless specifier fails the
  // whole module load — before the handler runs, which surfaces as a bare 500.
} from './_graph.js';

const REGISTER_TABLE = 'AssetRegister';
const SCANLOG_TABLE = 'ScanLog';

/** Written into the register's Stocktake Status column. */
const CITED = 'Cited';
const MISSING = 'Missing';

interface IncomingScan {
  sessionId: string;
  scanId: string;
  assetId: string;
  scannedValue: string;
  result: string;
  scannedBy: string;
  timestamp: string;
  actualLocation: string;
  notes: string;
}

/** Must stay in lockstep with normalizeTag() in src/capture.ts. */
function normalize(raw: unknown): string {
  return String(raw ?? '')
    .replace(/\s+/g, '')
    .replace(/^['#]+/, '')
    .toUpperCase();
}

/** Timing-safe enough for a short shared code, and never short-circuits on length. */
function codeMatches(supplied: string, expected: string): boolean {
  let diff = supplied.length ^ expected.length;
  for (let i = 0; i < Math.max(supplied.length, expected.length); i++) {
    diff |= (supplied.charCodeAt(i) || 0) ^ (expected.charCodeAt(i) || 0);
  }
  return diff === 0;
}

interface TableRows {
  headers: string[];
  index: Map<string, number>;
  rows: unknown[][];
  /** e.g. `Register!A2:O41`; null when the table has no data rows yet. */
  dataBodyAddress: string | null;
}

async function readTable(ctx: GraphCtx, table: string): Promise<TableRows> {
  const base = `${workbookBase()}/workbook/tables('${table}')`;

  const header = await graph<{ values: unknown[][] }>(ctx, `${base}/headerRowRange`);
  const headers = (header.values[0] ?? []).map((h) => String(h ?? '').trim());

  let dataBodyAddress: string | null = null;
  try {
    const body = await graph<{ address: string }>(ctx, `${base}/dataBodyRange`);
    dataBodyAddress = body.address ?? null;
  } catch {
    dataBodyAddress = null; // empty table
  }

  // Graph pages large tables; a register of any real size will need this.
  const rows: unknown[][] = [];
  let path: string | null = `${base}/rows`;
  while (path) {
    const page: { value: { values: unknown[][] }[]; '@odata.nextLink'?: string } = await graph(ctx, path);
    for (const r of page.value) rows.push(r.values[0] ?? []);
    const next = page['@odata.nextLink'];
    path = next ? next.replace('https://graph.microsoft.com/v1.0', '') : null;
  }

  return { headers, index: headerMap(headers), rows, dataBodyAddress };
}

function sheetNameFrom(address: string): string {
  const i = address.indexOf('!');
  return i === -1 ? address : address.slice(0, i).replace(/^'|'$/g, '');
}

// ---- actions ----

async function getRegister(ctx: GraphCtx): Promise<{ rows: Record<string, string>[] }> {
  const { headers, rows } = await readTable(ctx, REGISTER_TABLE);
  return {
    rows: rows.map((r) => {
      const o: Record<string, string> = {};
      headers.forEach((h, i) => {
        o[h] = r[i] === null || r[i] === undefined ? '' : String(r[i]);
      });
      return o;
    }),
  };
}

/**
 * Idempotent append. Every scan carries a client-generated Scan ID, so a lost
 * response costs nothing: the identical batch arrives again and each row that is
 * already present is skipped.
 */
async function submitScans(ctx: GraphCtx, scans: IncomingScan[]): Promise<{ accepted: number }> {
  if (scans.length === 0) return { accepted: 0 };

  const { headers, index, rows } = await readTable(ctx, SCANLOG_TABLE);
  const scanIdCol = columnIndex(index, 'Scan ID');
  const existing = new Set(rows.map((r) => String(r[scanIdCol] ?? '')));

  const fresh = scans.filter((s) => !existing.has(s.scanId));
  if (fresh.length === 0) return { accepted: 0 };

  const field: Record<string, (s: IncomingScan) => string> = {
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

  const values = fresh.map((s) =>
    headers.map((h) => {
      const get = field[h.trim().toLowerCase()];
      return get ? get(s) : '';
    }),
  );

  await graph(ctx, `${workbookBase()}/workbook/tables('${SCANLOG_TABLE}')/rows/add`, {
    method: 'POST',
    body: { values },
  });
  return { accepted: fresh.length };
}

/**
 * Graph has no "update row by key", so the register is read once to build an
 * Asset Id → spreadsheet row map, then each asset's three stocktake cells are
 * PATCHed as a contiguous range. Range writes are idempotent, which is what lets
 * a half-finished close be replayed verbatim.
 */
async function closeSession(
  ctx: GraphCtx,
  missingAssetIds: string[],
  sightedUpdates: { assetId: string; verifiedDate: string; verifiedBy: string }[],
): Promise<{ closed: true; updated: number }> {
  const { index, rows, dataBodyAddress } = await readTable(ctx, REGISTER_TABLE);
  if (!dataBodyAddress) throw new ConfigError('The AssetRegister table has no data rows.');

  const assetIdCol = columnIndex(index, 'Asset Id');
  const statusCol = columnIndex(index, 'Stocktake Status');
  const dateCol = columnIndex(index, 'Last Verified Date');
  const byCol = columnIndex(index, 'Last Verified By');

  // The three are written as one range, which assumes they sit side by side —
  // exactly how the setup instructions add them.
  const first = Math.min(statusCol, dateCol, byCol);
  const last = Math.max(statusCol, dateCol, byCol);
  if (last - first !== 2) {
    throw new ConfigError(
      'Stocktake Status, Last Verified Date and Last Verified By must be three adjacent columns.',
    );
  }

  const sheet = sheetNameFrom(dataBodyAddress);
  const startRow = firstDataRow(dataBodyAddress);
  const rowOf = new Map<string, number>();
  rows.forEach((r, i) => {
    const key = normalize(r[assetIdCol]);
    if (key !== '') rowOf.set(key, startRow + i);
  });

  const cells = (status: string, date: string, by: string): string[] => {
    const out = ['', '', ''];
    out[statusCol - first] = status;
    out[dateCol - first] = date;
    out[byCol - first] = by;
    return out;
  };

  const requests: BatchRequest[] = [];
  const push = (assetId: string, values: string[]) => {
    const row = rowOf.get(normalize(assetId));
    if (row === undefined) return; // no longer in the register — nothing to update
    const address = `${columnLetter(first)}${row}:${columnLetter(last)}${row}`;
    requests.push({
      id: `r${requests.length}`,
      method: 'PATCH',
      url: `${workbookBase()}/workbook/worksheets('${encodeURIComponent(sheet)}')/range(address='${address}')`,
      body: { values: [values] },
    });
  };

  for (const u of sightedUpdates) push(u.assetId, cells(CITED, u.verifiedDate, u.verifiedBy));
  for (const id of missingAssetIds) push(id, cells(MISSING, '', ''));

  await graphBatch(ctx, requests);
  return { closed: true, updated: requests.length };
}

// ---- handler ----

export default async function handler(req: VercelRequest, res: VercelResponse): Promise<void> {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'POST only' });
    return;
  }

  const expected = process.env.STOCKTAKE_ACCESS_CODE ?? '';
  const supplied = String(req.headers['x-stocktake-code'] ?? '');
  if (expected === '' || !codeMatches(supplied, expected)) {
    res.status(401).json({ error: 'invalid_access_code' });
    return;
  }

  const body = (req.body ?? {}) as Record<string, unknown>;
  const action = String(body.action ?? '');
  let ctx: GraphCtx | null = null;

  try {
    const token = await getAccessToken();
    // Reads don't need a workbook session; writes benefit from one.
    const sessionId = action === 'getRegister' ? undefined : await openWorkbookSession(token);
    ctx = { token, sessionId };

    switch (action) {
      case 'getRegister':
        res.status(200).json(await getRegister(ctx));
        return;
      case 'submitScans':
        res.status(200).json(await submitScans(ctx, (body.scans ?? []) as IncomingScan[]));
        return;
      case 'closeSession':
        res.status(200).json(
          await closeSession(
            ctx,
            (body.missingAssetIds ?? []) as string[],
            (body.sightedUpdates ?? []) as { assetId: string; verifiedDate: string; verifiedBy: string }[],
          ),
        );
        return;
      default:
        res.status(400).json({ error: `Unknown action "${action}"` });
        return;
    }
  } catch (e) {
    if (e instanceof ReauthRequired) {
      res.status(401).json({ error: 'reauth_required', detail: e.message });
      return;
    }
    if (e instanceof ConfigError) {
      res.status(500).json({ error: e.message });
      return;
    }
    res.status(502).json({ error: e instanceof Error ? e.message : 'Upstream failure' });
  } finally {
    if (ctx) await closeWorkbookSession(ctx);
  }
}

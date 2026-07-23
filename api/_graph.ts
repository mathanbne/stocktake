/**
 * Everything that touches Microsoft lives on the server. The phone never holds a
 * credential, never signs in, and never learns the workbook's identity — it only
 * knows about /api/excel.
 *
 * Auth is app-only (client credentials): the app registration *is* the identity,
 * so there is no user, no refresh token, and nothing to re-consent when a person
 * leaves. This is only available on a work/school tenant.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';

/** The app registration's access has broken — an admin has to fix it, not a scanner. */
export class ReauthRequired extends Error {
  constructor(detail: string) {
    super(detail);
    this.name = 'ReauthRequired';
  }
}

export class ConfigError extends Error {
  constructor(detail: string) {
    super(detail);
    this.name = 'ConfigError';
  }
}

function required(name: string): string {
  const v = process.env[name];
  if (!v) throw new ConfigError(`Missing environment variable ${name}`);
  return v;
}

// ---- access token ----

/** Warm-instance cache. Saves a token round trip on most invocations. */
let cached: { token: string; expiresAt: number } | null = null;

export async function getAccessToken(): Promise<string> {
  if (cached && Date.now() < cached.expiresAt) return cached.token;

  const tenant = required('MS_TENANT_ID');
  const res = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: required('MS_CLIENT_ID'),
      client_secret: required('MS_CLIENT_SECRET'),
      grant_type: 'client_credentials',
      scope: 'https://graph.microsoft.com/.default',
    }).toString(),
  });

  const body = (await res.json()) as Record<string, unknown>;
  if (!res.ok) {
    const detail = `${body.error ?? res.status}: ${body.error_description ?? 'token request failed'}`;
    // An expired client secret or withdrawn admin consent is an operational
    // problem, not a transient one — flag it distinctly so the app can say so.
    if (res.status === 400 || res.status === 401) throw new ReauthRequired(detail);
    throw new Error(detail);
  }

  const token = body as unknown as { access_token: string; expires_in: number };
  cached = {
    token: token.access_token,
    expiresAt: Date.now() + Math.max(0, token.expires_in - 300) * 1000,
  };
  return cached.token;
}

// ---- Graph plumbing ----

/**
 * App-only tokens have no `/me`, so the drive must be named explicitly.
 * See the README for the two Graph queries that produce these two ids.
 */
export function workbookBase(): string {
  return `/drives/${required('EXCEL_DRIVE_ID')}/items/${required('EXCEL_ITEM_ID')}`;
}

export interface GraphCtx {
  token: string;
  sessionId?: string;
}

export async function graph<T>(
  ctx: GraphCtx,
  path: string,
  init: { method?: string; body?: unknown } = {},
): Promise<T> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${ctx.token}`,
    'Content-Type': 'application/json',
  };
  if (ctx.sessionId) headers['workbook-session-id'] = ctx.sessionId;

  const res = await fetch(`${GRAPH}${path}`, {
    method: init.method ?? 'GET',
    headers,
    body: init.body === undefined ? undefined : JSON.stringify(init.body),
  });
  if (!res.ok) {
    const text = await res.text();
    if (res.status === 401 || res.status === 403) {
      throw new ReauthRequired(`Graph refused the app registration: ${text}`);
    }
    throw new Error(`Graph ${res.status} on ${path}: ${text}`);
  }
  return res.status === 204 ? (undefined as T) : ((await res.json()) as T);
}

/**
 * A persistent workbook session batches Excel's own recalculation instead of
 * reloading the file per request. Best-effort: if it can't be created we fall
 * back to sessionless calls, which are correct, just slower.
 */
export async function openWorkbookSession(token: string): Promise<string | undefined> {
  try {
    const r = await graph<{ id: string }>({ token }, `${workbookBase()}/workbook/createSession`, {
      method: 'POST',
      body: { persistChanges: true },
    });
    return r.id;
  } catch {
    return undefined;
  }
}

export async function closeWorkbookSession(ctx: GraphCtx): Promise<void> {
  if (!ctx.sessionId) return;
  try {
    await graph(ctx, `${workbookBase()}/workbook/closeSession`, { method: 'POST' });
  } catch {
    // Sessions expire on their own; a failed close is not worth failing the request.
  }
}

// ---- batching ----

export interface BatchRequest {
  id: string;
  method: string;
  url: string;
  body?: unknown;
}

/** Graph caps $batch at 20 sub-requests; chunking is the caller's only concern. */
export async function graphBatch(ctx: GraphCtx, requests: BatchRequest[]): Promise<void> {
  for (let i = 0; i < requests.length; i += 20) {
    const chunk = requests.slice(i, i + 20);
    const payload = {
      requests: chunk.map((r) => ({
        id: r.id,
        method: r.method,
        url: r.url,
        body: r.body,
        headers: {
          'Content-Type': 'application/json',
          ...(ctx.sessionId ? { 'workbook-session-id': ctx.sessionId } : {}),
        },
      })),
    };
    const out = await graph<{ responses: { id: string; status: number; body?: unknown }[] }>(
      ctx,
      '/$batch',
      { method: 'POST', body: payload },
    );
    const failed = out.responses.filter((r) => r.status >= 400);
    if (failed.length > 0) {
      throw new Error(
        `Batch write failed for ${failed.length} of ${chunk.length}: ${JSON.stringify(failed[0].body)}`,
      );
    }
  }
}

// ---- spreadsheet helpers ----

/** 0-based column index → Excel letters (0 → A, 26 → AA). */
export function columnLetter(index: number): string {
  let n = index;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** `Register!A2:O41` → 2. The first spreadsheet row holding table data. */
export function firstDataRow(dataBodyAddress: string): number {
  const m = /![A-Z]+(\d+):/.exec(dataBodyAddress);
  if (!m) throw new Error(`Could not parse table address: ${dataBodyAddress}`);
  return Number(m[1]);
}

export function headerMap(headers: string[]): Map<string, number> {
  const m = new Map<string, number>();
  headers.forEach((h, i) => m.set(String(h).trim().toLowerCase(), i));
  return m;
}

export function columnIndex(map: Map<string, number>, header: string): number {
  const i = map.get(header.toLowerCase());
  if (i === undefined) throw new ConfigError(`Workbook is missing the "${header}" column.`);
  return i;
}

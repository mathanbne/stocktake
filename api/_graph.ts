/**
 * Everything that touches Microsoft lives on the server. The phone never holds a
 * credential, never signs in, and never learns the workbook's identity — it only
 * knows about /api/excel.
 *
 * Auth is **delegated**: the app acts as you. You sign in once from a laptop at
 * /api/auth/start, and the resulting refresh token is stored server-side and used
 * from then on. Nothing is granted that you couldn't already open yourself, which
 * is why this needs no admin consent — unlike app-only auth, which would give the
 * app its own tenant-wide identity.
 *
 * Consequence worth knowing: every write lands in the workbook's version history
 * under *your* name, and the connection dies if your account is disabled or your
 * password is reset.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';

/**
 * Delegated Files.ReadWrite.All reaches both OneDrive for Business and any
 * SharePoint library you can already open, and — unlike its application-level
 * namesake — does not require admin consent. offline_access is what makes the
 * refresh token possible.
 */
export const SCOPES = 'Files.ReadWrite.All offline_access';

/** `organizations` accepts any work/school account without pinning a tenant id. */
export function authority(): string {
  return process.env.MS_TENANT_ID || 'organizations';
}

/** The stored sign-in has broken — someone has to sign in again at /api/auth/start. */
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

export function required(name: string): string {
  const v = process.env[name];
  if (!v) throw new ConfigError(`Missing environment variable ${name}`);
  return v;
}

// ---- refresh-token store ----

const KV_KEY = 'stocktake:ms_refresh_token';

async function kv(command: unknown[]): Promise<{ result: string | null }> {
  const res = await fetch(required('KV_REST_API_URL'), {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${required('KV_REST_API_TOKEN')}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(command),
  });
  if (!res.ok) throw new Error(`KV ${res.status}: ${await res.text()}`);
  return (await res.json()) as { result: string | null };
}

export async function readRefreshToken(): Promise<string | null> {
  return (await kv(['GET', KV_KEY])).result;
}

export async function writeRefreshToken(token: string): Promise<void> {
  await kv(['SET', KV_KEY, token]);
}

// ---- access tokens ----

/** Warm-instance cache. Saves a token round trip on most invocations. */
let cached: { token: string; expiresAt: number } | null = null;

export async function requestToken(form: Record<string, string>): Promise<{
  access_token: string;
  expires_in: number;
  refresh_token?: string;
}> {
  const res = await fetch(`https://login.microsoftonline.com/${authority()}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(form).toString(),
  });
  const body = (await res.json()) as Record<string, unknown>;
  if (!res.ok) {
    const detail = `${body.error ?? res.status}: ${body.error_description ?? 'token request failed'}`;
    // A dead refresh token needs a human to sign in again — never a retry.
    if (res.status === 400 || res.status === 401) throw new ReauthRequired(detail);
    throw new Error(detail);
  }
  return body as unknown as { access_token: string; expires_in: number; refresh_token?: string };
}

export async function getAccessToken(): Promise<string> {
  if (cached && Date.now() < cached.expiresAt) return cached.token;

  const refresh = await readRefreshToken();
  if (!refresh) {
    throw new ReauthRequired('Nobody has signed in yet. Visit /api/auth/start to connect Excel.');
  }

  const token = await requestToken({
    client_id: required('MS_CLIENT_ID'),
    client_secret: required('MS_CLIENT_SECRET'),
    grant_type: 'refresh_token',
    refresh_token: refresh,
    scope: SCOPES,
  });

  // Entra rotates the refresh token on redemption; persisting the new one is
  // mandatory, or the next cold start is locked out.
  if (token.refresh_token) await writeRefreshToken(token.refresh_token);

  cached = {
    token: token.access_token,
    expiresAt: Date.now() + Math.max(0, token.expires_in - 300) * 1000,
  };
  return cached.token;
}

// ---- Graph plumbing ----

/**
 * `/me/drive` is the signed-in user's OneDrive for Business. Set EXCEL_DRIVE_ID
 * instead when the workbook lives in a SharePoint or Teams document library.
 */
export function workbookBase(): string {
  const item = required('EXCEL_ITEM_ID');
  const drive = process.env.EXCEL_DRIVE_ID;
  return drive ? `/drives/${drive}/items/${item}` : `/me/drive/items/${item}`;
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
    if (res.status === 401) throw new ReauthRequired(`Graph rejected the token: ${text}`);
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

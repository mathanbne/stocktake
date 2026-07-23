/**
 * Everything that touches Microsoft lives on the server. The phone never holds a
 * credential, never signs in, and never learns the workbook's identity — it only
 * knows about /api/excel.
 *
 * Two auth modes, chosen by which env vars are present:
 *
 *   work M365   MS_TENANT_ID set  → client credentials (app-only). No user, no
 *                                   refresh token, no extra storage.
 *   personal    MS_TENANT_ID unset → delegated refresh-token flow. Consumer
 *                                   OneDrive does not support app-only, and MSA
 *                                   refresh tokens rotate on every use, so the
 *                                   current one must be persisted (KV).
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';
const DELEGATED_SCOPE = 'Files.ReadWrite offline_access';

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

export function isAppOnly(): boolean {
  const t = process.env.MS_TENANT_ID;
  return !!t && t !== 'consumers';
}

// ---- token store (KV, personal path only) ----

const KV_KEY = 'stocktake:ms_refresh_token';

async function kv(command: unknown[]): Promise<{ result: string | null }> {
  const url = required('KV_REST_API_URL');
  const token = required('KV_REST_API_TOKEN');
  const res = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
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

async function requestToken(form: Record<string, string>, tenant: string): Promise<{
  access_token: string;
  expires_in: number;
  refresh_token?: string;
}> {
  const res = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(form).toString(),
  });
  const body = (await res.json()) as Record<string, unknown>;
  if (!res.ok) {
    const detail = `${body.error ?? res.status}: ${body.error_description ?? 'token request failed'}`;
    // A dead refresh token is the owner's problem to fix, not a transient error —
    // surface it distinctly so the client can say "reconnect" instead of "retry".
    if (res.status === 400 || res.status === 401) throw new ReauthRequired(detail);
    throw new Error(detail);
  }
  return body as unknown as { access_token: string; expires_in: number; refresh_token?: string };
}

export async function getAccessToken(): Promise<string> {
  if (cached && Date.now() < cached.expiresAt) return cached.token;

  const clientId = required('MS_CLIENT_ID');
  const clientSecret = required('MS_CLIENT_SECRET');
  let token: { access_token: string; expires_in: number; refresh_token?: string };

  if (isAppOnly()) {
    token = await requestToken(
      {
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: 'client_credentials',
        scope: 'https://graph.microsoft.com/.default',
      },
      required('MS_TENANT_ID'),
    );
  } else {
    const refresh = await readRefreshToken();
    if (!refresh) {
      throw new ReauthRequired('No stored refresh token. Run the one-time consent at /api/auth/start.');
    }
    token = await requestToken(
      {
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: 'refresh_token',
        refresh_token: refresh,
        scope: DELEGATED_SCOPE,
      },
      'consumers',
    );
    // MSA rotates the refresh token on every use — persisting the new one is
    // mandatory, or the next cold start is locked out.
    if (token.refresh_token) await writeRefreshToken(token.refresh_token);
  }

  cached = {
    token: token.access_token,
    expiresAt: Date.now() + Math.max(0, token.expires_in - 300) * 1000,
  };
  return cached.token;
}

// ---- Graph plumbing ----

/** `/drives/{id}/items/{id}` app-only; `/me/drive/items/{id}` for personal. */
export function workbookBase(): string {
  const item = required('EXCEL_ITEM_ID');
  const drive = process.env.EXCEL_DRIVE_ID;
  if (isAppOnly() && !drive) {
    throw new ConfigError('EXCEL_DRIVE_ID is required in app-only mode — there is no /me to resolve.');
  }
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

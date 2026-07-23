import type { VercelRequest, VercelResponse } from '@vercel/node';
import { SCOPES, authority } from '../_graph.js';

/**
 * One-time sign-in, run by you from a laptop — never from the scanning phone.
 * The token it produces is what /api/excel uses from then on, which is what
 * keeps the phones credential-free and able to scan offline.
 *
 *   /api/auth/start?code=<STOCKTAKE_ACCESS_CODE>
 */
export default function handler(req: VercelRequest, res: VercelResponse): void {
  const expected = process.env.STOCKTAKE_ACCESS_CODE ?? '';
  if (expected === '' || String(req.query.code ?? '') !== expected) {
    res.status(401).json({ error: 'invalid_access_code' });
    return;
  }

  const clientId = process.env.MS_CLIENT_ID;
  if (!clientId) {
    res.status(500).json({ error: 'Missing MS_CLIENT_ID' });
    return;
  }

  const state = crypto.randomUUID();
  const url = new URL(`https://login.microsoftonline.com/${authority()}/oauth2/v2.0/authorize`);
  url.searchParams.set('client_id', clientId);
  url.searchParams.set('response_type', 'code');
  url.searchParams.set('redirect_uri', `https://${req.headers.host}/api/auth/callback`);
  url.searchParams.set('response_mode', 'query');
  url.searchParams.set('scope', SCOPES);
  url.searchParams.set('state', state);
  // Force the consent screen so offline_access is definitely granted — without a
  // refresh token the connection would silently die within the hour.
  url.searchParams.set('prompt', 'consent');

  res.setHeader(
    'Set-Cookie',
    `stocktake_state=${state}; Path=/api/auth; HttpOnly; Secure; SameSite=Lax; Max-Age=600`,
  );
  res.redirect(302, url.toString());
}

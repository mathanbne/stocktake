import type { VercelRequest, VercelResponse } from '@vercel/node';
import { isAppOnly } from '../_graph';

/**
 * One-time consent, run by the workbook owner from a laptop — never from the
 * scanning phone. Needed only on the personal-OneDrive path: consumer accounts
 * can't use app-only auth, so the server needs a delegated refresh token to hold
 * on the owner's behalf.
 *
 *   /api/auth/start?code=<STOCKTAKE_ACCESS_CODE>
 */
export default function handler(req: VercelRequest, res: VercelResponse): void {
  if (isAppOnly()) {
    res.status(400).json({ error: 'MS_TENANT_ID is set — app-only auth needs no consent flow.' });
    return;
  }

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

  const origin = `https://${req.headers.host}`;
  const state = crypto.randomUUID();

  const url = new URL('https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize');
  url.searchParams.set('client_id', clientId);
  url.searchParams.set('response_type', 'code');
  url.searchParams.set('redirect_uri', `${origin}/api/auth/callback`);
  url.searchParams.set('response_mode', 'query');
  url.searchParams.set('scope', 'Files.ReadWrite offline_access');
  url.searchParams.set('state', state);
  // Force the consent screen so offline_access is definitely granted.
  url.searchParams.set('prompt', 'consent');

  res.setHeader('Set-Cookie', `stocktake_state=${state}; Path=/api/auth; HttpOnly; Secure; SameSite=Lax; Max-Age=600`);
  res.redirect(302, url.toString());
}

import type { VercelRequest, VercelResponse } from '@vercel/node';
import { SCOPES, requestToken, writeRefreshToken } from '../_graph.js';

function cookie(header: string | undefined, name: string): string {
  for (const part of (header ?? '').split(';')) {
    const [k, ...v] = part.trim().split('=');
    if (k === name) return v.join('=');
  }
  return '';
}

/** Redirect target for /api/auth/start. Exchanges the code and stores the refresh token. */
export default async function handler(req: VercelRequest, res: VercelResponse): Promise<void> {
  const state = String(req.query.state ?? '');
  if (state === '' || state !== cookie(req.headers.cookie, 'stocktake_state')) {
    res.status(400).send('State mismatch — start again at /api/auth/start.');
    return;
  }

  const code = String(req.query.code ?? '');
  if (!code) {
    res
      .status(400)
      .send(`Sign-in failed: ${req.query.error_description ?? req.query.error ?? 'no code returned'}`);
    return;
  }

  try {
    const token = await requestToken({
      client_id: process.env.MS_CLIENT_ID ?? '',
      client_secret: process.env.MS_CLIENT_SECRET ?? '',
      grant_type: 'authorization_code',
      code,
      redirect_uri: `https://${req.headers.host}/api/auth/callback`,
      scope: SCOPES,
    });
    if (!token.refresh_token) {
      res.status(500).send('No refresh token returned — was offline_access granted?');
      return;
    }
    await writeRefreshToken(token.refresh_token);
  } catch (e) {
    res.status(500).send(`Could not complete sign-in: ${e instanceof Error ? e.message : 'unknown'}`);
    return;
  }

  res.setHeader('Set-Cookie', 'stocktake_state=; Path=/api/auth; HttpOnly; Secure; SameSite=Lax; Max-Age=0');
  res.status(200).send('Excel connected. You can close this tab — the app can now sync.');
}

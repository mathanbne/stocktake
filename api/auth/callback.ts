import type { VercelRequest, VercelResponse } from '@vercel/node';
import { isAppOnly, writeRefreshToken } from '../_graph';

function cookie(header: string | undefined, name: string): string {
  for (const part of (header ?? '').split(';')) {
    const [k, ...v] = part.trim().split('=');
    if (k === name) return v.join('=');
  }
  return '';
}

/** Redirect target for /api/auth/start. Exchanges the code and stores the refresh token. */
export default async function handler(req: VercelRequest, res: VercelResponse): Promise<void> {
  if (isAppOnly()) {
    res.status(400).json({ error: 'MS_TENANT_ID is set — app-only auth needs no consent flow.' });
    return;
  }

  const state = String(req.query.state ?? '');
  if (state === '' || state !== cookie(req.headers.cookie, 'stocktake_state')) {
    res.status(400).json({ error: 'State mismatch — start again at /api/auth/start.' });
    return;
  }

  const code = String(req.query.code ?? '');
  if (!code) {
    res.status(400).json({
      error: `Authorization failed: ${req.query.error_description ?? req.query.error ?? 'no code returned'}`,
    });
    return;
  }

  const origin = `https://${req.headers.host}`;
  const token = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: process.env.MS_CLIENT_ID ?? '',
      client_secret: process.env.MS_CLIENT_SECRET ?? '',
      grant_type: 'authorization_code',
      code,
      redirect_uri: `${origin}/api/auth/callback`,
      scope: 'Files.ReadWrite offline_access',
    }).toString(),
  });

  const body = (await token.json()) as { refresh_token?: string; error_description?: string };
  if (!token.ok || !body.refresh_token) {
    res.status(500).json({ error: body.error_description ?? 'No refresh token returned' });
    return;
  }

  await writeRefreshToken(body.refresh_token);
  res.setHeader('Set-Cookie', 'stocktake_state=; Path=/api/auth; HttpOnly; Secure; SameSite=Lax; Max-Age=0');
  res.status(200).send('Excel connected. You can close this tab — the app can now sync.');
}

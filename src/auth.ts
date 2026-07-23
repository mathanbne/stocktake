import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  type AccountInfo,
} from '@azure/msal-browser';
import { config } from './config';

/**
 * Sign-in happens on the device, as the scanner. There is no server and no
 * client secret — a public client proves itself with PKCE instead, which is why
 * the client id below is safe to ship in public JavaScript. Client ids are
 * identifiers, not credentials.
 */
export const SCOPES = ['Files.ReadWrite.All', 'User.Read'];

let msal: PublicClientApplication | null = null;

export class NotSignedInError extends Error {
  constructor() {
    super('Sign in to Microsoft to sync.');
    this.name = 'NotSignedInError';
  }
}

async function client(): Promise<PublicClientApplication> {
  if (!msal) {
    if (!config.msClientId) {
      throw new Error('VITE_MS_CLIENT_ID is not set — the app cannot sign in to Microsoft.');
    }
    msal = new PublicClientApplication({
      auth: {
        clientId: config.msClientId,
        // "organizations" accepts any work or school account without pinning a
        // tenant. Set VITE_MS_TENANT_ID to restrict it to yours.
        authority: `https://login.microsoftonline.com/${config.msTenantId || 'organizations'}`,
        redirectUri: window.location.origin,
      },
      cache: {
        // localStorage, not session: a scanner that reopens the app tomorrow
        // should still be signed in.
        cacheLocation: 'localStorage',
      },
    });
    await msal.initialize();
    // Completes a redirect sign-in when the app reloads back into itself.
    await msal.handleRedirectPromise();
  }
  return msal;
}

/** Whoever is signed in on this device, or null. */
export async function currentAccount(): Promise<AccountInfo | null> {
  const c = await client();
  return c.getAllAccounts()[0] ?? null;
}

/**
 * Redirect rather than popup: installed PWAs run without browser chrome, and
 * popup windows are unreliable — often blocked outright — in that context.
 */
export async function signIn(): Promise<void> {
  const c = await client();
  await c.loginRedirect({ scopes: SCOPES });
}

export async function signOut(): Promise<void> {
  const c = await client();
  const account = c.getAllAccounts()[0];
  if (account) await c.logoutRedirect({ account });
}

/**
 * Silent by default; falls back to an interactive redirect only when the
 * refresh token is genuinely spent. Offline this throws, which is correct —
 * every caller is a network operation anyway.
 */
export async function getToken(): Promise<string> {
  const c = await client();
  const account = c.getAllAccounts()[0];
  if (!account) throw new NotSignedInError();

  try {
    const r = await c.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      await c.acquireTokenRedirect({ scopes: SCOPES, account });
      throw new NotSignedInError(); // the redirect navigates away; nothing returns
    }
    throw e;
  }
}

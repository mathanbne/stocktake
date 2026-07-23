function num(v: string | undefined, fallback: number): number {
  const n = Number(v);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

export const config = {
  /**
   * Entra application (client) id. Public by design: a public client holds no
   * secret and proves itself with PKCE, so this identifies the app rather than
   * authorising it. Safe to ship in the bundle.
   */
  msClientId: (import.meta.env.VITE_MS_CLIENT_ID as string) ?? '',
  /** Optional. Blank accepts any work/school account; set to pin one tenant. */
  msTenantId: (import.meta.env.VITE_MS_TENANT_ID as string) ?? '',
  syncIntervalMs: num(import.meta.env.VITE_SYNC_INTERVAL_MS, 30_000),
  syncBatchSize: num(import.meta.env.VITE_SYNC_BATCH_SIZE, 50),
  /** Suppress re-detection of the same tag while it sits in frame. */
  scanCooldownMs: num(import.meta.env.VITE_SCAN_COOLDOWN_MS, 2_500),
};

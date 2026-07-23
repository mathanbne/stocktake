function num(v: string | undefined, fallback: number): number {
  const n = Number(v);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

export const config = {
  /**
   * Relative by design. The Vercel function at this path is the only thing that
   * holds Microsoft credentials — nothing secret is ever built into this bundle,
   * because anything prefixed VITE_ is inlined into public JavaScript.
   */
  apiUrl: '/api/excel',
  syncIntervalMs: num(import.meta.env.VITE_SYNC_INTERVAL_MS, 30_000),
  syncBatchSize: num(import.meta.env.VITE_SYNC_BATCH_SIZE, 50),
  /** Suppress re-detection of the same tag while it sits in frame. */
  scanCooldownMs: num(import.meta.env.VITE_SCAN_COOLDOWN_MS, 2_500),
};

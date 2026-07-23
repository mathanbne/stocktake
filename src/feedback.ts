import type { ScanResult } from './types';

let ctx: AudioContext | null = null;

/** Must be called from a user gesture once — iOS refuses audio otherwise. */
export function unlockAudio(): void {
  if (!ctx) ctx = new AudioContext();
  if (ctx.state === 'suspended') void ctx.resume();
}

function tone(freq: number, ms: number, when = 0): void {
  if (!ctx) return;
  const osc = ctx.createOscillator();
  const gain = ctx.createGain();
  osc.frequency.value = freq;
  osc.type = 'square';
  gain.gain.value = 0.15;
  osc.connect(gain).connect(ctx.destination);
  const t = ctx.currentTime + when;
  osc.start(t);
  osc.stop(t + ms / 1000);
}

export function feedbackFor(result: ScanResult): void {
  switch (result) {
    case 'sighted':
    case 'manual_entry':
      navigator.vibrate?.(60);
      tone(1200, 90);
      break;
    case 'sighted_wrong_location':
      navigator.vibrate?.([60, 60, 60]);
      tone(900, 90);
      tone(700, 90, 0.12);
      break;
    case 'duplicate':
      navigator.vibrate?.(30);
      tone(500, 60);
      break;
    case 'unknown_tag':
      navigator.vibrate?.([120, 60, 120]);
      tone(300, 220);
      break;
    case 'missing':
      break; // never produced while scanning
  }
}

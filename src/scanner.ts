import { config } from './config';

// 1D symbologies only — tags are linear barcodes, so excluding 2D formats
// cuts decode time and false positives.
const NATIVE_FORMATS = ['code_128', 'code_39', 'code_93', 'ean_13', 'ean_8', 'upc_a', 'upc_e', 'itf', 'codabar'];

export interface ScannerHandle {
  stop(): void;
}

type OnScan = (value: string) => void;

interface NativeDetector {
  detect(source: CanvasImageSource): Promise<{ rawValue: string }[]>;
}

declare global {
  interface Window {
    BarcodeDetector?: new (opts: { formats: string[] }) => NativeDetector;
  }
}

export async function nativeDetectorSupported(): Promise<boolean> {
  if (!('BarcodeDetector' in window)) return false;
  try {
    const supported = await (window.BarcodeDetector as unknown as {
      getSupportedFormats(): Promise<string[]>;
    }).getSupportedFormats();
    return NATIVE_FORMATS.some((f) => supported.includes(f));
  } catch {
    return false;
  }
}

/**
 * Continuous scan loop. The cooldown map suppresses re-firing while the same
 * barcode sits in frame; scanning the same tag again *after* the cooldown is a
 * legitimate event (classified upstream as a duplicate scan).
 */
export async function startScanner(video: HTMLVideoElement, onScan: OnScan): Promise<ScannerHandle> {
  const lastSeen = new Map<string, number>();

  const emit = (raw: string) => {
    const value = raw.trim();
    if (!value) return;
    const now = Date.now();
    const prev = lastSeen.get(value) ?? 0;
    if (now - prev < config.scanCooldownMs) return;
    lastSeen.set(value, now);
    onScan(value);
  };

  if (await nativeDetectorSupported()) {
    const stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: 'environment', width: { ideal: 1280 } },
    });
    video.srcObject = stream;
    await video.play();
    const detector = new window.BarcodeDetector!({ formats: NATIVE_FORMATS });
    let running = true;
    const loop = async () => {
      while (running) {
        if (video.readyState >= 2) {
          try {
            const codes = await detector.detect(video);
            for (const c of codes) emit(c.rawValue);
          } catch {
            // transient decode errors are expected between frames
          }
        }
        await new Promise((r) => setTimeout(r, 120));
      }
    };
    void loop();
    return {
      stop() {
        running = false;
        stream.getTracks().forEach((t) => t.stop());
      },
    };
  }

  // iOS Safari path — ZXing is ~400KB, so it's fetched (and thereafter served
  // from the service worker precache) only on browsers that need it.
  const [{ BrowserMultiFormatReader }, { BarcodeFormat, DecodeHintType }] = await Promise.all([
    import('@zxing/browser'),
    import('@zxing/library'),
  ]);
  const hints = new Map();
  hints.set(DecodeHintType.POSSIBLE_FORMATS, [
    BarcodeFormat.CODE_128, BarcodeFormat.CODE_39, BarcodeFormat.CODE_93,
    BarcodeFormat.EAN_13, BarcodeFormat.EAN_8, BarcodeFormat.UPC_A,
    BarcodeFormat.UPC_E, BarcodeFormat.ITF, BarcodeFormat.CODABAR,
  ]);
  const reader = new BrowserMultiFormatReader(hints);
  const controls = await reader.decodeFromVideoDevice(
    undefined,
    video,
    (result) => {
      if (result) emit(result.getText());
    },
  );
  return {
    stop() {
      controls.stop();
    },
  };
}

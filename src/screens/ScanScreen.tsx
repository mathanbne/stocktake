import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { RegisterIndex, captureScan, normalizeTag } from '../capture';
import { scansForSession } from '../db';
import { feedbackFor } from '../feedback';
import { startScanner, type ScannerHandle } from '../scanner';
import { refreshPending } from '../sync';
import { assetLabel, type ScanRecord, type ScanResult, type Session } from '../types';

const RESULT_LABEL: Record<ScanResult, string> = {
  sighted: 'SIGHTED',
  sighted_wrong_location: 'WRONG LOCATION',
  unknown_tag: 'UNKNOWN TAG',
  duplicate: 'ALREADY SCANNED',
  manual_entry: 'MANUAL ENTRY',
  missing: 'MISSING',
};

export default function ScanScreen(props: {
  session: Session;
  index: RegisterIndex;
  onCloseSession: () => void;
}) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const [scans, setScans] = useState<ScanRecord[]>([]);
  const [flash, setFlash] = useState<{ result: ScanResult; tag: string; detail: string } | null>(null);
  const [manualValue, setManualValue] = useState('');
  const [actualLocation, setActualLocation] = useState(props.session.scopeValue);
  const [cameraError, setCameraError] = useState<string | null>(null);

  // Rebuilt from the DB on mount so a resumed session's duplicate detection
  // and counters are exact, not remembered. Keyed on Asset Id, so the same
  // asset read by barcode and then by serial still counts once.
  const alreadyScanned = useMemo(() => {
    const set = new Set<string>();
    for (const s of scans) {
      if (s.result !== 'unknown_tag' && s.assetId !== '') set.add(normalizeTag(s.assetId));
    }
    return set;
  }, [scans]);

  useEffect(() => {
    void scansForSession(props.session.sessionId).then(setScans);
  }, [props.session.sessionId]);

  /**
   * The capture closure must see fresh state without being rebuilt, because
   * rebuilding it would restart the camera. A ref refreshed on every render is
   * what decouples the two.
   */
  const latest = useRef({ session: props.session, index: props.index, alreadyScanned, actualLocation });
  useEffect(() => {
    latest.current = { session: props.session, index: props.index, alreadyScanned, actualLocation };
  });

  const handleValue = useCallback(async (value: string, manual: boolean) => {
    const { session, index, alreadyScanned: seen, actualLocation: loc } = latest.current;
    const { scan, asset } = await captureScan({
      scannedValue: value,
      session,
      index,
      alreadyScanned: seen,
      actualLocation: loc,
      manual,
    });
    feedbackFor(scan.result);
    setFlash({
      result: scan.result,
      tag: asset ? asset.assetId : scan.scannedValue,
      detail:
        scan.result === 'sighted_wrong_location'
          ? scan.notes
          : asset
            ? assetLabel(asset)
            : 'Not in the register',
    });
    setScans((prev) => [...prev, scan]);
    void refreshPending();
  }, []);

  // Empty deps on purpose: the scanner mounts once for the life of the screen.
  // Any dependency here would tear down getUserMedia and re-initialise the
  // decoder on every scan — a visible stall on a phone.
  useEffect(() => {
    let handle: ScannerHandle | null = null;
    let cancelled = false;
    if (videoRef.current) {
      startScanner(videoRef.current, (v) => void handleValue(v, false))
        .then((h) => {
          if (cancelled) h.stop();
          else handle = h;
        })
        .catch(() => setCameraError('Camera unavailable. Use manual entry below.'));
    }
    return () => {
      cancelled = true;
      handle?.stop();
    };
  }, [handleValue]);

  const counters = useMemo(() => {
    const c = { total: 0, sighted: 0, wrong: 0, unknown: 0, dup: 0 };
    for (const s of scans) {
      if (s.result === 'missing') continue;
      c.total++;
      if (s.result === 'sighted' || s.result === 'manual_entry') c.sighted++;
      if (s.result === 'sighted_wrong_location') c.wrong++;
      if (s.result === 'unknown_tag') c.unknown++;
      if (s.result === 'duplicate') c.dup++;
    }
    return c;
  }, [scans]);

  useEffect(() => {
    if (!flash) return;
    const t = setTimeout(() => setFlash(null), 1800);
    return () => clearTimeout(t);
  }, [flash]);

  const submitManual = () => {
    if (!manualValue.trim()) return;
    void handleValue(manualValue, true);
    setManualValue('');
  };

  return (
    <main className="screen scan-screen">
      <div className="camera-wrap">
        <video ref={videoRef} playsInline muted />
        {cameraError && <div className="camera-error">{cameraError}</div>}
        {flash && (
          <div className={`flash flash-${flash.result}`}>
            <div className="flash-verdict">{RESULT_LABEL[flash.result]}</div>
            <div className="flash-tag">{flash.tag}</div>
            <div className="flash-detail">{flash.detail}</div>
          </div>
        )}
      </div>

      <div className="counters">
        <div><b>{counters.total}</b><span>scans</span></div>
        <div><b>{counters.sighted}</b><span>sighted</span></div>
        <div><b>{counters.wrong}</b><span>wrong loc</span></div>
        <div><b>{counters.unknown}</b><span>unknown</span></div>
        <div><b>{counters.dup}</b><span>dupes</span></div>
      </div>

      <label className="loc-field">
        Current location
        <input value={actualLocation} onChange={(e) => setActualLocation(e.target.value)} />
      </label>

      <div className="manual-row">
        <input
          value={manualValue}
          inputMode="text"
          autoCapitalize="characters"
          placeholder="Type a tag manually"
          onChange={(e) => setManualValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') submitManual();
          }}
        />
        <button className="btn" disabled={!manualValue.trim()} onClick={submitManual}>
          Add
        </button>
      </div>

      <ul className="recent">
        {scans.slice(-6).reverse().map((s) => (
          <li key={s.scanId} className={`recent-${s.result}`}>
            <code>{s.assetId || s.scannedValue}</code> {RESULT_LABEL[s.result]}
          </li>
        ))}
      </ul>

      <button className="btn btn-danger" onClick={props.onCloseSession}>
        Finish &amp; reconcile
      </button>
    </main>
  );
}

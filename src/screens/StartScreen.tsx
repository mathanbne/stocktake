import { useEffect, useState } from 'react';
import { AccessDeniedError, ReauthRequiredError, fetchRegister } from '../api';
import {
  getAccessCode,
  loadRegister,
  putSession,
  registerFetchedAt,
  replaceRegister,
  scansForSession,
  setAccessCode,
} from '../db';
import { unlockAudio } from '../feedback';
import type { ScopeType, Session } from '../types';

export default function StartScreen(props: {
  resumable: Session | null;
  onEnter: (s: Session) => void;
}) {
  const [name, setName] = useState('');
  const [scopeType, setScopeType] = useState<ScopeType>('room');
  const [scopeValue, setScopeValue] = useState('');
  const [code, setCode] = useState('');
  const [cachedAt, setCachedAt] = useState<string | null>(null);
  const [assetCount, setAssetCount] = useState<number | null>(null);
  const [fetching, setFetching] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [resumeCounts, setResumeCounts] = useState<number | null>(null);

  useEffect(() => {
    void registerFetchedAt().then(setCachedAt);
    void getAccessCode().then(setCode);
    void loadRegister().then((a) => setAssetCount(a.length));
    if (props.resumable) {
      void scansForSession(props.resumable.sessionId).then((s) => setResumeCounts(s.length));
    }
  }, [props.resumable]);

  const refreshRegister = async () => {
    setFetching(true);
    setError(null);
    try {
      await setAccessCode(code.trim());
      const assets = await fetchRegister();
      await replaceRegister(assets);
      setCachedAt(new Date().toISOString());
      setAssetCount(assets.length);
      if (assets.length === 0) {
        setError('The register came back empty. Check the AssetRegister table in the workbook.');
      }
    } catch (e) {
      if (e instanceof AccessDeniedError) setError(e.message);
      else if (e instanceof ReauthRequiredError) setError(e.message);
      else setError(e instanceof Error ? e.message : 'Could not fetch the register');
    } finally {
      setFetching(false);
    }
  };

  const begin = async () => {
    if (!name.trim() || !scopeValue.trim()) return;
    if (!cachedAt || !assetCount) {
      setError('Download the register once before starting — scanning runs against the local copy.');
      return;
    }
    unlockAudio();
    const session: Session = {
      sessionId: crypto.randomUUID(),
      scopeType,
      scopeValue: scopeValue.trim(),
      scannedBy: name.trim(),
      startedAt: new Date().toISOString(),
      status: 'open',
    };
    await putSession(session);
    props.onEnter(session);
  };

  return (
    <main className="screen">
      {props.resumable && (
        <section className="resume card">
          <h2>Open session found</h2>
          <p>
            <b>{props.resumable.scopeType}</b>: {props.resumable.scopeValue}
            <br />
            Started {new Date(props.resumable.startedAt).toLocaleString()} by {props.resumable.scannedBy}
            {resumeCounts !== null && <><br />{resumeCounts} scans recorded</>}
            {props.resumable.status === 'closing' && <><br />Close pending — reopen to finish it</>}
          </p>
          <button
            className="btn"
            onClick={() => {
              unlockAudio();
              props.onEnter(props.resumable!);
            }}
          >
            Resume session
          </button>
        </section>
      )}

      <section className="card">
        <h2>New session</h2>
        <label>
          Your name
          <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Scanned by" />
        </label>
        <label>
          Scope
          <div className="scope-row">
            <select value={scopeType} onChange={(e) => setScopeType(e.target.value as ScopeType)}>
              <option value="site">Site</option>
              <option value="floor">Floor</option>
              <option value="room">Room</option>
            </select>
            <input
              value={scopeValue}
              onChange={(e) => setScopeValue(e.target.value)}
              placeholder="e.g. A BLOCK or 01G04"
            />
          </div>
          <small>
            Matches Facility Group / Facility. Everything in scope and unscanned will be flagged missing at
            close.
          </small>
        </label>

        <label>
          Device access code
          <input
            value={code}
            onChange={(e) => setCode(e.target.value)}
            type="password"
            autoComplete="off"
            placeholder="Needed only to sync"
          />
          <small>Stored on this device. Scanning works offline without it.</small>
        </label>

        <div className="register-state">
          {cachedAt ? (
            <small>
              {assetCount ?? 0} assets · cached {new Date(cachedAt).toLocaleString()}
            </small>
          ) : (
            <small>Register not downloaded yet</small>
          )}
          <button className="btn btn-quiet" onClick={() => void refreshRegister()} disabled={fetching}>
            {fetching ? 'Downloading…' : cachedAt ? 'Refresh register' : 'Download register'}
          </button>
        </div>

        {error && <p className="error">{error}</p>}

        <button className="btn btn-big" onClick={() => void begin()} disabled={!name.trim() || !scopeValue.trim()}>
          Start scanning
        </button>
      </section>
    </main>
  );
}

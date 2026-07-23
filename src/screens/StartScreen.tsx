import { useEffect, useState } from 'react';
import { fetchRegister } from '../api';
import { currentAccount, signIn, signOut } from '../auth';
import {
  getWorkbook,
  loadRegister,
  putSession,
  registerFetchedAt,
  replaceRegister,
  scansForSession,
  setWorkbook,
} from '../db';
import { findWorkbooks, validateWorkbook, type WorkbookChoice, type WorkbookRef } from '../graph';
import { unlockAudio } from '../feedback';
import type { ScopeType, Session } from '../types';

function message(e: unknown): string {
  return e instanceof Error ? e.message : 'Something went wrong';
}

export default function StartScreen(props: {
  resumable: Session | null;
  onEnter: (s: Session) => void;
}) {
  const [account, setAccount] = useState<{ name: string; username: string } | null>(null);
  const [workbook, setWorkbookState] = useState<WorkbookRef | null>(null);
  const [picking, setPicking] = useState(false);
  const [term, setTerm] = useState('');
  const [results, setResults] = useState<WorkbookChoice[] | null>(null);
  const [busy, setBusy] = useState<string | null>(null);

  const [name, setName] = useState('');
  const [scopeType, setScopeType] = useState<ScopeType>('room');
  const [scopeValue, setScopeValue] = useState('');
  const [cachedAt, setCachedAt] = useState<string | null>(null);
  const [assetCount, setAssetCount] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [resumeCounts, setResumeCounts] = useState<number | null>(null);

  useEffect(() => {
    void (async () => {
      const acc = await currentAccount();
      if (acc) {
        setAccount({ name: acc.name ?? acc.username, username: acc.username });
        // Pre-fill the scanner's name from the signed-in identity; the ScanLog
        // should say who was holding the phone, and they already told us.
        setName((n) => n || (acc.name ?? acc.username));
      }
      setWorkbookState(await getWorkbook());
      setCachedAt(await registerFetchedAt());
      setAssetCount((await loadRegister()).length);
    })();
    if (props.resumable) {
      void scansForSession(props.resumable.sessionId).then((s) => setResumeCounts(s.length));
    }
  }, [props.resumable]);

  const search = async () => {
    setBusy('Searching…');
    setError(null);
    try {
      setResults(await findWorkbooks(term));
    } catch (e) {
      setError(message(e));
    } finally {
      setBusy(null);
    }
  };

  const choose = async (c: WorkbookChoice) => {
    setBusy('Checking workbook…');
    setError(null);
    try {
      const ref: WorkbookRef = { driveId: c.driveId, itemId: c.itemId, name: c.name };
      await validateWorkbook(ref); // fail here, not mid-stocktake
      await setWorkbook(ref);
      setWorkbookState(ref);
      setPicking(false);
      setResults(null);
    } catch (e) {
      setError(message(e));
    } finally {
      setBusy(null);
    }
  };

  const refreshRegister = async () => {
    setBusy('Downloading…');
    setError(null);
    try {
      const assets = await fetchRegister();
      await replaceRegister(assets);
      setCachedAt(new Date().toISOString());
      setAssetCount(assets.length);
      if (assets.length === 0) {
        setError('The register came back empty — check the AssetRegister table has rows.');
      }
    } catch (e) {
      setError(message(e));
    } finally {
      setBusy(null);
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

  // Signed out is a hard stop for setup, but never for a session already in
  // flight — a resumable session stays reachable so captured scans aren't
  // stranded behind a login the scanner may not be able to complete right now.
  if (!account) {
    return (
      <main className="screen">
        {props.resumable && (
          <section className="resume card">
            <h2>Open session found</h2>
            <p>{resumeCounts ?? 0} scans recorded. Sign in when you can to sync them.</p>
            <button className="btn" onClick={() => { unlockAudio(); props.onEnter(props.resumable!); }}>
              Resume session
            </button>
          </section>
        )}
        <section className="card">
          <h2>Sign in</h2>
          <p>
            Sign in with your work account to reach the stocktake workbook. Scans are stored on this
            device and uploaded as you go.
          </p>
          <button className="btn btn-big" onClick={() => void signIn()}>
            Sign in with Microsoft
          </button>
          {error && <p className="error">{error}</p>}
        </section>
      </main>
    );
  }

  return (
    <main className="screen">
      <section className="card account-card">
        <div>
          <small>Signed in as</small>
          <div className="account-name">{account.name}</div>
        </div>
        <button className="btn btn-quiet" onClick={() => void signOut()}>Sign out</button>
      </section>

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
          <button className="btn" onClick={() => { unlockAudio(); props.onEnter(props.resumable!); }}>
            Resume session
          </button>
        </section>
      )}

      <section className="card">
        <h2>Workbook</h2>
        {workbook && !picking ? (
          <div className="register-state">
            <small>
              <code>{workbook.name}</code>
              <br />
              {cachedAt ? `${assetCount ?? 0} assets · cached ${new Date(cachedAt).toLocaleString()}` : 'Register not downloaded yet'}
            </small>
            <button className="btn btn-quiet" onClick={() => setPicking(true)}>Change</button>
          </div>
        ) : (
          <>
            <label>
              Find your workbook
              <div className="scope-row scope-row-wide">
                <input
                  value={term}
                  placeholder="stock.xlsx"
                  onChange={(e) => setTerm(e.target.value)}
                  onKeyDown={(e) => { if (e.key === 'Enter') void search(); }}
                />
                <button className="btn" onClick={() => void search()} disabled={busy !== null}>
                  Search
                </button>
              </div>
              <small>Searches your OneDrive and any SharePoint site you can open.</small>
            </label>

            {results && (
              <ul className="picker">
                {results.length === 0 && <li className="picker-empty">No .xlsx files matched.</li>}
                {results.map((r) => (
                  <li key={`${r.driveId}:${r.itemId}`}>
                    <button onClick={() => void choose(r)} disabled={busy !== null}>
                      <b>{r.name}</b>
                      <small>{r.location}</small>
                    </button>
                  </li>
                ))}
              </ul>
            )}
            {workbook && (
              <button className="btn btn-quiet" onClick={() => { setPicking(false); setResults(null); }}>
                Cancel
              </button>
            )}
          </>
        )}

        {workbook && !picking && (
          <button className="btn btn-quiet" onClick={() => void refreshRegister()} disabled={busy !== null}>
            {cachedAt ? 'Refresh register' : 'Download register'}
          </button>
        )}
      </section>

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
            Matches Facility Group / Facility. Everything in scope and unscanned will be flagged missing
            at close.
          </small>
        </label>

        <button
          className="btn btn-big"
          onClick={() => void begin()}
          disabled={!name.trim() || !scopeValue.trim()}
        >
          Start scanning
        </button>
      </section>

      {busy && <p className="busy">{busy}</p>}
      {error && <p className="error">{error}</p>}
    </main>
  );
}

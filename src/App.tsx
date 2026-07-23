import { useEffect, useState } from 'react';
import { RegisterIndex } from './capture';
import { StorageFullError, findResumableSession, loadRegister } from './db';
import { subscribeSync, type SyncStatus } from './sync';
import type { Session } from './types';
import StartScreen from './screens/StartScreen';
import ScanScreen from './screens/ScanScreen';
import CloseScreen from './screens/CloseScreen';

type Screen =
  | { name: 'boot' }
  | { name: 'start'; resumable: Session | null }
  | { name: 'scan'; session: Session; index: RegisterIndex }
  | { name: 'close'; session: Session; index: RegisterIndex }
  | { name: 'closed'; session: Session };

export default function App() {
  const [screen, setScreen] = useState<Screen>({ name: 'boot' });
  const [sync, setSync] = useState<SyncStatus>({ pending: 0, lastError: null, flushing: false });
  const [fatal, setFatal] = useState<string | null>(null);

  useEffect(() => subscribeSync(setSync), []);

  useEffect(() => {
    void (async () => {
      const resumable = await findResumableSession();
      setScreen({ name: 'start', resumable });
    })();
  }, []);

  // Storage exhaustion anywhere becomes a blocking, unmissable stop.
  useEffect(() => {
    const handler = (ev: PromiseRejectionEvent) => {
      if (ev.reason instanceof StorageFullError) {
        setFatal(ev.reason.message);
        ev.preventDefault();
      }
    };
    window.addEventListener('unhandledrejection', handler);
    return () => window.removeEventListener('unhandledrejection', handler);
  }, []);

  const enterSession = async (session: Session) => {
    const index = new RegisterIndex(await loadRegister());
    if (session.status === 'closing') {
      setScreen({ name: 'close', session, index });
    } else {
      setScreen({ name: 'scan', session, index });
    }
  };

  if (fatal) {
    return (
      <div className="fatal">
        <h1>Storage full</h1>
        <p>{fatal}</p>
        <p>Free up space on this device, then reopen the app. Nothing already saved has been lost.</p>
      </div>
    );
  }

  return (
    <div className="app">
      <header className="topbar">
        <span className="brand">STOCKTAKE</span>
        <span className={`sync ${sync.lastError ? 'sync-err' : sync.pending > 0 ? 'sync-wait' : 'sync-ok'}`}>
          {sync.flushing
            ? 'Syncing…'
            : sync.pending > 0
              ? `${sync.pending} queued`
              : navigator.onLine
                ? 'Synced'
                : 'Offline'}
        </span>
      </header>

      {screen.name === 'boot' && <div className="center">Loading…</div>}

      {screen.name === 'start' && (
        <StartScreen resumable={screen.resumable} onEnter={(s) => void enterSession(s)} />
      )}

      {screen.name === 'scan' && (
        <ScanScreen
          session={screen.session}
          index={screen.index}
          onCloseSession={() => setScreen({ name: 'close', session: screen.session, index: screen.index })}
        />
      )}

      {screen.name === 'close' && (
        <CloseScreen
          session={screen.session}
          index={screen.index}
          onBack={() => setScreen({ name: 'scan', session: screen.session, index: screen.index })}
          onClosed={(s) => setScreen({ name: 'closed', session: s })}
        />
      )}

      {screen.name === 'closed' && (
        <div className="center done">
          <h1>Session closed</h1>
          <p>All results are in the workbook.</p>
          <button className="btn" onClick={() => setScreen({ name: 'start', resumable: null })}>
            Start a new session
          </button>
        </div>
      )}
    </div>
  );
}

import { useEffect, useState } from 'react';
import { RegisterIndex } from '../capture';
import { getPendingClose } from '../db';
import { confirmClose, previewMissing } from '../reconcile';
import { requestClose, subscribeSync, type SyncStatus } from '../sync';
import { assetLabel, type RegisterAsset, type Session } from '../types';

export default function CloseScreen(props: {
  session: Session;
  index: RegisterIndex;
  onBack: () => void;
  onClosed: (s: Session) => void;
}) {
  const [missing, setMissing] = useState<RegisterAsset[] | null>(null);
  const [confirmed, setConfirmed] = useState(props.session.status === 'closing');
  const [sync, setSync] = useState<SyncStatus>({ pending: 0, lastError: null, flushing: false });

  useEffect(() => subscribeSync(setSync), []);

  useEffect(() => {
    void (async () => {
      // A session already in 'closing' resumes its saved payload — the preview
      // shown is exactly what was previously confirmed, never recomputed.
      if (props.session.status === 'closing') {
        const pc = await getPendingClose(props.session.sessionId);
        if (pc) {
          setMissing(
            pc.missingAssetIds
              .map((id) => props.index.byAssetId(id))
              .filter((a): a is RegisterAsset => !!a),
          );
          requestClose(props.session, props.onClosed);
          return;
        }
      }
      setMissing(await previewMissing(props.session, props.index));
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const confirm = async () => {
    if (!missing) return;
    const { session } = await confirmClose(props.session, missing);
    setConfirmed(true);
    requestClose(session, props.onClosed);
  };

  if (missing === null) return <div className="center">Preparing reconciliation…</div>;

  return (
    <main className="screen">
      <section className="card">
        <h2>{confirmed ? 'Closing session' : 'Close session?'}</h2>
        <p>
          {missing.length === 0
            ? 'Every in-scope asset was accounted for. Nothing will be marked missing.'
            : `${missing.length} asset${missing.length === 1 ? '' : 's'} in scope ${
                missing.length === 1 ? 'was' : 'were'
              } never scanned and will be marked MISSING:`}
        </p>

        {missing.length > 0 && (
          <ul className="missing-list">
            {missing.map((a) => (
              <li key={a.assetId}>
                <code>{a.assetId}</code>
                <span>{assetLabel(a)}</span>
                <small>{a.expectedLocation}</small>
              </li>
            ))}
          </ul>
        )}

        {!confirmed ? (
          <>
            <button className="btn btn-danger btn-big" onClick={() => void confirm()}>
              Confirm — mark {missing.length} missing and close
            </button>
            <button className="btn btn-quiet" onClick={props.onBack}>
              Keep scanning
            </button>
          </>
        ) : (
          <div className="closing-state">
            {sync.lastError ? (
              <p className="error">
                Not delivered yet: {sync.lastError}. Saved locally — it will retry automatically, and this
                screen is safe to leave and come back to.
              </p>
            ) : (
              <p>
                {sync.pending > 0
                  ? `Uploading ${sync.pending} remaining scans…`
                  : 'Writing results to the workbook…'}
              </p>
            )}
          </div>
        )}
      </section>
    </main>
  );
}

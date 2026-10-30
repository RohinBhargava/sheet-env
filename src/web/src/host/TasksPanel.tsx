// Optional side panel (?panel=1): the scenario's tasks and a Reset button.
import { useEffect, useState } from 'react';
import { api, type Task } from '../api';

export function TasksPanel({ onOpen, onReset, ask }: { onOpen: (path: string) => void; onReset: () => void; ask: (text: string, buttons: string[], title?: string) => Promise<string> }) {
  const [tasks, setTasks] = useState<Task[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  useEffect(() => { api.tasks().then(setTasks, e => setErr(String(e?.message ?? e))); }, []);
  const reset = async () => {
    if (await ask('Reset every workbook to the scenario\'s starting files? Changes you made are discarded.', ['Reset', 'Cancel'], 'Reset scenario') !== 'Reset') return;
    setBusy(true);
    try { await api.reset('fund-admin'); onReset(); }
    catch (e) { setErr(`Reset failed: ${e instanceof Error ? e.message : String(e)}`); }
    finally { setBusy(false); }
  };
  return (
    <aside className="tasks" aria-label="Tasks">
      <div className="tasks-head">
        <h2>Tasks</h2>
        <button className="home-btn" onClick={() => void reset()} disabled={busy}>{busy ? 'Resetting…' : 'Reset'}</button>
      </div>
      {err && <div className="tasks-err">{err}</div>}
      {!err && !tasks && <div className="tasks-empty">Loading…</div>}
      {tasks && !tasks.length && <div className="tasks-empty">No tasks.</div>}
      <ol className="tasks-list">
        {tasks?.map(t => (
          <li key={t.id} className="task">
            <div className="task-title">{t.title}</div>
            <div className="task-text">{t.instructions}</div>
            <button className="task-file" onClick={() => onOpen(t.file)} title={t.file}>Open {t.file.split('/').pop()}</button>
          </li>
        ))}
      </ol>
    </aside>
  );
}

import React, { useState } from 'react';
import { invoke } from '@forge/bridge';
import { SectionCard, Button, Skeleton, useToast, formatError } from '../common';

// Status / Comment / Log Time / Assign / Edit Fields — self-contained like TeamsTab.
// `onIssueChange` lets successful actions (transition, assign, edit) update the summary card
// that lives in the parent, without lifting all of this tab's form state up to get there.
export default function ActionsTab({ issueKey, issue, onIssueChange }) {
  const showToast = useToast();
  const [mode,        setMode]        = useState(null);
  const [submitting,  setSubmitting]  = useState(false);
  const [err,         setErr]         = useState('');

  const [transitions,  setTransitions]  = useState([]);
  const [commentText,  setCommentText]  = useState('');
  const [timeSpent,    setTimeSpent]    = useState('');
  const [workDesc,     setWorkDesc]     = useState('');
  const [assignEmail,  setAssignEmail]  = useState('');
  const [priorities,   setPriorities]   = useState([]);
  const [editPriority, setEditPriority] = useState('');
  const [editLabels,   setEditLabels]   = useState((issue?.labels || []).join(', '));
  const [editDueDate,  setEditDueDate]  = useState(issue?.dueDate || '');

  function reset() { setMode(null); setErr(''); }

  function openMode(m) {
    setMode(m); setErr('');
    if (m === 'status') invoke('getIssueTransitions', { issueKey }).then(res => setTransitions(res.transitions || []));
    if (m === 'editfields') {
      // Re-sync from the current issue each time this opens, not just on first mount — the
      // issue may have changed (e.g. a transition applied) since the last time this was open.
      setEditLabels((issue?.labels || []).join(', '));
      setEditDueDate(issue?.dueDate || '');
      setEditPriority('');
      invoke('getPriorities').then(res => setPriorities(res.priorities || []));
    }
  }

  async function applyTransition(id, name) {
    setErr(''); setSubmitting(true);
    try {
      await invoke('transitionIssue', { issueKey, transitionId: id });
      showToast(`Status updated to "${name}".`);
      onIssueChange(prev => ({ ...prev, status: name }));
      setMode(null);
    } catch (e) { setErr(formatError(e, 'Transition failed.')); }
    finally { setSubmitting(false); }
  }

  async function submit() {
    setErr(''); setSubmitting(true);
    try {
      if (mode === 'comment') {
        if (!commentText.trim()) { setErr('Enter a comment.'); return; }
        await invoke('addIssueComment', { issueKey, commentText: commentText.trim() });
        showToast('Comment added.'); setCommentText(''); setMode(null);
      } else if (mode === 'logtime') {
        if (!timeSpent.trim()) { setErr('Enter time (e.g. 2h, 30m).'); return; }
        await invoke('logIssueWork', { issueKey, timeSpent: timeSpent.trim(), description: workDesc });
        showToast(`Logged ${timeSpent} on ${issueKey}.`); setTimeSpent(''); setWorkDesc(''); setMode(null);
      } else if (mode === 'assign') {
        if (!assignEmail.trim()) { setErr('Enter an email address.'); return; }
        const res = await invoke('assignIssue', { issueKey, email: assignEmail.trim() });
        showToast(`Assigned to ${res.displayName}.`);
        onIssueChange(prev => ({ ...prev, assignee: { name: res.displayName, email: assignEmail.trim() } }));
        setMode(null);
      } else if (mode === 'editfields') {
        const payload = { issueKey, labels: editLabels.split(',').map(s => s.trim()).filter(Boolean) };
        if (editPriority) payload.priority = { id: editPriority };
        if (editDueDate) payload.dueDate = editDueDate;
        await invoke('updateIssueFields', payload);
        const priorityName = priorities.find(p => p.id === editPriority)?.name;
        showToast('Issue fields updated.');
        onIssueChange(prev => ({ ...prev, priority: priorityName || prev.priority, labels: payload.labels, dueDate: editDueDate || prev.dueDate }));
        setMode(null);
      }
    } catch (e) { setErr(formatError(e, 'Failed.')); }
    finally { setSubmitting(false); }
  }

  return (
    <div className="space-y-2">
      {err && <div className="alert-err">{err}</div>}

      {!mode && (
        <>
          <Button className="w-full" onClick={() => openMode('status')}>Update Status</Button>
          <Button variant="success" className="w-full" onClick={() => openMode('comment')}>Add Comment</Button>
          <Button variant="purple" className="w-full" onClick={() => openMode('logtime')}>Log Time</Button>
          <Button variant="dark" className="w-full" onClick={() => openMode('assign')}>Assign Issue</Button>
          <Button variant="purple" className="w-full" onClick={() => openMode('editfields')}>Edit Fields</Button>
        </>
      )}

      {mode === 'status' && (
        <SectionCard title="Select new status">
          <div className="space-y-1.5">
            {transitions.length === 0 && (
              <>
                <Skeleton className="h-8 w-full" />
                <Skeleton className="h-8 w-full" />
              </>
            )}
            {transitions.map(t => (
              <Button key={t.id} variant="ghost" className="w-full text-left" onClick={() => applyTransition(t.id, t.name)} loading={submitting}>
                {t.name}
              </Button>
            ))}
          </div>
          <Button variant="ghost" className="w-full mt-1.5" onClick={reset}>Cancel</Button>
        </SectionCard>
      )}

      {mode === 'comment' && (
        <SectionCard title="Add Comment">
          <textarea className="form-textarea w-full h-20" value={commentText} onChange={e => setCommentText(e.target.value)} placeholder="Enter your comment…" />
          <div className="flex gap-2 mt-3">
            <Button variant="success" className="flex-1" onClick={submit} loading={submitting}>Add Comment</Button>
            <Button variant="ghost" onClick={reset}>Cancel</Button>
          </div>
        </SectionCard>
      )}

      {mode === 'logtime' && (
        <SectionCard title="Log Time">
          <label className="label mt-0">Time Spent</label>
          <input className="form-input" value={timeSpent} onChange={e => setTimeSpent(e.target.value)} placeholder="e.g. 2h, 30m, 1h 30m" />
          <label className="label">Description (optional)</label>
          <textarea className="form-textarea w-full h-14" value={workDesc} onChange={e => setWorkDesc(e.target.value)} placeholder="What did you work on?" />
          <div className="flex gap-2 mt-3">
            <Button variant="purple" className="flex-1" onClick={submit} loading={submitting}>Log Time</Button>
            <Button variant="ghost" onClick={reset}>Cancel</Button>
          </div>
        </SectionCard>
      )}

      {mode === 'assign' && (
        <SectionCard title="Assign Issue">
          <label className="label mt-0">Assign to (email)</label>
          <input className="form-input" value={assignEmail} onChange={e => setAssignEmail(e.target.value)} placeholder="user@company.com" />
          <div className="flex gap-2 mt-3">
            <Button variant="dark" className="flex-1" onClick={submit} loading={submitting}>Assign</Button>
            <Button variant="ghost" onClick={reset}>Cancel</Button>
          </div>
        </SectionCard>
      )}

      {mode === 'editfields' && (
        <SectionCard title="Edit Fields">
          <label className="label mt-0">Priority</label>
          <select className="form-select" value={editPriority} onChange={e => setEditPriority(e.target.value)}>
            <option value="">— no change —</option>
            {priorities.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          <label className="label">Labels (comma-separated)</label>
          <input className="form-input" value={editLabels} onChange={e => setEditLabels(e.target.value)} placeholder="backend, urgent" />
          <label className="label">Due Date</label>
          <input type="date" className="form-input" value={editDueDate} onChange={e => setEditDueDate(e.target.value)} />
          <div className="flex gap-2 mt-3">
            <Button variant="purple" className="flex-1" onClick={submit} loading={submitting}>Save Changes</Button>
            <Button variant="ghost" onClick={reset}>Cancel</Button>
          </div>
        </SectionCard>
      )}
    </div>
  );
}

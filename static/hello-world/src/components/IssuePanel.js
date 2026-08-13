import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';

// ── Issue Context Panel ───────────────────────────────────────────────────────
export default function IssuePanel({ issueKey }) {
  const [issue,       setIssue]       = useState(null);
  const [currentUser, setCurrentUser] = useState(null);
  const [tab,         setTab]         = useState('teams');
  const [mode,        setMode]        = useState(null);
  const [toEmail,     setToEmail]     = useState('');
  const [groupEmails, setGroupEmails] = useState('');
  const [transitions, setTransitions] = useState([]);
  const [commentText, setCommentText] = useState('');
  const [timeSpent,   setTimeSpent]   = useState('');
  const [workDesc,    setWorkDesc]    = useState('');
  const [assignEmail, setAssignEmail] = useState('');
  const [activity,    setActivity]    = useState(null);
  const [priorities,  setPriorities]  = useState([]);
  const [editPriority, setEditPriority] = useState('');
  const [editLabels,   setEditLabels]   = useState('');
  const [editDueDate,  setEditDueDate]  = useState('');
  const [loading,     setLoading]     = useState(true);
  const [submitting,  setSubmitting]  = useState(false);
  const [err,         setErr]         = useState('');
  const [msg,         setMsg]         = useState('');

  useEffect(() => {
    if (!issueKey) return;
    Promise.all([
      invoke('getIssueDetails', { issueKey }),
      invoke('getCurrentUser'),
    ]).then(([issueData, userData]) => {
      setIssue(issueData);
      setCurrentUser(userData);
      setToEmail(issueData.assignee?.email || '');
      setGroupEmails([issueData.assignee?.email, issueData.reporter?.email].filter(Boolean).join(', '));
      setLoading(false);
    }).catch(e => { setErr(e.message || 'Failed to load issue.'); setLoading(false); });
  }, [issueKey]);

  function reset() { setMode(null); setErr(''); setMsg(''); }

  async function loadTransitions() {
    const res = await invoke('getIssueTransitions', { issueKey });
    setTransitions(res.transitions || []);
  }

  async function loadActivity() {
    setActivity(null);
    const res = await invoke('getIssueActivity', { issueKey });
    setActivity(res.comments || []);
  }

  async function loadEditFields() {
    const res = await invoke('getPriorities');
    setPriorities(res.priorities || []);
    setEditLabels((issue?.labels || []).join(', '));
    setEditDueDate(issue?.dueDate || '');
    setEditPriority('');
  }

  function switchTab(t) {
    setTab(t); reset();
    if (t === 'activity') loadActivity();
  }

  async function submit() {
    setErr(''); setMsg(''); setSubmitting(true);
    try {
      if (mode === 'dm') {
        if (!toEmail.trim()) { setErr('Enter a recipient email.'); return; }
        await invoke('startDM', { fromEmail: currentUser.email, toEmail: toEmail.trim(), issueKey, issueSummary: issue.summary });
        setMsg(`DM sent to ${toEmail.trim()}`); setMode(null);
      } else if (mode === 'group') {
        const emails = groupEmails.split(',').map(e => e.trim()).filter(Boolean);
        if (emails.length < 2) { setErr('Enter at least 2 emails.'); return; }
        await invoke('startGroupChat', { emails, issueKey, issueSummary: issue.summary });
        setMsg(`Group chat created with ${emails.length} members.`); setMode(null);
      } else if (mode === 'channel') {
        await invoke('postToChannelManual', { issueKey, issueSummary: issue.summary });
        setMsg('Posted to Teams channel.'); setMode(null);
      } else if (mode === 'comment') {
        if (!commentText.trim()) { setErr('Enter a comment.'); return; }
        await invoke('addIssueComment', { issueKey, commentText: commentText.trim() });
        setMsg('Comment added.'); setCommentText(''); setMode(null);
      } else if (mode === 'logtime') {
        if (!timeSpent.trim()) { setErr('Enter time (e.g. 2h, 30m).'); return; }
        await invoke('logIssueWork', { issueKey, timeSpent: timeSpent.trim(), description: workDesc });
        setMsg(`Logged ${timeSpent} on ${issueKey}.`); setTimeSpent(''); setWorkDesc(''); setMode(null);
      } else if (mode === 'assign') {
        if (!assignEmail.trim()) { setErr('Enter an email address.'); return; }
        const res = await invoke('assignIssue', { issueKey, email: assignEmail.trim() });
        setMsg(`Assigned to ${res.displayName}.`);
        setIssue(prev => ({ ...prev, assignee: { name: res.displayName, email: assignEmail.trim() } }));
        setMode(null);
      } else if (mode === 'editfields') {
        const payload = { issueKey, labels: editLabels.split(',').map(s => s.trim()).filter(Boolean) };
        if (editPriority) payload.priority = { id: editPriority };
        if (editDueDate) payload.dueDate = editDueDate;
        await invoke('updateIssueFields', payload);
        const priorityName = priorities.find(p => p.id === editPriority)?.name;
        setMsg('Issue fields updated.');
        setIssue(prev => ({ ...prev, priority: priorityName || prev.priority, labels: payload.labels, dueDate: editDueDate || prev.dueDate }));
        setMode(null);
      }
    } catch (e) { setErr(e.message || 'Failed.'); }
    finally { setSubmitting(false); }
  }

  async function applyTransition(id, name) {
    setErr(''); setMsg(''); setSubmitting(true);
    try {
      await invoke('transitionIssue', { issueKey, transitionId: id });
      setMsg(`Status updated to "${name}".`);
      setIssue(prev => ({ ...prev, status: name }));
      setMode(null);
    } catch (e) { setErr(e.message || 'Transition failed.'); }
    finally { setSubmitting(false); }
  }

  const statusBadgeClass = status => {
    const map = { 'In Progress': 'bg-jira-blue text-white', 'Done': 'bg-jira-green text-white', 'In Review': 'bg-jira-purple text-white' };
    return `text-xs font-semibold px-2 py-0.5 rounded ${map[status] || 'bg-jira-border text-jira-dark'}`;
  };

  const tabClass = active => `tab-btn ${active ? 'border-jira-blue text-jira-blue' : 'border-transparent text-jira-grey hover:text-jira-dark'}`;

  if (loading) return <div className="p-4 text-jira-grey text-sm">Loading…</div>;

  return (
    <div className="p-3 font-sans text-sm">

      {/* Issue summary */}
      {issue && (
        <div className="card mb-3">
          <div className="font-bold text-jira-dark text-sm">{issue.key}</div>
          <div className="text-jira-dark mt-1 mb-2 leading-snug text-xs">{issue.summary}</div>
          <div className="flex flex-wrap gap-2 items-center">
            <span className={statusBadgeClass(issue.status)}>{issue.status}</span>
            {issue.priority && <span className="text-xs text-jira-grey">P: {issue.priority}</span>}
            {issue.issueType && <span className="text-xs text-jira-grey">{issue.issueType}</span>}
          </div>
          {issue.assignee && (
            <div className="mt-2 text-xs text-jira-grey">
              Assignee: <span className="font-semibold text-jira-dark">{issue.assignee.name}</span>
            </div>
          )}
        </div>
      )}

      {/* Tabs */}
      <div className="flex border-b border-jira-border mb-3">
        <button className={tabClass(tab === 'teams')}    onClick={() => switchTab('teams')}>Teams</button>
        <button className={tabClass(tab === 'actions')}  onClick={() => switchTab('actions')}>Actions</button>
        <button className={tabClass(tab === 'activity')} onClick={() => switchTab('activity')}>Activity</button>
      </div>

      {err && <div className="alert-err">{err}</div>}
      {msg && <div className="alert-ok">{msg}</div>}

      {/* ── TEAMS TAB ── */}
      {tab === 'teams' && (
        <div className="space-y-2">
          {!mode && (
            <>
              <button className="btn-blue w-full" onClick={() => { setMode('dm'); reset(); }}>
                DM {issue?.assignee ? issue.assignee.name : 'Someone'}
              </button>
              <button className="btn-purple w-full" onClick={() => { setMode('group'); reset(); }}>
                Group Chat
              </button>
              <button className="btn-green w-full" onClick={() => { setMode('channel'); reset(); }}>
                Post to Channel
              </button>
            </>
          )}

          {mode === 'dm' && (
            <div className="space-y-2">
              <label className="label">To (email)</label>
              <input className="form-input" value={toEmail} onChange={e => setToEmail(e.target.value)} placeholder="user@company.com" />
              <p className="text-xs text-jira-grey">From: {currentUser?.email || '—'}</p>
              <div className="flex gap-2">
                <button className="btn-blue flex-1" onClick={submit} disabled={submitting}>{submitting ? 'Sending…' : 'Send DM'}</button>
                <button className="btn-ghost" onClick={reset}>Cancel</button>
              </div>
            </div>
          )}

          {mode === 'group' && (
            <div className="space-y-2">
              <label className="label">Emails (comma-separated)</label>
              <textarea className="form-textarea w-full h-16" value={groupEmails} onChange={e => setGroupEmails(e.target.value)} placeholder="user1@co.com, user2@co.com" />
              <div className="flex gap-2">
                <button className="btn-purple flex-1" onClick={submit} disabled={submitting}>{submitting ? 'Creating…' : 'Create Group Chat'}</button>
                <button className="btn-ghost" onClick={reset}>Cancel</button>
              </div>
            </div>
          )}

          {mode === 'channel' && (
            <div className="space-y-2">
              <p className="text-xs text-jira-grey">Posts to the channel configured in Teams Connector settings.</p>
              <div className="flex gap-2">
                <button className="btn-green flex-1" onClick={submit} disabled={submitting}>{submitting ? 'Posting…' : 'Post Now'}</button>
                <button className="btn-ghost" onClick={reset}>Cancel</button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── ACTIONS TAB ── */}
      {tab === 'actions' && (
        <div className="space-y-2">
          {!mode && (
            <>
              <button className="btn-blue w-full" onClick={() => { setMode('status'); setErr(''); setMsg(''); loadTransitions(); }}>
                Update Status
              </button>
              <button className="btn-green w-full" onClick={() => { setMode('comment'); reset(); }}>
                Add Comment
              </button>
              <button className="btn-purple w-full" onClick={() => { setMode('logtime'); reset(); }}>
                Log Time
              </button>
              <button className="btn-dark w-full" onClick={() => { setMode('assign'); reset(); }}>
                Assign Issue
              </button>
              <button className="btn-purple w-full" onClick={() => { setMode('editfields'); reset(); loadEditFields(); }}>
                Edit Fields
              </button>
            </>
          )}

          {mode === 'status' && (
            <div className="space-y-2">
              <div className="section-title">Select new status</div>
              {transitions.length === 0 && <p className="text-xs text-jira-grey">Loading…</p>}
              {transitions.map(t => (
                <button key={t.id} className="btn-ghost w-full text-left" onClick={() => applyTransition(t.id, t.name)} disabled={submitting}>
                  {t.name}
                </button>
              ))}
              <button className="btn-ghost w-full" onClick={reset}>Cancel</button>
            </div>
          )}

          {mode === 'comment' && (
            <div className="space-y-2">
              <label className="label">Comment</label>
              <textarea className="form-textarea w-full h-20" value={commentText} onChange={e => setCommentText(e.target.value)} placeholder="Enter your comment…" />
              <div className="flex gap-2">
                <button className="btn-green flex-1" onClick={submit} disabled={submitting}>{submitting ? 'Saving…' : 'Add Comment'}</button>
                <button className="btn-ghost" onClick={reset}>Cancel</button>
              </div>
            </div>
          )}

          {mode === 'logtime' && (
            <div className="space-y-2">
              <label className="label">Time Spent</label>
              <input className="form-input" value={timeSpent} onChange={e => setTimeSpent(e.target.value)} placeholder="e.g. 2h, 30m, 1h 30m" />
              <label className="label">Description (optional)</label>
              <textarea className="form-textarea w-full h-14" value={workDesc} onChange={e => setWorkDesc(e.target.value)} placeholder="What did you work on?" />
              <div className="flex gap-2">
                <button className="btn-purple flex-1" onClick={submit} disabled={submitting}>{submitting ? 'Logging…' : 'Log Time'}</button>
                <button className="btn-ghost" onClick={reset}>Cancel</button>
              </div>
            </div>
          )}

          {mode === 'assign' && (
            <div className="space-y-2">
              <label className="label">Assign to (email)</label>
              <input className="form-input" value={assignEmail} onChange={e => setAssignEmail(e.target.value)} placeholder="user@company.com" />
              <div className="flex gap-2">
                <button className="btn-dark flex-1" onClick={submit} disabled={submitting}>{submitting ? 'Assigning…' : 'Assign'}</button>
                <button className="btn-ghost" onClick={reset}>Cancel</button>
              </div>
            </div>
          )}

          {mode === 'editfields' && (
            <div className="space-y-2">
              <label className="label">Priority</label>
              <select className="form-select" value={editPriority} onChange={e => setEditPriority(e.target.value)}>
                <option value="">— no change —</option>
                {priorities.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
              <label className="label">Labels (comma-separated)</label>
              <input className="form-input" value={editLabels} onChange={e => setEditLabels(e.target.value)} placeholder="backend, urgent" />
              <label className="label">Due Date</label>
              <input type="date" className="form-input" value={editDueDate} onChange={e => setEditDueDate(e.target.value)} />
              <div className="flex gap-2">
                <button className="btn-purple flex-1" onClick={submit} disabled={submitting}>{submitting ? 'Saving…' : 'Save Changes'}</button>
                <button className="btn-ghost" onClick={reset}>Cancel</button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── ACTIVITY TAB ── */}
      {tab === 'activity' && (
        <div>
          {activity === null && <p className="text-xs text-jira-grey">Loading activity…</p>}
          {activity !== null && activity.length === 0 && <p className="text-xs text-jira-grey">No comments yet.</p>}
          {activity !== null && activity.map(c => (
            <div key={c.id} className="border-b border-jira-border pb-3 mb-3 last:border-0">
              <div className="flex justify-between items-center mb-1">
                <span className="text-xs font-semibold text-jira-dark">{c.author}</span>
                <span className="text-xs text-jira-grey">{new Date(c.created).toLocaleDateString()}</span>
              </div>
              <div className="text-xs text-jira-dark leading-relaxed">{c.text}</div>
            </div>
          ))}
          {activity !== null && activity.length > 0 && (
            <button className="btn-ghost text-xs mt-1" onClick={loadActivity}>Refresh</button>
          )}
        </div>
      )}
    </div>
  );
}

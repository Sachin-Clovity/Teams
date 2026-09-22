import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { SectionCard, Button, useToast, Skeleton, EmptyState, formatError } from './common';

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
  const showToast = useToast();

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
    }).catch(e => { setErr(formatError(e, 'Failed to load issue.')); setLoading(false); });
  }, [issueKey]);

  function reset() { setMode(null); setErr(''); }

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
    setErr(''); setSubmitting(true);
    try {
      if (mode === 'dm') {
        if (!toEmail.trim()) { setErr('Enter a recipient email.'); return; }
        await invoke('startDM', { fromEmail: currentUser.email, toEmail: toEmail.trim(), issueKey, issueSummary: issue.summary });
        showToast(`DM sent to ${toEmail.trim()}`); setMode(null);
      } else if (mode === 'group') {
        const emails = groupEmails.split(',').map(e => e.trim()).filter(Boolean);
        if (emails.length < 2) { setErr('Enter at least 2 emails.'); return; }
        await invoke('startGroupChat', { emails, issueKey, issueSummary: issue.summary });
        showToast(`Group chat created with ${emails.length} members.`); setMode(null);
      } else if (mode === 'channel') {
        await invoke('postToChannelManual', { issueKey, issueSummary: issue.summary });
        showToast('Posted to Teams channel.'); setMode(null);
      } else if (mode === 'comment') {
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
        setIssue(prev => ({ ...prev, assignee: { name: res.displayName, email: assignEmail.trim() } }));
        setMode(null);
      } else if (mode === 'editfields') {
        const payload = { issueKey, labels: editLabels.split(',').map(s => s.trim()).filter(Boolean) };
        if (editPriority) payload.priority = { id: editPriority };
        if (editDueDate) payload.dueDate = editDueDate;
        await invoke('updateIssueFields', payload);
        const priorityName = priorities.find(p => p.id === editPriority)?.name;
        showToast('Issue fields updated.');
        setIssue(prev => ({ ...prev, priority: priorityName || prev.priority, labels: payload.labels, dueDate: editDueDate || prev.dueDate }));
        setMode(null);
      }
    } catch (e) { setErr(formatError(e, 'Failed.')); }
    finally { setSubmitting(false); }
  }

  async function applyTransition(id, name) {
    setErr(''); setSubmitting(true);
    try {
      await invoke('transitionIssue', { issueKey, transitionId: id });
      showToast(`Status updated to "${name}".`);
      setIssue(prev => ({ ...prev, status: name }));
      setMode(null);
    } catch (e) { setErr(formatError(e, 'Transition failed.')); }
    finally { setSubmitting(false); }
  }

  const statusBadgeClass = status => {
    const map = { 'In Progress': 'bg-jira-blue text-white', 'Done': 'bg-jira-green text-white', 'In Review': 'bg-jira-purple text-white' };
    return `text-xs font-semibold px-2 py-0.5 rounded ${map[status] || 'bg-jira-border text-jira-dark'}`;
  };

  const tabClass = active => `tab-btn ${active ? 'border-jira-blue text-jira-blue' : 'border-transparent text-jira-grey hover:text-jira-dark'}`;

  if (loading) {
    return (
      <div className="p-3" aria-busy="true" aria-label="Loading">
        <div className="card mb-3">
          <Skeleton className="h-4 w-20 mb-2" />
          <Skeleton className="h-3 w-full mb-1" />
          <Skeleton className="h-3 w-2/3 mb-2" />
          <Skeleton className="h-5 w-16" />
        </div>
        <div className="flex gap-3 mb-3">
          <Skeleton className="h-6 w-14" /><Skeleton className="h-6 w-14" /><Skeleton className="h-6 w-14" />
        </div>
        <Skeleton className="h-8 w-full mb-2" />
        <Skeleton className="h-8 w-full mb-2" />
        <Skeleton className="h-8 w-full" />
      </div>
    );
  }

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

      {/* ── TEAMS TAB ── */}
      {tab === 'teams' && (
        <div className="space-y-2">
          {!mode && (
            <>
              <Button className="w-full" onClick={() => { setMode('dm'); reset(); }}>
                DM {issue?.assignee ? issue.assignee.name : 'Someone'}
              </Button>
              <Button variant="purple" className="w-full" onClick={() => { setMode('group'); reset(); }}>
                Group Chat
              </Button>
              <Button variant="success" className="w-full" onClick={() => { setMode('channel'); reset(); }}>
                Post to Channel
              </Button>
            </>
          )}

          {mode === 'dm' && (
            <SectionCard title="Send a Teams DM">
              <label className="label">To (email)</label>
              <input className="form-input" value={toEmail} onChange={e => setToEmail(e.target.value)} placeholder="user@company.com" />
              <p className="text-xs text-jira-grey mt-1">From: {currentUser?.email || '—'}</p>
              <div className="flex gap-2 mt-3">
                <Button className="flex-1" onClick={submit} loading={submitting}>Send DM</Button>
                <Button variant="ghost" onClick={reset}>Cancel</Button>
              </div>
            </SectionCard>
          )}

          {mode === 'group' && (
            <SectionCard title="Create a Group Chat">
              <label className="label">Emails (comma-separated)</label>
              <textarea className="form-textarea w-full h-16" value={groupEmails} onChange={e => setGroupEmails(e.target.value)} placeholder="user1@co.com, user2@co.com" />
              <div className="flex gap-2 mt-3">
                <Button variant="purple" className="flex-1" onClick={submit} loading={submitting}>Create Group Chat</Button>
                <Button variant="ghost" onClick={reset}>Cancel</Button>
              </div>
            </SectionCard>
          )}

          {mode === 'channel' && (
            <SectionCard title="Post to Channel" description="Posts to the channel configured in Teams Connector settings.">
              <div className="flex gap-2">
                <Button variant="success" className="flex-1" onClick={submit} loading={submitting}>Post Now</Button>
                <Button variant="ghost" onClick={reset}>Cancel</Button>
              </div>
            </SectionCard>
          )}
        </div>
      )}

      {/* ── ACTIONS TAB ── */}
      {tab === 'actions' && (
        <div className="space-y-2">
          {!mode && (
            <>
              <Button className="w-full" onClick={() => { setMode('status'); setErr(''); loadTransitions(); }}>
                Update Status
              </Button>
              <Button variant="success" className="w-full" onClick={() => { setMode('comment'); reset(); }}>
                Add Comment
              </Button>
              <Button variant="purple" className="w-full" onClick={() => { setMode('logtime'); reset(); }}>
                Log Time
              </Button>
              <Button variant="dark" className="w-full" onClick={() => { setMode('assign'); reset(); }}>
                Assign Issue
              </Button>
              <Button variant="purple" className="w-full" onClick={() => { setMode('editfields'); reset(); loadEditFields(); }}>
                Edit Fields
              </Button>
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
      )}

      {/* ── ACTIVITY TAB ── */}
      {tab === 'activity' && (
        <div>
          {activity === null && (
            <>
              <Skeleton className="h-10 w-full mb-2" />
              <Skeleton className="h-10 w-full" />
            </>
          )}
          {activity !== null && activity.length === 0 && <EmptyState icon="💬" title="No comments yet" />}
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
            <Button variant="ghost" className="text-xs mt-1" onClick={loadActivity}>Refresh</Button>
          )}
        </div>
      )}
    </div>
  );
}

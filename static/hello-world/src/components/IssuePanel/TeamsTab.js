import React, { useState } from 'react';
import { invoke } from '@forge/bridge';
import { SectionCard, Button, useToast, formatError } from '../common';

// DM / Group Chat / Post to Channel — self-contained: unmounting this tab (switching away)
// naturally discards its mode/form state instead of needing an explicit reset() call.
export default function TeamsTab({ issueKey, issue, currentUser }) {
  const showToast = useToast();
  const [mode,        setMode]        = useState(null);
  const [toEmail,     setToEmail]     = useState(issue?.assignee?.email || '');
  const [groupEmails, setGroupEmails] = useState([issue?.assignee?.email, issue?.reporter?.email].filter(Boolean).join(', '));
  const [submitting,  setSubmitting]  = useState(false);
  const [err,         setErr]         = useState('');

  function reset() { setMode(null); setErr(''); }

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
      }
    } catch (e) { setErr(formatError(e, 'Failed.')); }
    finally { setSubmitting(false); }
  }

  return (
    <div className="space-y-2">
      {err && <div className="alert-err">{err}</div>}

      {!mode && (
        <>
          <Button className="w-full" onClick={() => { setMode('dm'); setErr(''); }}>
            DM {issue?.assignee ? issue.assignee.name : 'Someone'}
          </Button>
          <Button variant="purple" className="w-full" onClick={() => { setMode('group'); setErr(''); }}>
            Group Chat
          </Button>
          <Button variant="success" className="w-full" onClick={() => { setMode('channel'); setErr(''); }}>
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
  );
}

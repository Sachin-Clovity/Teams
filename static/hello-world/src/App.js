import React, { useEffect, useState } from 'react';
import { invoke, view } from '@forge/bridge';

// ── Admin / Global Settings Page ─────────────────────────────────────────────
function AdminPage() {
  const [connected,            setConnected]            = useState(false);
  const [teams,                setTeams]                = useState([]);
  const [channels,             setChannels]             = useState([]);
  const [teamId,               setTeamId]               = useState('');
  const [channelId,            setChannelId]            = useState('');
  const [webhookUrl,           setWebhookUrlInput]      = useState('');
  const [savedConfig,          setSavedConfig]          = useState(null);
  const [automationWebhookUrl, setAutomationWebhookUrl] = useState('');
  const [loading,              setLoading]              = useState(false);
  const [err,                  setErr]                  = useState('');
  const [msg,                  setMsg]                  = useState('');
  const [consoleLogs,          setConsoleLogs]          = useState([]);
  const [showConsole,          setShowConsole]          = useState(false);
  const [customMsg,            setCustomMsg]            = useState('');
  const [notifSettings,        setNotifSettings]        = useState({ created: true, updated: true, deleted: false, commented: false });
  const [notifSaved,           setNotifSaved]           = useState(false);

  useEffect(() => {
    invoke('getConfig').then(({ config }) => { if (config) setSavedConfig(config); }).catch(() => {});
    invoke('getWebhookUrl').then(({ url }) => setAutomationWebhookUrl(url)).catch(() => {});
    invoke('getNotificationSettings').then(({ settings }) => { if (settings) setNotifSettings(settings); }).catch(() => {});
  }, []);

  const clear = () => { setErr(''); setMsg(''); };

  async function connect() {
    clear(); setLoading(true);
    try {
      const res = await invoke('getTeams');
      if (!res.teams?.length) { setErr('Connected but no Teams found. Check Azure Application permissions.'); return; }
      setTeams(res.teams);
      setConnected(true);
      setMsg('Connected. Select a Team below.');
    } catch (e) { setErr(e.message || 'Connection failed. Check Forge logs.'); }
    finally { setLoading(false); }
  }

  async function onTeamChange(e) {
    const id = e.target.value;
    setTeamId(id); setChannelId(''); setChannels([]);
    if (!id) return;
    try {
      const res = await invoke('getChannels', { teamId: id });
      setChannels(res.channels || []);
    } catch (e) { setErr(e.message || 'Failed to load channels.'); }
  }

  async function save() {
    clear();
    if (!teamId || !channelId) { setErr('Select both a Team and a Channel.'); return; }
    setLoading(true);
    try {
      const team    = teams.find(t => t.id === teamId);
      const channel = channels.find(c => c.id === channelId);
      await invoke('saveConfig', { teamId, channelId, teamName: team?.displayName || teamId, channelName: channel?.displayName || channelId, webhookUrl });
      setSavedConfig({ teamId, channelId, teamName: team?.displayName || teamId, channelName: channel?.displayName || channelId });
      setMsg('Saved. Jira issue events will now post to this channel.');
    } catch (e) { setErr(e.message || 'Save failed.'); }
    finally { setLoading(false); }
  }

  async function test() {
    clear(); setLoading(true);
    try {
      const res = await invoke('testConnection');
      res.success ? setMsg('Test message sent. Check your Teams channel.') : setErr(`Test failed: ${res.error}`);
    } catch (e) { setErr(e.message || 'Test failed.'); }
    finally { setLoading(false); }
  }

  async function sendCustomMessage() {
    clear(); setLoading(true);
    try {
      const res = await invoke('sendCustomMessage', { message: customMsg });
      res.success ? setMsg('Message sent to Teams!') : setErr(`Failed: ${res.error}`);
    } catch (e) { setErr(e.message || 'Send failed.'); }
    finally { setLoading(false); }
  }

  async function checkAuthStatus() {
    clear();
    try {
      const res = await invoke('getAuthStatus');
      const lines = res.authenticated
        ? [{ key: 'authenticated', val: 'true', isErr: false }, { key: 'org', val: res.org || '—', isErr: false }]
        : [{ key: 'authenticated', val: 'false', isErr: true }, { key: 'error', val: res.error || 'Token request failed', isErr: true }];
      setConsoleLogs(prev => [{ time: new Date().toLocaleTimeString(), lines }, ...prev].slice(0, 20));
      setShowConsole(true);
    } catch (e) {
      setConsoleLogs(prev => [{ time: new Date().toLocaleTimeString(), lines: [{ key: 'error', val: e.message, isErr: true }] }, ...prev].slice(0, 20));
      setShowConsole(true);
    }
  }

  async function checkBotDebug() {
    clear();
    try {
      const res  = await invoke('getBotDebug');
      const data = res.debug;
      const lines = typeof data === 'string'
        ? [{ key: 'info', val: data, isErr: false }]
        : Object.entries(data).map(([k, v]) => ({ key: k, val: typeof v === 'boolean' ? String(v) : String(v ?? '—'), isErr: k === 'error' || k === 'detail' }));
      setConsoleLogs(prev => [{ time: new Date().toLocaleTimeString(), lines }, ...prev].slice(0, 20));
      setShowConsole(true);
    } catch (e) {
      setConsoleLogs(prev => [{ time: new Date().toLocaleTimeString(), lines: [{ key: 'error', val: e.message, isErr: true }] }, ...prev].slice(0, 20));
      setShowConsole(true);
    }
  }

  async function disconnect() {
    clear();
    await invoke('clearConfig').catch(() => {});
    setConnected(false); setTeams([]); setChannels([]);
    setTeamId(''); setChannelId(''); setSavedConfig(null);
    setMsg('Disconnected and configuration cleared.');
  }

  return (
    <div className="max-w-xl mx-auto p-6 font-sans">

      {/* Header */}
      <div className="mb-6">
        <div className="flex items-center gap-3 mb-1">
          <div className="w-8 h-8 bg-jira-blue rounded flex items-center justify-center text-white font-bold text-sm">J</div>
          <h1 className="text-lg font-bold text-jira-dark">Microsoft Teams Connector</h1>
        </div>
        <p className="text-xs text-jira-grey ml-11">Send Jira issue events to Teams via Microsoft Graph API</p>
      </div>

      {err && <div className="alert-err">{err}</div>}
      {msg && <div className="alert-ok">{msg}</div>}

      {/* Step 1 — Connect */}
      <div className="section-title">Step 1 — Connect to Microsoft Teams</div>
      {!connected ? (
        <button className="btn-blue w-full" onClick={connect} disabled={loading}>
          {loading ? 'Connecting…' : 'Connect Microsoft Teams'}
        </button>
      ) : (
        <>
          <div className="card flex items-center gap-3">
            <div className="w-2 h-2 rounded-full bg-jira-green"></div>
            <span className="text-sm font-semibold text-jira-dark">Connected to Microsoft Teams</span>
          </div>

          <div className="divider" />
          <div className="section-title">Step 2 — Select Destination</div>

          <label className="label">Team</label>
          <select className="form-select mb-3" value={teamId} onChange={onTeamChange}>
            <option value="">— choose a team —</option>
            {teams.map(t => <option key={t.id} value={t.id}>{t.displayName}</option>)}
          </select>

          {channels.length > 0 && (
            <>
              <label className="label">Channel</label>
              <select className="form-select mb-3" value={channelId} onChange={e => setChannelId(e.target.value)}>
                <option value="">— choose a channel —</option>
                {channels.map(c => <option key={c.id} value={c.id}>{c.displayName}</option>)}
              </select>
            </>
          )}

          <label className="label">Teams Incoming Webhook URL</label>
          <input
            className="form-input mb-1 text-xs"
            type="text"
            placeholder="https://outlook.office.com/webhook/..."
            value={webhookUrl}
            onChange={e => setWebhookUrlInput(e.target.value)}
          />
          <p className="text-xs text-jira-grey mb-4">Teams channel → ··· → Connectors → Incoming Webhook → Configure → Copy URL</p>

          <div className="flex gap-2">
            <button
              className={(!teamId || !channelId || !webhookUrl || loading) ? 'btn bg-jira-border text-jira-grey cursor-not-allowed flex-1' : 'btn-blue flex-1'}
              onClick={save}
              disabled={!teamId || !channelId || !webhookUrl || loading}
            >
              {loading ? 'Saving…' : 'Save Configuration'}
            </button>
            <button className="btn-red" onClick={disconnect}>Disconnect</button>
          </div>
        </>
      )}

      {/* Step 3 — Active Config */}
      {savedConfig && (
        <>
          <div className="divider" />
          <div className="section-title">Step 3 — Active Configuration</div>
          <div className="card">
            <div className="grid grid-cols-2 gap-3">
              <div>
                <div className="text-xs text-jira-grey uppercase tracking-wide mb-1">Team</div>
                <div className="text-sm font-semibold text-jira-dark">{savedConfig.teamName}</div>
              </div>
              <div>
                <div className="text-xs text-jira-grey uppercase tracking-wide mb-1">Channel</div>
                <div className="text-sm font-semibold text-jira-dark">{savedConfig.channelName}</div>
              </div>
            </div>
          </div>

          <button className="btn-green w-full mb-3" onClick={test} disabled={loading}>
            {loading ? 'Sending…' : 'Send Test Message to Teams'}
          </button>

          {/* Custom message */}
          <div className="divider" />
          <div className="section-title">Send Custom Message</div>
          <textarea
            className="form-textarea w-full mb-2 h-20"
            placeholder="Type your message here..."
            value={customMsg}
            onChange={e => setCustomMsg(e.target.value)}
          />
          <button
            className={(!customMsg.trim() || loading) ? 'btn bg-jira-border text-jira-grey cursor-not-allowed w-full' : 'btn-blue w-full'}
            onClick={sendCustomMessage}
            disabled={!customMsg.trim() || loading}
          >
            {loading ? 'Sending…' : 'Send to Teams'}
          </button>

          {/* Notification Settings */}
          <div className="divider" />
          <div className="section-title">Notification Settings</div>
          <p className="text-xs text-jira-grey mb-3">Choose which Jira events send a Teams notification.</p>
          <div className="card space-y-3">
            {[
              { key: 'created',   label: 'Issue Created',   desc: 'Notify when a new issue is created' },
              { key: 'updated',   label: 'Issue Updated',   desc: 'Notify when an issue is modified' },
              { key: 'deleted',   label: 'Issue Deleted',   desc: 'Notify when an issue is deleted' },
              { key: 'commented', label: 'Comment Added',   desc: 'Notify when a comment is added via bot' },
            ].map(({ key, label, desc }) => (
              <label key={key} className="flex items-start gap-3 cursor-pointer group">
                <input
                  type="checkbox"
                  className="mt-0.5 w-4 h-4 accent-jira-blue"
                  checked={!!notifSettings[key]}
                  onChange={e => setNotifSettings(prev => ({ ...prev, [key]: e.target.checked }))}
                />
                <div>
                  <div className="text-sm font-medium text-jira-dark">{label}</div>
                  <div className="text-xs text-jira-grey">{desc}</div>
                </div>
              </label>
            ))}
          </div>
          <div className="flex items-center gap-3 mt-2">
            <button
              className="btn-blue"
              onClick={async () => {
                await invoke('saveNotificationSettings', { settings: notifSettings }).catch(() => {});
                setNotifSaved(true);
                setTimeout(() => setNotifSaved(false), 2000);
              }}
            >
              Save Notification Settings
            </button>
            {notifSaved && <span className="text-jira-green text-xs font-semibold">Saved!</span>}
          </div>

          {/* Automation Webhook URL */}
          {webhookUrl && (
            <>
              <div className="divider" />
              <div className="section-title">Automation Webhook URL</div>
              <div className="card">
                <div className="text-xs text-jira-grey mb-1">Use this URL in Jira Automation rules</div>
                <div className="text-xs font-mono text-jira-dark break-all bg-white border border-jira-border rounded p-2 mt-1">
                  {automationWebhookUrl}
                </div>
              </div>
              <p className="text-xs text-jira-grey">
                In Jira Automation: Add action → Send web request → POST to this URL<br />
                Body: <code className="bg-jira-light px-1 rounded">{"{'title':'{{issue.summary}}','issueKey':'{{issue.key}}'}"}</code>
              </p>
            </>
          )}

          {/* Debug */}
          <div className="divider" />
          <div className="flex gap-2 flex-wrap">
            <button className="btn-blue text-xs" onClick={checkAuthStatus}>Check MS Auth</button>
            <button className="btn-dark text-xs" onClick={checkBotDebug}>Bot Console</button>
            <button className="btn-ghost text-xs" onClick={() => { setConsoleLogs([]); setShowConsole(false); }}>Clear</button>
          </div>
        </>
      )}

      {/* Debug Console */}
      {showConsole && (
        <>
          <div className="divider" />
          <div className="section-title">Debug Console</div>
          <div className="bg-gray-900 rounded-lg p-3 font-mono text-xs max-h-56 overflow-y-auto space-y-2">
            {consoleLogs.length === 0 && <div className="text-gray-500">No activity yet.</div>}
            {consoleLogs.map((entry, i) => (
              <div key={i} className="border-b border-gray-700 pb-2">
                <div className="text-gray-500 mb-1">── {entry.time} ──</div>
                {entry.lines.map((line, j) => (
                  <div key={j}>
                    <span className={line.isErr ? 'text-red-400' : 'text-blue-400'}>{line.key}: </span>
                    <span className={line.isErr ? 'text-red-300' : 'text-green-300'}>{line.val}</span>
                  </div>
                ))}
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

// ── Issue Context Panel ───────────────────────────────────────────────────────
function IssuePanel({ issueKey }) {
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

// ── Root ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [moduleKey, setModuleKey] = useState(null);
  const [issueKey,  setIssueKey]  = useState(null);

  useEffect(() => {
    view.getContext()
      .then(ctx => {
        setModuleKey(ctx.moduleKey || 'admin');
        setIssueKey(ctx.extension?.issue?.key || null);
      })
      .catch(() => setModuleKey('admin'));
  }, []);

  if (!moduleKey) return <div className="p-4 text-jira-grey text-sm">Loading…</div>;
  if (moduleKey === 'teams-issue-context') return <IssuePanel issueKey={issueKey} />;
  return <AdminPage />;
}

import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { PageHeader, SectionCard, Stepper } from './common';

const NOTIFICATION_FIELDS = [
  { key: 'status', label: 'Status' }, { key: 'priority', label: 'Priority' },
  { key: 'issueType', label: 'Type' }, { key: 'assignee', label: 'Assignee' },
  { key: 'reporter', label: 'Reporter' }, { key: 'dueDate', label: 'Due Date' },
  { key: 'labels', label: 'Labels' },
];

// ── Admin / Global Settings Page ─────────────────────────────────────────────
export default function AdminPage() {
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
  const [fields,                setFields]              = useState(['status', 'priority', 'issueType', 'assignee', 'reporter']);

  useEffect(() => {
    invoke('getConfig').then(({ config }) => { if (config) { setSavedConfig(config); if (config.fields?.length) setFields(config.fields); } }).catch(() => {});
    invoke('getWebhookUrl').then(({ url }) => setAutomationWebhookUrl(url)).catch(() => {});
    invoke('getNotificationSettings').then(({ settings }) => { if (settings) setNotifSettings(settings); }).catch(() => {});
  }, []);

  function toggleField(key) {
    setFields(prev => prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]);
  }

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
      await invoke('saveConfig', { teamId, channelId, teamName: team?.displayName || teamId, channelName: channel?.displayName || channelId, webhookUrl, fields });
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

  const stepIndex = savedConfig ? 2 : connected ? 1 : 0;

  return (
    <div className="max-w-xl mx-auto p-6 font-sans">
      <PageHeader title="Microsoft Teams Connector" subtitle="Send Jira issue events to Teams via Microsoft Graph API" />

      <Stepper steps={['Connect', 'Destination', 'Configure']} currentIndex={stepIndex} />

      {err && <div className="alert-err">{err}</div>}
      {msg && <div className="alert-ok">{msg}</div>}

      {/* Step 1 — Connect */}
      <SectionCard title="Step 1 — Connect to Microsoft Teams" tone={connected ? 'default' : 'highlight'}>
        {!connected ? (
          <button className="btn-blue w-full" onClick={connect} disabled={loading}>
            {loading ? 'Connecting…' : 'Connect Microsoft Teams'}
          </button>
        ) : (
          <div className="flex items-center gap-3">
            <div className="w-2 h-2 rounded-full bg-jira-green"></div>
            <span className="text-sm font-semibold text-jira-dark">Connected to Microsoft Teams</span>
          </div>
        )}
      </SectionCard>

      {/* Step 2 — Destination */}
      {connected && (
        <SectionCard title="Step 2 — Select Destination" tone={savedConfig ? 'default' : 'highlight'}>
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
          <p className="text-xs text-jira-grey mb-4">Teams channel → ··· → Workflows → "Send webhook alerts to a channel" template → Save → Copy URL</p>

          <label className="label">Fields to include in notifications</label>
          <div className="card space-y-2 mb-4">
            {NOTIFICATION_FIELDS.map(f => (
              <label key={f.key} className="flex items-center gap-2 cursor-pointer">
                <input type="checkbox" className="w-4 h-4 accent-jira-blue" checked={fields.includes(f.key)} onChange={() => toggleField(f.key)} />
                <span className="text-sm text-jira-dark">{f.label}</span>
              </label>
            ))}
          </div>

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
        </SectionCard>
      )}

      {/* Step 3 — Active Config + management */}
      {savedConfig && (
        <>
          <SectionCard title="Step 3 — Active Configuration" tone="highlight">
            <div className="grid grid-cols-2 gap-3 mb-3">
              <div>
                <div className="text-xs text-jira-grey uppercase tracking-wide mb-1">Team</div>
                <div className="text-sm font-semibold text-jira-dark">{savedConfig.teamName}</div>
              </div>
              <div>
                <div className="text-xs text-jira-grey uppercase tracking-wide mb-1">Channel</div>
                <div className="text-sm font-semibold text-jira-dark">{savedConfig.channelName}</div>
              </div>
            </div>
            <button className="btn-green w-full" onClick={test} disabled={loading}>
              {loading ? 'Sending…' : 'Send Test Message to Teams'}
            </button>
          </SectionCard>

          <SectionCard title="Send Custom Message">
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
          </SectionCard>

          <SectionCard title="Notification Settings" description="Choose which Jira events send a Teams notification.">
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
                    onChange={e => { const checked = e.target.checked; setNotifSettings(prev => ({ ...prev, [key]: checked })); }}
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
          </SectionCard>

          {/* Automation Webhook URL */}
          {webhookUrl && (
            <SectionCard title="Automation Webhook URL" description="Use this URL in Jira Automation rules.">
              <div className="card">
                <div className="text-xs font-mono text-jira-dark break-all bg-white border border-jira-border rounded p-2">
                  {automationWebhookUrl}
                </div>
              </div>
              <p className="text-xs text-jira-grey">
                In Jira Automation: Add action → Send web request → POST to this URL<br />
                Body: <code className="bg-jira-light px-1 rounded">{"{'title':'{{issue.summary}}','issueKey':'{{issue.key}}'}"}</code>
              </p>
            </SectionCard>
          )}

          <SectionCard title="Debug Tools">
            <div className="flex gap-2 flex-wrap">
              <button className="btn-blue text-xs" onClick={checkAuthStatus}>Check MS Auth</button>
              <button className="btn-dark text-xs" onClick={checkBotDebug}>Bot Console</button>
              <button className="btn-ghost text-xs" onClick={() => { setConsoleLogs([]); setShowConsole(false); }}>Clear</button>
            </div>

            {showConsole && (
              <div className="bg-gray-900 rounded-lg p-3 font-mono text-xs max-h-56 overflow-y-auto space-y-2 mt-3">
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
            )}
          </SectionCard>
        </>
      )}
    </div>
  );
}

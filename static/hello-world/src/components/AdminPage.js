import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { PageHeader, SectionCard, Stepper, Button, useToast, LogTable, formatError } from './common';

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
  const [consoleLogs,          setConsoleLogs]          = useState([]);
  const [showConsole,          setShowConsole]          = useState(false);
  const [customMsg,            setCustomMsg]            = useState('');
  const [notifSettings,        setNotifSettings]        = useState({ created: true, updated: true, deleted: false, commented: false });
  const [fields,                setFields]              = useState(['status', 'priority', 'issueType', 'assignee', 'reporter']);
  const showToast = useToast();

  useEffect(() => {
    invoke('getConfig').then(({ config }) => { if (config) { setSavedConfig(config); if (config.fields?.length) setFields(config.fields); } }).catch(() => {});
    invoke('getWebhookUrl').then(({ url }) => setAutomationWebhookUrl(url)).catch(() => {});
    invoke('getNotificationSettings').then(({ settings }) => { if (settings) setNotifSettings(settings); }).catch(() => {});
  }, []);

  function toggleField(key) {
    setFields(prev => prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]);
  }

  const clear = () => setErr('');

  async function connect() {
    clear(); setLoading(true);
    try {
      const res = await invoke('getTeams');
      if (!res.teams?.length) { setErr('Connected but no Teams found. Check Azure Application permissions.'); return; }
      setTeams(res.teams);
      setConnected(true);
      showToast('Connected. Select a Team below.');
    } catch (e) { setErr(formatError(e, 'Connection failed. Check Forge logs.')); }
    finally { setLoading(false); }
  }

  async function onTeamChange(e) {
    const id = e.target.value;
    setTeamId(id); setChannelId(''); setChannels([]);
    if (!id) return;
    try {
      const res = await invoke('getChannels', { teamId: id });
      setChannels(res.channels || []);
    } catch (e) { setErr(formatError(e, 'Failed to load channels.')); }
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
      showToast('Saved. Jira issue events will now post to this channel.');
    } catch (e) { setErr(formatError(e, 'Save failed.')); }
    finally { setLoading(false); }
  }

  async function test() {
    clear(); setLoading(true);
    try {
      const res = await invoke('testConnection');
      res.success ? showToast('Test message sent. Check your Teams channel.') : showToast(formatError({ message: res.error }, 'Test failed.'), 'error');
    } catch (e) { showToast(formatError(e, 'Test failed.'), 'error'); }
    finally { setLoading(false); }
  }

  async function sendCustomMessage() {
    clear(); setLoading(true);
    try {
      const res = await invoke('sendCustomMessage', { message: customMsg });
      res.success ? showToast('Message sent to Teams!') : showToast(formatError({ message: res.error }, 'Send failed.'), 'error');
    } catch (e) { showToast(formatError(e, 'Send failed.'), 'error'); }
    finally { setLoading(false); }
  }

  async function checkAuthStatus() {
    clear();
    try {
      const res = await invoke('getAuthStatus');
      const lines = res.authenticated
        ? [{ key: 'authenticated', val: 'true', isErr: false }, { key: 'org', val: res.org || '—', isErr: false }, { key: 'tenantId', val: res.tenantId || '(default — no bot traffic seen yet)', isErr: false }]
        : [{ key: 'authenticated', val: 'false', isErr: true }, { key: 'error', val: res.error || 'Token request failed', isErr: true }];
      setConsoleLogs(prev => [{ time: new Date().toLocaleTimeString(), lines }, ...prev].slice(0, 20));
      setShowConsole(true);
    } catch (e) {
      setConsoleLogs(prev => [{ time: new Date().toLocaleTimeString(), lines: [{ key: 'error', val: formatError(e), isErr: true }] }, ...prev].slice(0, 20));
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
      setConsoleLogs(prev => [{ time: new Date().toLocaleTimeString(), lines: [{ key: 'error', val: formatError(e), isErr: true }] }, ...prev].slice(0, 20));
      setShowConsole(true);
    }
  }

  async function disconnect() {
    clear();
    try {
      await invoke('clearConfig');
      setConnected(false); setTeams([]); setChannels([]);
      setTeamId(''); setChannelId(''); setSavedConfig(null);
      showToast('Disconnected and configuration cleared.');
    } catch (e) { showToast(formatError(e, 'Disconnect failed.'), 'error'); }
  }

  const stepIndex = savedConfig ? 2 : connected ? 1 : 0;

  return (
    <div className="max-w-6xl mx-auto p-6 lg:p-8 font-sans">
      <PageHeader title="Microsoft Teams Connector" subtitle="Send Jira issue events to Teams via Microsoft Graph API" />

      <Stepper steps={['Connect', 'Destination', 'Configure']} currentIndex={stepIndex} />

      {err && <div className="alert-err">{err}</div>}

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-start">
        {/* ── Main column — the setup flow itself ── */}
        <div className="lg:col-span-2">
          {/* Step 1 — Connect */}
          <SectionCard title="Step 1 — Connect to Microsoft Teams" tone={connected ? 'default' : 'highlight'}>
            {!connected ? (
              <Button className="w-full" onClick={connect} loading={loading}>Connect Microsoft Teams</Button>
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
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4">
                <div>
                  <label className="label mt-0">Team</label>
                  <select className="form-select mb-3" value={teamId} onChange={onTeamChange}>
                    <option value="">— choose a team —</option>
                    {teams.map(t => <option key={t.id} value={t.id}>{t.displayName}</option>)}
                  </select>
                </div>

                {channels.length > 0 && (
                  <div>
                    <label className="label mt-0">Channel</label>
                    <select className="form-select mb-3" value={channelId} onChange={e => setChannelId(e.target.value)}>
                      <option value="">— choose a channel —</option>
                      {channels.map(c => <option key={c.id} value={c.id}>{c.displayName}</option>)}
                    </select>
                  </div>
                )}
              </div>

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
              <div className="card grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-2 mb-4">
                {NOTIFICATION_FIELDS.map(f => (
                  <label key={f.key} className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" className="w-4 h-4 accent-jira-blue" checked={fields.includes(f.key)} onChange={() => toggleField(f.key)} />
                    <span className="text-sm text-jira-dark">{f.label}</span>
                  </label>
                ))}
              </div>

              <div className="flex gap-2">
                <Button className="flex-1" onClick={save} loading={loading} disabled={!teamId || !channelId || !webhookUrl}>
                  Save Configuration
                </Button>
                <Button variant="danger" onClick={disconnect}>Disconnect</Button>
              </div>
            </SectionCard>
          )}

          {/* Step 3 — Active Config */}
          {savedConfig && (
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
              <Button variant="success" className="w-full" onClick={test} loading={loading}>Send Test Message to Teams</Button>
            </SectionCard>
          )}

          {savedConfig && (
            <SectionCard title="Notification Settings" description="Choose which Jira events send a Teams notification.">
              <div className="card grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-3">
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
              <Button
                onClick={async () => {
                  try { await invoke('saveNotificationSettings', { settings: notifSettings }); showToast('Notification settings saved.'); }
                  catch (e) { showToast(formatError(e, 'Save failed.'), 'error'); }
                }}
              >
                Save Notification Settings
              </Button>
            </SectionCard>
          )}
        </div>

        {/* ── Sidebar column — supplementary tools, only relevant once configured ── */}
        {savedConfig && (
          <div>
            <SectionCard title="Send Custom Message">
              <textarea
                className="form-textarea w-full mb-2 h-20"
                placeholder="Type your message here..."
                value={customMsg}
                onChange={e => setCustomMsg(e.target.value)}
              />
              <Button className="w-full" onClick={sendCustomMessage} loading={loading} disabled={!customMsg.trim()}>Send to Teams</Button>
            </SectionCard>

            {webhookUrl && (
              <SectionCard title="Automation Webhook URL" description="Use this URL in Jira Automation rules.">
                <div className="card">
                  <div className="text-xs font-mono text-jira-dark break-all bg-white border border-jira-border rounded p-2">
                    {automationWebhookUrl}
                  </div>
                </div>
                <p className="text-xs text-jira-grey">
                  In Jira Automation: Add action → Send web request → POST to this URL<br />
                  Body: <code className="bg-jira-light px-1 rounded break-all">{'{"title":"{{issue.summary}}","issueKey":"{{issue.key}}","issueUrl":"{{baseUrl}}/browse/{{issue.key}}"}'}</code>
                  <br />
                  This works from any Jira site — no app install needed there. Include <code className="bg-jira-light px-1 rounded">issueUrl</code> using the <code className="bg-jira-light px-1 rounded">{'{{baseUrl}}'}</code> smart value so the link points at the right site.
                </p>
              </SectionCard>
            )}

            <SectionCard title="Debug Tools">
              <div className="flex gap-2 flex-wrap mb-3">
                <Button className="text-xs" onClick={checkAuthStatus}>Check MS Auth</Button>
                <Button variant="dark" className="text-xs" onClick={checkBotDebug}>Bot Console</Button>
                <Button variant="ghost" className="text-xs" onClick={() => { setConsoleLogs([]); setShowConsole(false); }}>Clear</Button>
              </div>

              {showConsole && <LogTable entries={consoleLogs} />}
            </SectionCard>
          </div>
        )}
      </div>
    </div>
  );
}

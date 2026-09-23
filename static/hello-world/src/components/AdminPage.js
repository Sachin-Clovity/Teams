import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { PageHeader, SectionCard, Stepper, Button, useToast, LogTable, formatError } from './common';

const NOTIFICATION_FIELDS = [
  { key: 'status', label: 'Status' }, { key: 'priority', label: 'Priority' },
  { key: 'issueType', label: 'Type' }, { key: 'assignee', label: 'Assignee' },
  { key: 'reporter', label: 'Reporter' }, { key: 'dueDate', label: 'Due Date' },
  { key: 'labels', label: 'Labels' },
];

// Sentinel used in the routing form's Project dropdown to mean "the global default,"
// as opposed to any specific project key.
const DEFAULT_ROUTE = '';

// ── Admin / Global Settings Page ─────────────────────────────────────────────
export default function AdminPage() {
  const [connected,            setConnected]            = useState(false);
  const [teams,                setTeams]                = useState([]);
  const [loading,              setLoading]              = useState(false);
  const [err,                  setErr]                  = useState('');
  const [savedConfig,          setSavedConfig]          = useState(null);
  const [automationWebhookUrl, setAutomationWebhookUrl] = useState('');
  const [consoleLogs,          setConsoleLogs]          = useState([]);
  const [showConsole,          setShowConsole]          = useState(false);
  const [customMsg,            setCustomMsg]            = useState('');
  const [notifSettings,        setNotifSettings]        = useState({ created: true, updated: true, deleted: false, commented: false });
  const [fields,                setFields]              = useState(['status', 'priority', 'issueType', 'assignee', 'reporter']);

  // ── Channel routing — one unified table. The global default and every project-specific
  // override are really the same shape of data (team/channel/webhook); they used to be two
  // separate sections with duplicated form code, which read as two different features when
  // they're really one "where does this go" question answered at two different scopes. ──
  const [mapProjects,    setMapProjects]    = useState([]);   // every project on the site, for the picker
  const [projectConfigs, setProjectConfigs] = useState([]);   // saved project-specific routes
  const [routeProjectKey, setRouteProjectKey] = useState(DEFAULT_ROUTE);
  const [routeTeamId,     setRouteTeamId]     = useState('');
  const [routeChannelId,  setRouteChannelId]  = useState('');
  const [routeChannels,   setRouteChannels]   = useState([]);
  const [routeWebhookUrl, setRouteWebhookUrl] = useState('');
  const [routeSaving,     setRouteSaving]     = useState(false);

  const showToast = useToast();

  function loadProjectConfigs() {
    invoke('getAllProjectConfigs').then(({ configs }) => setProjectConfigs(configs || [])).catch(() => {});
  }

  useEffect(() => {
    // A saved global config is durable proof that Connect already succeeded at some point —
    // restore that state on every load instead of always starting from "not connected," which
    // only reflected whether this particular React mount happened to still remember an
    // earlier click (inconsistent, since Jira doesn't always tear the custom-UI iframe down
    // between navigations).
    invoke('getConfig').then(async ({ config }) => {
      if (!config) return;
      setSavedConfig(config);
      if (config.fields?.length) setFields(config.fields);
      setConnected(true);
      try {
        const { teams: loadedTeams } = await invoke('getTeams');
        setTeams(loadedTeams || []);
      } catch (e) {
        // Non-fatal — the saved team/channel names still show in the routing table below;
        // only the picker dropdowns would be missing their live choice list until reloaded.
      }
    }).catch(() => {});
    invoke('getWebhookUrl').then(({ url }) => setAutomationWebhookUrl(url)).catch(() => {});
    invoke('getNotificationSettings').then(({ settings }) => { if (settings) setNotifSettings(settings); }).catch(() => {});
    invoke('searchProjects').then(({ projects }) => setMapProjects(projects || [])).catch(() => {});
    loadProjectConfigs();
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
      showToast('Connected. Set up channel routing below.');
    } catch (e) { setErr(formatError(e, 'Connection failed. Check Forge logs.')); }
    finally { setLoading(false); }
  }

  async function onRouteTeamChange(e) {
    const id = e.target.value;
    setRouteTeamId(id); setRouteChannelId(''); setRouteChannels([]);
    if (!id) return;
    try {
      const res = await invoke('getChannels', { teamId: id });
      setRouteChannels(res.channels || []);
    } catch (e) { showToast(formatError(e, 'Failed to load channels.'), 'error'); }
  }

  function resetRouteForm() {
    setRouteProjectKey(DEFAULT_ROUTE); setRouteTeamId(''); setRouteChannelId(''); setRouteChannels([]); setRouteWebhookUrl('');
  }

  // Populates the form from an existing row (Default or a specific project) so "Edit" can
  // change it, instead of only ever being able to add brand new routes.
  async function editRoute(projectKey) {
    setRouteProjectKey(projectKey);
    const cfg = projectKey === DEFAULT_ROUTE ? savedConfig : projectConfigs.find(p => p.projectKey === projectKey)?.config;
    setRouteTeamId(cfg?.teamId || ''); setRouteChannelId(cfg?.channelId || ''); setRouteWebhookUrl(cfg?.webhookUrl || '');
    setRouteChannels([]);
    if (cfg?.teamId) {
      try {
        const res = await invoke('getChannels', { teamId: cfg.teamId });
        setRouteChannels(res.channels || []);
      } catch (e) { showToast(formatError(e, 'Failed to load channels.'), 'error'); }
    }
  }

  async function saveRoute() {
    if (!routeTeamId || !routeChannelId || !routeWebhookUrl.trim()) {
      showToast('Pick a team, channel, and webhook URL first.', 'error');
      return;
    }
    setRouteSaving(true);
    try {
      const team    = teams.find(t => t.id === routeTeamId);
      const channel = routeChannels.find(c => c.id === routeChannelId);
      const teamName = team?.displayName || routeTeamId;
      const channelName = channel?.displayName || routeChannelId;
      if (routeProjectKey === DEFAULT_ROUTE) {
        await invoke('saveConfig', { teamId: routeTeamId, channelId: routeChannelId, teamName, channelName, webhookUrl: routeWebhookUrl.trim(), fields });
        setSavedConfig({ teamId: routeTeamId, channelId: routeChannelId, teamName, channelName, webhookUrl: routeWebhookUrl.trim() });
        showToast(`Default channel set to ${channelName}.`);
      } else {
        await invoke('saveProjectConfig', { projectKey: routeProjectKey, teamId: routeTeamId, channelId: routeChannelId, teamName, channelName, webhookUrl: routeWebhookUrl.trim() });
        const project = mapProjects.find(p => p.key === routeProjectKey);
        showToast(`${project?.name || routeProjectKey} will now notify ${channelName}.`);
        loadProjectConfigs();
      }
      // Leaving the form as-is (not clearing it) so what you just saved stays visible right
      // where you set it, instead of vanishing and forcing you to go confirm it in the list
      // below. Use "Reset Form" when you're ready to add a different one.
    } catch (e) { showToast(formatError(e, 'Failed to save route.'), 'error'); }
    finally { setRouteSaving(false); }
  }

  async function removeRoute(projectKey) {
    try {
      await invoke('clearProjectConfig', { projectKey });
      showToast('Mapping removed — that project now uses the default channel.');
      loadProjectConfigs();
    } catch (e) { showToast(formatError(e, 'Failed to remove mapping.'), 'error'); }
  }

  async function clearDefault() {
    try {
      await invoke('clearConfig');
      setSavedConfig(null);
      resetRouteForm();
      showToast('Default channel cleared.');
    } catch (e) { showToast(formatError(e, 'Failed to clear default.'), 'error'); }
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

  const stepIndex = connected ? 1 : 0;
  const editingLabel = routeProjectKey === DEFAULT_ROUTE ? 'Default (all projects)' : (mapProjects.find(p => p.key === routeProjectKey)?.name || routeProjectKey);

  return (
    <div className="max-w-6xl mx-auto p-6 lg:p-8 font-sans">
      <PageHeader title="Microsoft Teams Connector" subtitle="Send Jira issue events to Teams via Microsoft Graph API" />

      <Stepper steps={['Connect', 'Channel Routing']} currentIndex={stepIndex} />

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

          {/* Step 2 — Channel Routing (default + per-project, one table) */}
          {connected && (
            <SectionCard title="Step 2 — Channel Routing" description="The default channel is where every project posts unless it has its own override below.">
              <p className="text-xs font-semibold text-jira-grey uppercase tracking-wide mb-2">
                {routeProjectKey === DEFAULT_ROUTE ? 'Editing: Default channel' : `Editing: ${editingLabel}`}
              </p>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4">
                <div>
                  <label className="label mt-0">Project</label>
                  <select className="form-select mb-3" value={routeProjectKey} onChange={e => setRouteProjectKey(e.target.value)}>
                    <option value={DEFAULT_ROUTE}>— Default (all projects) —</option>
                    {mapProjects.map(p => <option key={p.key} value={p.key}>{p.name} ({p.key})</option>)}
                  </select>
                </div>
                <div>
                  <label className="label mt-0">Team</label>
                  <select className="form-select mb-3" value={routeTeamId} onChange={onRouteTeamChange}>
                    <option value="">— choose a team —</option>
                    {teams.map(t => <option key={t.id} value={t.id}>{t.displayName}</option>)}
                  </select>
                </div>
              </div>

              {routeChannels.length > 0 && (
                <>
                  <label className="label mt-0">Channel</label>
                  <select className="form-select mb-3" value={routeChannelId} onChange={e => setRouteChannelId(e.target.value)}>
                    <option value="">— choose a channel —</option>
                    {routeChannels.map(c => <option key={c.id} value={c.id}>{c.displayName}</option>)}
                  </select>
                </>
              )}

              <label className="label">Teams Incoming Webhook URL for this channel</label>
              <input
                className="form-input mb-1 text-xs"
                type="text"
                placeholder="https://...api.powerplatform.com/powerautomate/..."
                value={routeWebhookUrl}
                onChange={e => setRouteWebhookUrl(e.target.value)}
              />
              <p className="text-xs text-jira-grey mb-4">Teams channel → ··· → Workflows → "Post to a channel when a webhook request is received" → Save → Copy URL</p>

              <div className="flex gap-2 mb-5">
                <Button className="flex-1" onClick={saveRoute} loading={routeSaving} disabled={!routeTeamId || !routeChannelId || !routeWebhookUrl.trim()}>
                  {routeProjectKey === DEFAULT_ROUTE ? 'Save Default Channel' : 'Add / Update Project Route'}
                </Button>
                <Button variant="ghost" onClick={resetRouteForm}>Reset Form</Button>
              </div>

              <label className="label mt-0">Current routing</label>
              <div className="space-y-2">
                <div className="flex items-center justify-between gap-3 bg-jira-light rounded px-3 py-2">
                  <div className="text-sm">
                    <span className="font-semibold text-jira-dark">Default (all other projects)</span>
                    {savedConfig ? <span className="text-jira-grey"> → {savedConfig.teamName} / {savedConfig.channelName}</span> : <span className="text-jira-grey"> — not set</span>}
                  </div>
                  <div className="flex gap-2 shrink-0">
                    <Button variant="ghost" className="text-xs" onClick={() => editRoute(DEFAULT_ROUTE)}>Edit</Button>
                    {savedConfig && <Button variant="danger" className="text-xs" onClick={clearDefault}>Clear</Button>}
                  </div>
                </div>
                {projectConfigs.map(({ projectKey, config }) => (
                  <div key={projectKey} className="flex items-center justify-between gap-3 bg-jira-light rounded px-3 py-2">
                    <div className="text-sm">
                      <span className="font-semibold text-jira-dark">{projectKey}</span>
                      <span className="text-jira-grey"> → {config.teamName} / {config.channelName}</span>
                    </div>
                    <div className="flex gap-2 shrink-0">
                      <Button variant="ghost" className="text-xs" onClick={() => editRoute(projectKey)}>Edit</Button>
                      <Button variant="danger" className="text-xs" onClick={() => removeRoute(projectKey)}>Remove</Button>
                    </div>
                  </div>
                ))}
              </div>

              {savedConfig && (
                <Button variant="success" className="w-full mt-4" onClick={test} loading={loading}>Send Test Message (Default Channel)</Button>
              )}
            </SectionCard>
          )}

          {savedConfig && (
            <SectionCard title="Fields & Notification Settings" description="Which fields show in a notification, and which Jira events trigger one — applies to the default channel.">
              <label className="label mt-0">Fields to include</label>
              <div className="card grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-2 mb-4">
                {NOTIFICATION_FIELDS.map(f => (
                  <label key={f.key} className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" className="w-4 h-4 accent-jira-blue" checked={fields.includes(f.key)} onChange={() => toggleField(f.key)} />
                    <span className="text-sm text-jira-dark">{f.label}</span>
                  </label>
                ))}
              </div>

              <label className="label mt-0">Notify on</label>
              <div className="card grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-3 mb-4">
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
                  try {
                    await Promise.all([
                      invoke('saveNotificationSettings', { settings: notifSettings }),
                      invoke('saveConfig', { ...savedConfig, fields }),
                    ]);
                    showToast('Settings saved.');
                  } catch (e) { showToast(formatError(e, 'Save failed.'), 'error'); }
                }}
              >
                Save Settings
              </Button>
            </SectionCard>
          )}
        </div>

        {/* ── Sidebar column — supplementary tools ── */}
        {connected && (
          <div>
            <SectionCard title="Send Custom Message" description="Posts to the default channel.">
              <textarea
                className="form-textarea w-full mb-2 h-20"
                placeholder="Type your message here..."
                value={customMsg}
                onChange={e => setCustomMsg(e.target.value)}
              />
              <Button className="w-full" onClick={sendCustomMessage} loading={loading} disabled={!customMsg.trim()}>Send to Teams</Button>
            </SectionCard>

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

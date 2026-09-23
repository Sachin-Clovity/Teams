import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { SectionCard, Button, useToast, formatError } from '../common';

// Sentinel used in the routing form's Project dropdown to mean "the global default,"
// as opposed to any specific project key.
const DEFAULT_ROUTE = '';

// ── Channel routing — one unified table. The global default and every project-specific
// override are really the same shape of data (team/channel/webhook); they used to be two
// separate sections with duplicated form code, which read as two different features when
// they're really one "where does this go" question answered at two different scopes.
//
// Self-contained: owns its own project list, saved routes, and form state. Only `teams`
// (loaded once by the parent's Connect step) and `savedConfig` (also read by the sibling
// Notification Settings section, so it stays lifted) come in as props. ──
export default function ChannelRouting({ teams, fields, savedConfig, onSavedConfigChange, loading, onTest }) {
  const showToast = useToast();

  const [mapProjects,    setMapProjects]    = useState([]); // every project on the site, for the picker
  const [projectConfigs, setProjectConfigs] = useState([]); // saved project-specific routes

  const [routeProjectKey, setRouteProjectKey] = useState(DEFAULT_ROUTE);
  const [routeTeamId,     setRouteTeamId]     = useState('');
  const [routeChannelId,  setRouteChannelId]  = useState('');
  const [routeChannels,   setRouteChannels]   = useState([]);
  const [routeWebhookUrl, setRouteWebhookUrl] = useState('');
  const [routeSaving,     setRouteSaving]     = useState(false);

  function loadProjectConfigs() {
    invoke('getAllProjectConfigs').then(({ configs }) => setProjectConfigs(configs || [])).catch(() => {});
  }

  useEffect(() => {
    invoke('searchProjects').then(({ projects }) => setMapProjects(projects || [])).catch(() => {});
    loadProjectConfigs();
  }, []);

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
        onSavedConfigChange({ teamId: routeTeamId, channelId: routeChannelId, teamName, channelName, webhookUrl: routeWebhookUrl.trim() });
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
      onSavedConfigChange(null);
      resetRouteForm();
      showToast('Default channel cleared.');
    } catch (e) { showToast(formatError(e, 'Failed to clear default.'), 'error'); }
  }

  const editingLabel = routeProjectKey === DEFAULT_ROUTE ? 'Default channel' : (mapProjects.find(p => p.key === routeProjectKey)?.name || routeProjectKey);

  return (
    <SectionCard title="Step 2 — Channel Routing" description="The default channel is where every project posts unless it has its own override below.">
      <p className="text-xs font-semibold text-jira-grey uppercase tracking-wide mb-2">Editing: {editingLabel}</p>

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
        <Button variant="success" className="w-full mt-4" onClick={onTest} loading={loading}>Send Test Message (Default Channel)</Button>
      )}
    </SectionCard>
  );
}

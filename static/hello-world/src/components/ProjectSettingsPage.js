import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';

const NOTIFICATION_FIELDS = [
  { key: 'status', label: 'Status' }, { key: 'priority', label: 'Priority' },
  { key: 'issueType', label: 'Type' }, { key: 'assignee', label: 'Assignee' },
  { key: 'reporter', label: 'Reporter' }, { key: 'dueDate', label: 'Due Date' },
  { key: 'labels', label: 'Labels' },
];

// ── Project Settings Page — per-project channel routing + notification filters ──
export default function ProjectSettingsPage({ projectKey }) {
  const [config,          setConfig]          = useState(null);
  const [teams,           setTeams]           = useState([]);
  const [channels,        setChannels]        = useState([]);
  const [teamId,          setTeamId]          = useState('');
  const [channelId,       setChannelId]       = useState('');
  const [webhookUrl,      setWebhookUrl]      = useState('');
  const [meta,            setMeta]            = useState({ issueTypes: [], statuses: [], priorities: [] });
  const [issueTypeFilter, setIssueTypeFilter] = useState([]);
  const [statusFilter,    setStatusFilter]    = useState([]);
  const [priorityFilter,  setPriorityFilter]  = useState([]);
  const [fields,          setFields]          = useState(['status', 'priority', 'issueType', 'assignee', 'reporter']);
  const [loading,         setLoading]         = useState(true);
  const [err,             setErr]             = useState('');
  const [msg,             setMsg]             = useState('');

  useEffect(() => {
    if (!projectKey) return;
    Promise.all([
      invoke('getProjectConfig', { projectKey }),
      invoke('getProjectMeta', { projectKey }),
    ]).then(([cfgRes, metaRes]) => {
      setMeta(metaRes);
      if (cfgRes.config) {
        const c = cfgRes.config;
        setConfig(c);
        setTeamId(c.teamId || ''); setChannelId(c.channelId || ''); setWebhookUrl(c.webhookUrl || '');
        setIssueTypeFilter(c.filters?.issueTypes || []);
        setStatusFilter(c.filters?.statuses || []);
        setPriorityFilter(c.filters?.priorities || []);
        if (c.fields?.length) setFields(c.fields);
      }
      setLoading(false);
    }).catch(e => { setErr(e.message || 'Failed to load.'); setLoading(false); });
  }, [projectKey]);

  async function loadTeams() {
    const res = await invoke('getTeams').catch(e => { setErr(e.message); return null; });
    if (res) setTeams(res.teams || []);
  }

  async function onTeamChange(id) {
    setTeamId(id); setChannelId(''); setChannels([]);
    if (!id) return;
    const res = await invoke('getChannels', { teamId: id }).catch(e => { setErr(e.message); return null; });
    if (res) setChannels(res.channels || []);
  }

  function toggleInList(list, setList, value) {
    setList(list.includes(value) ? list.filter(v => v !== value) : [...list, value]);
  }

  function toggleField(key) {
    setFields(prev => prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]);
  }

  const chipClass = active => `text-xs px-2 py-1 rounded border ${active ? 'border-jira-blue bg-jira-blue text-white' : 'border-jira-border text-jira-grey hover:border-jira-blue'}`;

  async function save() {
    setErr(''); setMsg('');
    try {
      const team    = teams.find(t => t.id === teamId);
      const channel = channels.find(c => c.id === channelId);
      await invoke('saveProjectConfig', {
        projectKey, teamId, channelId,
        teamName: team?.displayName || config?.teamName || '',
        channelName: channel?.displayName || config?.channelName || '',
        webhookUrl,
        filters: { issueTypes: issueTypeFilter, statuses: statusFilter, priorities: priorityFilter },
        fields,
      });
      setMsg('Project settings saved. This project now overrides the global channel.');
    } catch (e) { setErr(e.message || 'Save failed.'); }
  }

  async function clearAndUseGlobal() {
    await invoke('clearProjectConfig', { projectKey }).catch(() => {});
    setConfig(null); setTeamId(''); setChannelId(''); setWebhookUrl('');
    setIssueTypeFilter([]); setStatusFilter([]); setPriorityFilter([]);
    setMsg('Reverted — this project now uses the global default channel.');
  }

  if (loading) return <div className="p-4 text-jira-grey text-sm">Loading…</div>;

  return (
    <div className="max-w-xl mx-auto p-6 font-sans">
      <h1 className="text-lg font-bold text-jira-dark mb-1">Teams Connector — {projectKey} Settings</h1>
      <p className="text-xs text-jira-grey mb-4">Override the global Teams channel and notification rules for this project only. Leave unset to use the global default.</p>

      {err && <div className="alert-err">{err}</div>}
      {msg && <div className="alert-ok">{msg}</div>}

      <div className="section-title">Destination channel</div>
      {!teams.length && <button className="btn-blue w-full mb-3" onClick={loadTeams}>Load Teams</button>}
      {!!teams.length && (
        <>
          <label className="label">Team</label>
          <select className="form-select mb-3" value={teamId} onChange={e => onTeamChange(e.target.value)}>
            <option value="">— choose a team —</option>
            {teams.map(t => <option key={t.id} value={t.id}>{t.displayName}</option>)}
          </select>
        </>
      )}
      {!!channels.length && (
        <>
          <label className="label">Channel</label>
          <select className="form-select mb-3" value={channelId} onChange={e => setChannelId(e.target.value)}>
            <option value="">— choose a channel —</option>
            {channels.map(c => <option key={c.id} value={c.id}>{c.displayName}</option>)}
          </select>
        </>
      )}
      <label className="label">Webhook URL</label>
      <input className="form-input mb-4 text-xs" value={webhookUrl} onChange={e => setWebhookUrl(e.target.value)} placeholder="https://outlook.office.com/webhook/..." />

      <div className="divider" />
      <div className="section-title">Notify only for (leave empty = all)</div>

      <label className="label">Issue Types</label>
      <div className="flex flex-wrap gap-2 mb-3">
        {meta.issueTypes.map(t => (
          <button key={t} type="button" onClick={() => toggleInList(issueTypeFilter, setIssueTypeFilter, t)} className={chipClass(issueTypeFilter.includes(t))}>{t}</button>
        ))}
      </div>

      <label className="label">Statuses</label>
      <div className="flex flex-wrap gap-2 mb-3">
        {meta.statuses.map(s => (
          <button key={s} type="button" onClick={() => toggleInList(statusFilter, setStatusFilter, s)} className={chipClass(statusFilter.includes(s))}>{s}</button>
        ))}
      </div>

      <label className="label">Priorities</label>
      <div className="flex flex-wrap gap-2 mb-4">
        {meta.priorities.map(p => (
          <button key={p} type="button" onClick={() => toggleInList(priorityFilter, setPriorityFilter, p)} className={chipClass(priorityFilter.includes(p))}>{p}</button>
        ))}
      </div>

      <div className="divider" />
      <div className="section-title">Fields shown in notification</div>
      <div className="card space-y-2 mb-4">
        {NOTIFICATION_FIELDS.map(f => (
          <label key={f.key} className="flex items-center gap-2 cursor-pointer">
            <input type="checkbox" className="w-4 h-4 accent-jira-blue" checked={fields.includes(f.key)} onChange={() => toggleField(f.key)} />
            <span className="text-sm text-jira-dark">{f.label}</span>
          </label>
        ))}
      </div>

      <div className="flex gap-2">
        <button className="btn-blue flex-1" onClick={save}>Save Project Settings</button>
        <button className="btn-red" onClick={clearAndUseGlobal}>Use Global Default</button>
      </div>
    </div>
  );
}

import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { PageHeader, SectionCard, Button, useToast, PageSkeleton, formatError } from './common';

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
  const [meta,            setMeta]            = useState({ issueTypes: [], statuses: [], priorities: [], components: [] });
  const [issueTypeFilter, setIssueTypeFilter] = useState([]);
  const [statusFilter,    setStatusFilter]    = useState([]);
  const [priorityFilter,  setPriorityFilter]  = useState([]);
  const [labelFilterText, setLabelFilterText] = useState('');
  const [labelMatch,      setLabelMatch]      = useState('any');
  const [componentFilter, setComponentFilter] = useState([]);
  const [componentMatch,  setComponentMatch]  = useState('any');
  const [fields,          setFields]          = useState(['status', 'priority', 'issueType', 'assignee', 'reporter']);
  const [loading,         setLoading]         = useState(true);
  const [saving,          setSaving]          = useState(false);
  const [err,             setErr]             = useState('');
  const showToast = useToast();

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
        setLabelFilterText((c.filters?.labels || []).join(', '));
        setLabelMatch(c.filters?.labelMatch === 'all' ? 'all' : 'any');
        setComponentFilter(c.filters?.components || []);
        setComponentMatch(c.filters?.componentMatch === 'all' ? 'all' : 'any');
        if (c.fields?.length) setFields(c.fields);
      }
      setLoading(false);
    }).catch(e => { setErr(formatError(e, 'Failed to load.')); setLoading(false); });
  }, [projectKey]);

  async function loadTeams() {
    const res = await invoke('getTeams').catch(e => { setErr(formatError(e)); return null; });
    if (res) setTeams(res.teams || []);
  }

  async function onTeamChange(id) {
    setTeamId(id); setChannelId(''); setChannels([]);
    if (!id) return;
    const res = await invoke('getChannels', { teamId: id }).catch(e => { setErr(formatError(e)); return null; });
    if (res) setChannels(res.channels || []);
  }

  function toggleInList(list, setList, value) {
    setList(list.includes(value) ? list.filter(v => v !== value) : [...list, value]);
  }

  function toggleField(key) {
    setFields(prev => prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]);
  }

  const chipClass = active => `text-xs px-2 py-1 rounded border ${active ? 'border-jira-blue bg-jira-blue text-white' : 'border-jira-border text-jira-grey hover:border-jira-blue'}`;

  // Small ANY/ALL switch for multi-valued filters (labels, components) — an issue can carry
  // several of these at once, so unlike issue type/status/priority, "match" needs a strategy.
  function matchToggle(value, setValue) {
    return (
      <div className="flex text-[10px] font-semibold rounded overflow-hidden border border-jira-border">
        {['any', 'all'].map(opt => (
          <button
            key={opt} type="button" onClick={() => setValue(opt)}
            className={`px-2 py-0.5 uppercase ${value === opt ? 'bg-jira-blue text-white' : 'bg-white text-jira-grey'}`}
          >
            {opt}
          </button>
        ))}
      </div>
    );
  }

  async function save() {
    setErr(''); setSaving(true);
    try {
      const team    = teams.find(t => t.id === teamId);
      const channel = channels.find(c => c.id === channelId);
      await invoke('saveProjectConfig', {
        projectKey, teamId, channelId,
        teamName: team?.displayName || config?.teamName || '',
        channelName: channel?.displayName || config?.channelName || '',
        webhookUrl,
        filters: {
          issueTypes: issueTypeFilter, statuses: statusFilter, priorities: priorityFilter,
          labels: labelFilterText.split(',').map(s => s.trim()).filter(Boolean), labelMatch,
          components: componentFilter, componentMatch,
        },
        fields,
      });
      showToast('Project settings saved. This project now overrides the global channel.');
    } catch (e) { showToast(formatError(e, 'Save failed.'), 'error'); }
    finally { setSaving(false); }
  }

  async function clearAndUseGlobal() {
    setErr(''); setSaving(true);
    try {
      await invoke('clearProjectConfig', { projectKey });
      setConfig(null); setTeamId(''); setChannelId(''); setWebhookUrl('');
      setIssueTypeFilter([]); setStatusFilter([]); setPriorityFilter([]);
      setLabelFilterText(''); setLabelMatch('any'); setComponentFilter([]); setComponentMatch('any');
      showToast('Reverted — this project now uses the global default channel.');
    } catch (e) { showToast(formatError(e, 'Failed to revert.'), 'error'); }
    finally { setSaving(false); }
  }

  if (loading) return <PageSkeleton />;

  return (
    <div className="max-w-xl mx-auto p-6 font-sans">
      <PageHeader
        title={`Teams Connector — ${projectKey}`}
        subtitle="Override the global Teams channel and notification rules for this project only. Leave unset to use the global default."
      />

      {err && <div className="alert-err">{err}</div>}

      <SectionCard title="Destination channel">
        {!teams.length && <Button className="w-full mb-1" onClick={loadTeams}>Load Teams</Button>}
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
        <input className="form-input text-xs" value={webhookUrl} onChange={e => setWebhookUrl(e.target.value)} placeholder="https://outlook.office.com/webhook/..." />
      </SectionCard>

      <SectionCard title="Notify only for" description="Leave a group empty to match all values.">
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
        <div className="flex flex-wrap gap-2 mb-3">
          {meta.priorities.map(p => (
            <button key={p} type="button" onClick={() => toggleInList(priorityFilter, setPriorityFilter, p)} className={chipClass(priorityFilter.includes(p))}>{p}</button>
          ))}
        </div>

        <div className="flex items-center justify-between mt-3">
          <label className="label mt-0">Labels</label>
          {matchToggle(labelMatch, setLabelMatch)}
        </div>
        <input
          className="form-input text-xs mb-3"
          value={labelFilterText}
          onChange={e => setLabelFilterText(e.target.value)}
          placeholder="e.g. urgent, customer-facing (comma-separated)"
        />

        {!!meta.components.length && (
          <>
            <div className="flex items-center justify-between">
              <label className="label mt-0">Components</label>
              {matchToggle(componentMatch, setComponentMatch)}
            </div>
            <div className="flex flex-wrap gap-2">
              {meta.components.map(c => (
                <button key={c} type="button" onClick={() => toggleInList(componentFilter, setComponentFilter, c)} className={chipClass(componentFilter.includes(c))}>{c}</button>
              ))}
            </div>
          </>
        )}
      </SectionCard>

      <SectionCard title="Fields shown in notification">
        <div className="card space-y-2 mb-0">
          {NOTIFICATION_FIELDS.map(f => (
            <label key={f.key} className="flex items-center gap-2 cursor-pointer">
              <input type="checkbox" className="w-4 h-4 accent-jira-blue" checked={fields.includes(f.key)} onChange={() => toggleField(f.key)} />
              <span className="text-sm text-jira-dark">{f.label}</span>
            </label>
          ))}
        </div>
      </SectionCard>

      <div className="flex gap-2">
        <Button className="flex-1" onClick={save} loading={saving}>Save Project Settings</Button>
        <Button variant="danger" onClick={clearAndUseGlobal} loading={saving}>Use Global Default</Button>
      </div>
    </div>
  );
}

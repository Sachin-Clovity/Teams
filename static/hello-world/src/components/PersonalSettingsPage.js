import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { PageHeader, SectionCard, Button, useToast, PageSkeleton, formatError } from './common';

const DEFAULT_SETTINGS = {
  dmOnAssigned: true,
  dmOnStatusChange: false,
  dmOnReported: false,
  dmOnMentioned: true,
  dmOnWatching: false,
};

const TOGGLES = [
  { key: 'dmOnAssigned',     label: "DM me when I'm assigned an issue",       desc: 'Sent as a 1:1 Teams message when you become the assignee.' },
  { key: 'dmOnStatusChange', label: 'DM me when status changes on my issues', desc: 'Sent when an issue assigned to you changes status.' },
  { key: 'dmOnReported',     label: 'DM me when status changes on issues I reported', desc: "Sent when an issue you reported changes status, even if you're not the assignee." },
  { key: 'dmOnMentioned',    label: "DM me when I'm mentioned",               desc: 'Sent when someone @mentions you in a comment.' },
  { key: 'dmOnWatching',     label: 'DM me when status changes on issues I watch', desc: 'Sent when a status changes on an issue you’re watching.' },
];

// ── Personal Settings Page — per-user DM preferences ──────────────────────
export default function PersonalSettingsPage() {
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [loading,  setLoading]  = useState(true);
  const [saving,   setSaving]   = useState(false);
  const showToast = useToast();

  useEffect(() => {
    invoke('getPersonalConfig')
      .then(({ settings }) => { if (settings) setSettings(prev => ({ ...prev, ...settings })); setLoading(false); })
      .catch(() => setLoading(false));
  }, []);

  async function save() {
    setSaving(true);
    try { await invoke('savePersonalConfig', { settings }); showToast('Preferences saved.'); }
    catch (e) { showToast(formatError(e, 'Save failed.'), 'error'); }
    finally { setSaving(false); }
  }

  if (loading) return <PageSkeleton />;

  return (
    <div className="max-w-xl mx-auto p-6 font-sans">
      <PageHeader title="My Teams Notifications" subtitle="Choose when the Jira bot DMs you directly in Microsoft Teams." />

      <SectionCard>
        <div className="card space-y-3 mb-0">
          {TOGGLES.map(t => (
            <label key={t.key} className="flex items-start gap-3 cursor-pointer">
              <input type="checkbox" className="mt-0.5 w-4 h-4 accent-jira-blue" checked={!!settings[t.key]}
                onChange={e => { const checked = e.target.checked; setSettings(prev => ({ ...prev, [t.key]: checked })); }} />
              <div>
                <div className="text-sm font-medium text-jira-dark">{t.label}</div>
                <div className="text-xs text-jira-grey">{t.desc}</div>
              </div>
            </label>
          ))}
        </div>
      </SectionCard>

      <Button onClick={save} loading={saving}>Save</Button>
    </div>
  );
}

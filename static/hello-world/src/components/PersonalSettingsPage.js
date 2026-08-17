import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';

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
  const [saved,    setSaved]    = useState(false);

  useEffect(() => {
    invoke('getPersonalConfig')
      .then(({ settings }) => { if (settings) setSettings(prev => ({ ...prev, ...settings })); setLoading(false); })
      .catch(() => setLoading(false));
  }, []);

  async function save() {
    await invoke('savePersonalConfig', { settings }).catch(() => {});
    setSaved(true);
    setTimeout(() => setSaved(false), 2000);
  }

  if (loading) return <div className="p-4 text-jira-grey text-sm">Loading…</div>;

  return (
    <div className="max-w-xl mx-auto p-6 font-sans">
      <h1 className="text-lg font-bold text-jira-dark mb-1">My Teams Notifications</h1>
      <p className="text-xs text-jira-grey mb-4">Choose when the Jira bot DMs you directly in Microsoft Teams.</p>

      <div className="card space-y-3">
        {TOGGLES.map(t => (
          <label key={t.key} className="flex items-start gap-3 cursor-pointer">
            <input type="checkbox" className="mt-0.5 w-4 h-4 accent-jira-blue" checked={!!settings[t.key]}
              onChange={e => setSettings(prev => ({ ...prev, [t.key]: e.target.checked }))} />
            <div>
              <div className="text-sm font-medium text-jira-dark">{t.label}</div>
              <div className="text-xs text-jira-grey">{t.desc}</div>
            </div>
          </label>
        ))}
      </div>

      <div className="flex items-center gap-3 mt-3">
        <button className="btn-blue" onClick={save}>Save</button>
        {saved && <span className="text-jira-green text-xs font-semibold">Saved!</span>}
      </div>
    </div>
  );
}

import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';

// ── Personal Settings Page — per-user DM preferences ──────────────────────
export default function PersonalSettingsPage() {
  const [settings, setSettings] = useState({ dmOnAssigned: true, dmOnStatusChange: false });
  const [loading,  setLoading]  = useState(true);
  const [saved,    setSaved]    = useState(false);

  useEffect(() => {
    invoke('getPersonalConfig')
      .then(({ settings }) => { if (settings) setSettings(settings); setLoading(false); })
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
        <label className="flex items-start gap-3 cursor-pointer">
          <input type="checkbox" className="mt-0.5 w-4 h-4 accent-jira-blue" checked={settings.dmOnAssigned}
            onChange={e => setSettings(prev => ({ ...prev, dmOnAssigned: e.target.checked }))} />
          <div>
            <div className="text-sm font-medium text-jira-dark">DM me when I'm assigned an issue</div>
            <div className="text-xs text-jira-grey">Sent as a 1:1 Teams message when you become the assignee.</div>
          </div>
        </label>
        <label className="flex items-start gap-3 cursor-pointer">
          <input type="checkbox" className="mt-0.5 w-4 h-4 accent-jira-blue" checked={settings.dmOnStatusChange}
            onChange={e => setSettings(prev => ({ ...prev, dmOnStatusChange: e.target.checked }))} />
          <div>
            <div className="text-sm font-medium text-jira-dark">DM me when status changes on my issues</div>
            <div className="text-xs text-jira-grey">Sent when an issue assigned to you changes status.</div>
          </div>
        </label>
      </div>

      <div className="flex items-center gap-3 mt-3">
        <button className="btn-blue" onClick={save}>Save</button>
        {saved && <span className="text-jira-green text-xs font-semibold">Saved!</span>}
      </div>
    </div>
  );
}

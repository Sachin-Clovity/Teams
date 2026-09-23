import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { PageHeader, SectionCard, Stepper, Button, useToast, formatError } from '../common';
import ChannelRouting from './ChannelRouting';
import NotificationSettings from './NotificationSettings';
import DebugSidebar from './DebugSidebar';

// ── Admin / Global Settings Page ─────────────────────────────────────────────
// Orchestrator only — Step 1 (Connect) lives here since nothing else needs its own file for
// one button; Channel Routing, Notification Settings, and the sidebar tools are each their own
// component (see the sibling files in this folder), so this file stays a page layout instead
// of accumulating every feature's state and handlers in one place.
export default function AdminPage() {
  const [connected,   setConnected]   = useState(false);
  const [teams,        setTeams]       = useState([]);
  const [loading,      setLoading]     = useState(false);
  const [err,          setErr]         = useState('');
  const [savedConfig,  setSavedConfig] = useState(null);
  const [notifSettings, setNotifSettings] = useState({ created: true, updated: true, deleted: false, commented: false });
  const [fields,        setFields]        = useState(['status', 'priority', 'issueType', 'assignee', 'reporter']);

  const showToast = useToast();

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
    invoke('getNotificationSettings').then(({ settings }) => { if (settings) setNotifSettings(settings); }).catch(() => {});
  }, []);

  function toggleField(key) {
    setFields(prev => prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]);
  }

  async function connect() {
    setErr(''); setLoading(true);
    try {
      const res = await invoke('getTeams');
      if (!res.teams?.length) { setErr('Connected but no Teams found. Check Azure Application permissions.'); return; }
      setTeams(res.teams);
      setConnected(true);
      showToast('Connected. Set up channel routing below.');
    } catch (e) { setErr(formatError(e, 'Connection failed. Check Forge logs.')); }
    finally { setLoading(false); }
  }

  async function test() {
    setErr(''); setLoading(true);
    try {
      const res = await invoke('testConnection');
      res.success ? showToast('Test message sent. Check your Teams channel.') : showToast(formatError({ message: res.error }, 'Test failed.'), 'error');
    } catch (e) { showToast(formatError(e, 'Test failed.'), 'error'); }
    finally { setLoading(false); }
  }

  const stepIndex = connected ? 1 : 0;

  return (
    <div className="max-w-6xl mx-auto p-6 lg:p-8 font-sans">
      <PageHeader title="Microsoft Teams Connector" subtitle="Send Jira issue events to Teams via Microsoft Graph API" />

      <Stepper steps={['Connect', 'Channel Routing']} currentIndex={stepIndex} />

      {err && <div className="alert-err">{err}</div>}

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-start">
        <div className="lg:col-span-2">
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

          {connected && (
            <ChannelRouting
              teams={teams}
              fields={fields}
              savedConfig={savedConfig}
              onSavedConfigChange={setSavedConfig}
              loading={loading}
              onTest={test}
            />
          )}

          {savedConfig && (
            <NotificationSettings
              fields={fields}
              onToggleField={toggleField}
              savedConfig={savedConfig}
              notifSettings={notifSettings}
              onNotifSettingsChange={setNotifSettings}
            />
          )}
        </div>

        {connected && <DebugSidebar />}
      </div>
    </div>
  );
}

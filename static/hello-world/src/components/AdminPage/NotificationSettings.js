import React, { useState } from 'react';
import { invoke } from '@forge/bridge';
import { SectionCard, Button, useToast, formatError } from '../common';

const NOTIFICATION_FIELDS = [
  { key: 'status', label: 'Status' }, { key: 'priority', label: 'Priority' },
  { key: 'issueType', label: 'Type' }, { key: 'assignee', label: 'Assignee' },
  { key: 'reporter', label: 'Reporter' }, { key: 'dueDate', label: 'Due Date' },
  { key: 'labels', label: 'Labels' },
];

const EVENT_TOGGLES = [
  { key: 'created',   label: 'Issue Created',   desc: 'Notify when a new issue is created' },
  { key: 'updated',   label: 'Issue Updated',   desc: 'Notify when an issue is modified' },
  { key: 'deleted',   label: 'Issue Deleted',   desc: 'Notify when an issue is deleted' },
  { key: 'commented', label: 'Comment Added',   desc: 'Notify when a comment is added' },
];

// Which fields show in a notification, and which Jira events trigger one — both apply to the
// default channel. `fields` is lifted to the parent because saving the Default route (in
// ChannelRouting) persists the current field selection alongside it.
export default function NotificationSettings({ fields, onToggleField, savedConfig, notifSettings, onNotifSettingsChange }) {
  const showToast = useToast();
  const [saving, setSaving] = useState(false);

  async function save() {
    setSaving(true);
    try {
      await Promise.all([
        invoke('saveNotificationSettings', { settings: notifSettings }),
        invoke('saveConfig', { ...savedConfig, fields }),
      ]);
      showToast('Settings saved.');
    } catch (e) { showToast(formatError(e, 'Save failed.'), 'error'); }
    finally { setSaving(false); }
  }

  return (
    <SectionCard title="Fields & Notification Settings" description="Which fields show in a notification, and which Jira events trigger one — applies to the default channel.">
      <label className="label mt-0">Fields to include</label>
      <div className="card grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-2 mb-4">
        {NOTIFICATION_FIELDS.map(f => (
          <label key={f.key} className="flex items-center gap-2 cursor-pointer">
            <input type="checkbox" className="w-4 h-4 accent-jira-blue" checked={fields.includes(f.key)} onChange={() => onToggleField(f.key)} />
            <span className="text-sm text-jira-dark">{f.label}</span>
          </label>
        ))}
      </div>

      <label className="label mt-0">Notify on</label>
      <div className="card grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-3 mb-4">
        {EVENT_TOGGLES.map(({ key, label, desc }) => (
          <label key={key} className="flex items-start gap-3 cursor-pointer group">
            <input
              type="checkbox"
              className="mt-0.5 w-4 h-4 accent-jira-blue"
              checked={!!notifSettings[key]}
              onChange={e => { const checked = e.target.checked; onNotifSettingsChange(prev => ({ ...prev, [key]: checked })); }}
            />
            <div>
              <div className="text-sm font-medium text-jira-dark">{label}</div>
              <div className="text-xs text-jira-grey">{desc}</div>
            </div>
          </label>
        ))}
      </div>
      <Button onClick={save} loading={saving}>Save Settings</Button>
    </SectionCard>
  );
}

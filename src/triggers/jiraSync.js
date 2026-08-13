import { getGlobalConfig, getNotificationSettings, getProjectConfig, getPersonalConfig } from '../storage/kvsStore.js';
import { DEFAULT_FIELDS, FIELD_DEFS } from '../config/constants.js';
import { getJiraUserEmail } from '../jira/utils.js';
import { sendPersonalDM } from '../graph/chat.js';

// Channel notification — respects per-config field customization.
async function postChannelNotification(config, eventType, issue, fields) {
  const actionMap = {
    'jira:issueCreated': '🆕 Issue Created',
    'jira:issueUpdated': '✏️ Issue Updated',
    'jira:issueDeleted': '🗑️ Issue Deleted',
  };
  const action   = actionMap[eventType] || '🔔 Issue Event';
  const issueKey = issue.key || 'Unknown';
  const summary  = fields.summary || '(no summary)';
  const jiraUrl  = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
  const color    = eventType === 'jira:issueCreated' ? '0078D4' : eventType === 'jira:issueDeleted' ? 'D83B01' : 'FFB900';

  const fieldKeys = config.fields?.length ? config.fields : DEFAULT_FIELDS;
  const factLine  = fieldKeys.filter(k => FIELD_DEFS[k]).map(k => `${FIELD_DEFS[k].label}: ${FIELD_DEFS[k].get(fields)}`).join(' | ');

  // Power Automate (Teams Workflows) expects plain text field
  const payload = {
    text: `${action}: **${issueKey}** — ${summary}\n${factLine}\n🔗 ${jiraUrl}`,
    // Legacy MessageCard fields (for old incoming webhooks)
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    "themeColor": color,
    "summary": `${action}: ${issueKey}`,
  };

  const { fetch } = await import('@forge/api');
  const res = await fetch(config.webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  console.log('[jiraSync] webhook post status:', res.status, await res.text());
}

// Personal DMs — assignment and status-change, gated by each user's own preference.
// Relies on event.changelog (present on jira:issueUpdated) to detect an actual assignee/status change,
// rather than firing on every unrelated update to an already-assigned issue.
async function maybeSendPersonalNotifications(eventType, issue, fields, changelogItems) {
  const assignee = fields.assignee;
  if (!assignee?.accountId) return;

  const personalCfg = await getPersonalConfig(assignee.accountId);
  const issueKey = issue.key || 'Unknown';
  const jiraUrl  = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;

  const assigneeChanged = changelogItems.some(i => i.field === 'assignee');
  const statusChanged    = changelogItems.some(i => i.field === 'status');
  const isNewAssignment  = eventType === 'jira:issueCreated' || assigneeChanged;

  if (isNewAssignment && personalCfg.dmOnAssigned) {
    const email = await getJiraUserEmail(assignee.accountId);
    if (email) await sendPersonalDM(email, `<b>You were assigned:</b> <a href="${jiraUrl}">${issueKey}</a> — ${fields.summary || ''}`);
  }
  if (statusChanged && personalCfg.dmOnStatusChange) {
    const email = await getJiraUserEmail(assignee.accountId);
    if (email) await sendPersonalDM(email, `<b>Status changed:</b> <a href="${jiraUrl}">${issueKey}</a> is now <b>${fields.status?.name || 'Unknown'}</b>`);
  }
}

// jira:issueCreated / jira:issueUpdated / jira:issueDeleted trigger.
export async function jiraSync(event) {
  console.log('[jiraSync] event received:', event.eventType);

  const eventType  = event.eventType || 'jira:issueUpdated';
  const issue      = event.issue || {};
  const fields     = issue.fields || {};
  const projectKey = fields.project?.key;

  const notifSettings = await getNotificationSettings();
  if (eventType === 'jira:issueCreated' && !notifSettings.created) { console.log('[jiraSync] created notifications disabled'); return; }
  if (eventType === 'jira:issueUpdated' && !notifSettings.updated) { console.log('[jiraSync] updated notifications disabled'); return; }
  if (eventType === 'jira:issueDeleted' && !notifSettings.deleted) { console.log('[jiraSync] deleted notifications disabled'); return; }

  const projectConfig = projectKey ? await getProjectConfig(projectKey) : null;
  const globalConfig  = await getGlobalConfig();
  const routeConfig   = projectConfig?.webhookUrl ? projectConfig : globalConfig;

  if (!routeConfig?.webhookUrl) {
    console.log('[jiraSync] no channel configured for', projectKey || 'this site', '— skipping channel post');
  } else {
    const filters = projectConfig?.filters;
    const passes  = (list, value) => !list?.length || list.includes(value);
    const filteredOut = filters && (
      !passes(filters.issueTypes, fields.issuetype?.name || 'Issue') ||
      !passes(filters.statuses,   fields.status?.name   || 'Unknown') ||
      !passes(filters.priorities, fields.priority?.name || 'None')
    );
    if (filteredOut) {
      console.log('[jiraSync] filtered out by project notification filters');
    } else {
      await postChannelNotification(routeConfig, eventType, issue, fields);
    }
  }

  await maybeSendPersonalNotifications(eventType, issue, fields, event.changelog?.items || []);
}

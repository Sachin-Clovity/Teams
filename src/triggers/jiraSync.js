import { getGlobalConfig, getNotificationSettings, getProjectConfig, getPersonalConfig, getIssueNotifySub } from '../storage/kvsStore.js';
import { DEFAULT_FIELDS, FIELD_DEFS } from '../config/constants.js';
import { getJiraUserEmail, getIssueWatchers, extractMentionedAccountIds } from '../jira/utils.js';
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

async function dmAccountId(accountId, html) {
  const email = await getJiraUserEmail(accountId);
  if (email) await sendPersonalDM(email, html);
}

// Personal DMs — assigned / status-change / reported / mentioned / watching, each gated by
// that specific person's own preference. Relies on event.changelog to detect an actual field
// change (not just any update), and event.comment for mention scanning where Forge provides it.
async function maybeSendPersonalNotifications(eventType, issue, fields, event) {
  const issueKey = issue.key || 'Unknown';
  const jiraUrl  = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
  const summary  = fields.summary || '';
  const statusName = fields.status?.name || 'Unknown';

  const changelogItems  = event.changelog?.items || [];
  const assigneeChanged = changelogItems.some(i => i.field === 'assignee');
  const statusChanged   = changelogItems.some(i => i.field === 'status');
  const isNewAssignment = eventType === 'jira:issueCreated' || assigneeChanged;

  const assignee = fields.assignee;
  const reporter = fields.reporter;

  if (assignee?.accountId) {
    const cfg = await getPersonalConfig(assignee.accountId);
    if (isNewAssignment && cfg.dmOnAssigned) {
      await dmAccountId(assignee.accountId, `<b>You were assigned:</b> <a href="${jiraUrl}">${issueKey}</a> — ${summary}`);
    }
    if (statusChanged && cfg.dmOnStatusChange) {
      await dmAccountId(assignee.accountId, `<b>Status changed:</b> <a href="${jiraUrl}">${issueKey}</a> is now <b>${statusName}</b>`);
    }
  }

  // Reporter — notified on status change, skipped if they're also the assignee (already covered above)
  if (reporter?.accountId && reporter.accountId !== assignee?.accountId && statusChanged) {
    const cfg = await getPersonalConfig(reporter.accountId);
    if (cfg.dmOnReported) {
      await dmAccountId(reporter.accountId, `<b>Status changed on your issue:</b> <a href="${jiraUrl}">${issueKey}</a> is now <b>${statusName}</b>`);
    }
  }

  // Mentions — only when this update actually carried a comment (Forge includes it on the event)
  if (event.comment?.body) {
    const mentioned = extractMentionedAccountIds(event.comment.body);
    for (const accountId of mentioned) {
      const cfg = await getPersonalConfig(accountId);
      if (cfg.dmOnMentioned) {
        await dmAccountId(accountId, `<b>You were mentioned on:</b> <a href="${jiraUrl}">${issueKey}</a> — ${summary}`);
      }
    }
  }

  // Watchers — notified on status change, skipping whoever was already notified above
  if (statusChanged) {
    try {
      const watcherIds = await getIssueWatchers(issueKey);
      for (const accountId of watcherIds) {
        if (accountId === assignee?.accountId || accountId === reporter?.accountId) continue;
        const cfg = await getPersonalConfig(accountId);
        if (cfg.dmOnWatching) {
          await dmAccountId(accountId, `<b>Status changed on a watched issue:</b> <a href="${jiraUrl}">${issueKey}</a> is now <b>${statusName}</b>`);
        }
      }
    } catch (e) {
      console.log('[jiraSync] watcher lookup failed:', e.message);
    }
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

  // Per-issue subscription (the "Notify" card button) can force a channel post through even
  // when project-level filters would otherwise exclude it — the user explicitly asked for this issue.
  const issueSub = await getIssueNotifySub(issue.key);
  const commentAdded = !!event.comment;
  const subForcesNotify = !!issueSub && (
    (issueSub.notifyOnUpdate && eventType === 'jira:issueUpdated') ||
    (issueSub.notifyOnComment && commentAdded)
  );

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
    if (filteredOut && !subForcesNotify) {
      console.log('[jiraSync] filtered out by project notification filters');
    } else {
      await postChannelNotification(routeConfig, eventType, issue, fields);
    }
  }

  await maybeSendPersonalNotifications(eventType, issue, fields, event);
}

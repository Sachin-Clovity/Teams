import { getGlobalConfig, getNotificationSettings, getProjectConfig, getPersonalConfig, getIssueNotifySub, getMsTenantId } from '../storage/kvsStore.js';
import { DEFAULT_FIELDS, FIELD_DEFS } from '../config/constants.js';
import { getJiraUserEmail, getIssueWatchers, extractMentionedAccountIds } from '../jira/utils.js';
import { sendPersonalDM } from '../graph/chat.js';
import { fetchWithRetry } from '../utils/http.js';

// Current Forge event identifiers (ARI-style) — the older dot-notation names
// (jira:issueCreated/Updated/Deleted) are what this trigger used to be registered with, and
// silently never fired a single time under that registration; these are what manifest.yml's
// trigger.events must also list for the platform to actually deliver anything here.
const CREATED   = 'avi:jira:created:issue';
const UPDATED   = 'avi:jira:updated:issue';
const DELETED   = 'avi:jira:deleted:issue';
// A distinct event from UPDATED — adding/editing a comment does not fire avi:jira:updated:issue
// on its own, so without registering this separately, comments never reached this trigger at
// all (independent of the "Comment Added" toggle, which had nothing to gate — it was checked
// nowhere in this file until now).
const COMMENTED = 'avi:jira:commented:issue';

// Channel notification — respects per-config field customization.
async function postChannelNotification(config, eventType, issue, fields) {
  const actionMap = {
    [CREATED]:   '🆕 Issue Created',
    [UPDATED]:   '✏️ Issue Updated',
    [DELETED]:   '🗑️ Issue Deleted',
    [COMMENTED]: '💬 Comment Added',
  };
  const action   = actionMap[eventType] || '🔔 Issue Event';
  const issueKey = issue.key || 'Unknown';
  const summary  = fields.summary || '(no summary)';
  const jiraUrl  = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
  const color    = eventType === CREATED ? '0078D4' : eventType === DELETED ? 'D83B01' : 'FFB900';

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
  const res = await fetchWithRetry(fetch, config.webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }, `webhook post for ${issueKey}`);
  console.log('[jiraSync] webhook post status:', res.status, await res.text());
}

// True if the issue has a Security Level set — those exist specifically to restrict who can
// see the issue beyond normal project permissions, so broadcasting it to an entire Teams
// channel (whose membership Jira has no visibility into) could show restricted data to people
// who could never have opened the issue in Jira itself. This only gates the CHANNEL post —
// personal DMs already go only to the assignee/reporter/watchers, who by definition already
// have access to the issue, so they're not the same risk.
//
// A deleted issue can't be re-fetched, so a delete event trusts whatever security field the
// trigger payload already carried. For created/updated events this re-fetches fields=security
// directly rather than trusting the trigger payload, since Forge doesn't guarantee every field
// is present on every event and getting this wrong means a real data leak, not a UX bug.
async function isSecurityRestricted(eventType, issue, fields) {
  if (eventType === DELETED) return !!fields.security;
  try {
    const { default: api, route } = await import('@forge/api');
    const res = await api.asApp().requestJira(route`/rest/api/3/issue/${issue.key}?fields=security`);
    if (!res.ok) return !!fields.security;
    const data = await res.json();
    return !!data.fields?.security;
  } catch (e) {
    console.log('[jiraSync] security-level check failed, treating as restricted to be safe:', e.message);
    return true;
  }
}

async function dmAccountId(accountId, html) {
  const email = await getJiraUserEmail(accountId);
  if (!email) return;
  const tenantId = await getMsTenantId();
  await sendPersonalDM(email, html, tenantId);
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
  const isNewAssignment = eventType === CREATED || assigneeChanged;

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

// avi:jira:created:issue / avi:jira:updated:issue / avi:jira:deleted:issue / avi:jira:commented:issue trigger.
export async function jiraSync(event) {
  console.log('[jiraSync] event received:', event.eventType);

  const eventType  = event.eventType || UPDATED;
  const issue      = event.issue || {};
  const fields     = issue.fields || {};
  const projectKey = fields.project?.key;

  const notifSettings = await getNotificationSettings();
  if (eventType === CREATED   && !notifSettings.created)   { console.log('[jiraSync] created notifications disabled');   return; }
  if (eventType === UPDATED   && !notifSettings.updated)   { console.log('[jiraSync] updated notifications disabled');   return; }
  if (eventType === DELETED   && !notifSettings.deleted)   { console.log('[jiraSync] deleted notifications disabled');   return; }
  if (eventType === COMMENTED && !notifSettings.commented) { console.log('[jiraSync] commented notifications disabled'); return; }

  const projectConfig = projectKey ? await getProjectConfig(projectKey) : null;
  const globalConfig  = await getGlobalConfig();
  const routeConfig   = projectConfig?.webhookUrl ? projectConfig : globalConfig;

  // Per-issue subscription (the "Notify" card button) can force a channel post through even
  // when project-level filters would otherwise exclude it — the user explicitly asked for this issue.
  const issueSub = await getIssueNotifySub(issue.key);
  const commentAdded = !!event.comment || eventType === COMMENTED;
  const subForcesNotify = !!issueSub && (
    (issueSub.notifyOnUpdate && eventType === UPDATED) ||
    (issueSub.notifyOnComment && commentAdded)
  );

  if (!routeConfig?.webhookUrl) {
    console.log('[jiraSync] no channel configured for', projectKey || 'this site', '— skipping channel post');
  } else if (await isSecurityRestricted(eventType, issue, fields)) {
    // A hard boundary — overrides even an explicit per-issue "Notify" subscription, since
    // that subscription reflects "notify about this issue," not "ignore its access controls."
    console.log('[jiraSync] issue has a security level set — skipping channel notification to avoid exposing restricted data');
  } else {
    const filters = projectConfig?.filters;
    // Single-valued fields (an issue has exactly one type/status/priority): passes if the
    // issue's value is in the selected list, or the list is empty (no filter set).
    const passes  = (list, value) => !list?.length || list.includes(value);
    // Multi-valued fields (an issue can have several labels/components): "any" passes if at
    // least one selected value is present on the issue; "all" requires every selected value
    // to be present.
    const passesMulti = (list, match, issueValues) => {
      if (!list?.length) return true;
      return match === 'all'
        ? list.every(v => issueValues.includes(v))
        : list.some(v => issueValues.includes(v));
    };
    const issueLabels     = fields.labels || [];
    const issueComponents = (fields.components || []).map(c => c.name);
    const filteredOut = filters && (
      !passes(filters.issueTypes, fields.issuetype?.name || 'Issue') ||
      !passes(filters.statuses,   fields.status?.name   || 'Unknown') ||
      !passes(filters.priorities, fields.priority?.name || 'None') ||
      !passesMulti(filters.labels,     filters.labelMatch,     issueLabels) ||
      !passesMulti(filters.components, filters.componentMatch, issueComponents)
    );
    if (filteredOut && !subForcesNotify) {
      console.log('[jiraSync] filtered out by project notification filters');
    } else {
      await postChannelNotification(routeConfig, eventType, issue, fields);
    }
  }

  await maybeSendPersonalNotifications(eventType, issue, fields, event);
}

import { getGlobalConfig, getMsTenantId } from '../storage/kvsStore.js';
import { getAppToken } from '../graph/auth.js';

// Jira Automation rule calls this HTTP endpoint → posts to Teams. Unlike the native
// jira:issueCreated/Updated/Deleted trigger (which only fires for the site this Forge app is
// actually installed on), this webtrigger is a plain HTTPS URL any Jira site's Automation
// rules can POST to — no install required. That also means it can't assume a single
// JIRA_BASE_URL: a second site using this path needs its OWN issue links, so the caller
// should pass the ready-made link (Jira Automation's {{baseUrl}}/browse/{{issue.key}} smart
// values) as `issueUrl` rather than relying on this function to guess which site it's from.
export async function automationHandler(req) {
  try {
    const body = req.body ? JSON.parse(req.body) : {};
    const config = await getGlobalConfig();
    if (!config?.teamId || !config?.channelId) {
      return { statusCode: 400, body: JSON.stringify({ error: 'No channel configured' }) };
    }

    const tenantId = await getMsTenantId();
    const token = await getAppToken(tenantId);
    const { fetch } = await import('@forge/api');

    const title    = body.title   || body.summary  || 'Automation Rule Triggered';
    const detail   = body.detail  || body.message  || '';
    const issueKey = body.issueKey || body.issue?.key || '';
    // Prefer a fully-formed link from the caller (correct for whichever site sent it) —
    // JIRA_BASE_URL is only a fallback for older rules that don't pass one.
    const jiraUrl  = body.issueUrl || (issueKey ? `${process.env.JIRA_BASE_URL}/browse/${issueKey}` : process.env.JIRA_BASE_URL);

    const content = issueKey
      ? `<b>🤖 ${title}: <a href="${jiraUrl}">${issueKey}</a></b>${detail ? `<br/>${detail}` : ''}`
      : `<b>🤖 ${title}</b>${detail ? `<br/>${detail}` : ''}`;

    const res = await fetch(
      `https://graph.microsoft.com/v1.0/teams/${config.teamId}/channels/${config.channelId}/messages`,
      {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ body: { contentType: 'html', content } }),
      }
    );
    console.log('[automationHandler] post status:', res.status);
    return { statusCode: 200, body: JSON.stringify({ success: true }) };
  } catch (e) {
    console.log('[automationHandler] error:', e.message);
    return { statusCode: 500, body: JSON.stringify({ error: e.message }) };
  }
}

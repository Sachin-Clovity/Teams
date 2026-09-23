import { getGlobalConfig, getMsTenantId } from '../storage/kvsStore.js';
import { graphPost } from '../graph/client.js';

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

    const title    = body.title   || body.summary  || 'Automation Rule Triggered';
    const detail   = body.detail  || body.message  || '';
    const issueKey = body.issueKey || body.issue?.key || '';
    // Prefer a fully-formed link from the caller (correct for whichever site sent it) —
    // JIRA_BASE_URL is only a fallback for older rules that don't pass one.
    const jiraUrl  = body.issueUrl || (issueKey ? `${process.env.JIRA_BASE_URL}/browse/${issueKey}` : process.env.JIRA_BASE_URL);

    const content = issueKey
      ? `<b>🤖 ${title}: <a href="${jiraUrl}">${issueKey}</a></b>${detail ? `<br/>${detail}` : ''}`
      : `<b>🤖 ${title}</b>${detail ? `<br/>${detail}` : ''}`;

    // graphPost retries on 429/5xx and throws on a real failure — caught below, so a rule
    // that fires many times in a row (e.g. behind a bulk import) no longer silently drops
    // messages on the first rate-limit response, and a genuine failure is now reported back
    // to the calling Automation rule instead of always claiming success.
    await graphPost(`/v1.0/teams/${config.teamId}/channels/${config.channelId}/messages`, {
      body: { contentType: 'html', content },
    }, tenantId);
    return { statusCode: 200, body: JSON.stringify({ success: true }) };
  } catch (e) {
    console.log('[automationHandler] error:', e.message);
    return { statusCode: 500, body: JSON.stringify({ error: e.message }) };
  }
}

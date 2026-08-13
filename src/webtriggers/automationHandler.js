import { getGlobalConfig } from '../storage/kvsStore.js';
import { getAppToken } from '../graph/auth.js';

// Jira Automation rule calls this HTTP endpoint → posts to Teams.
export async function automationHandler(req) {
  try {
    const body = req.body ? JSON.parse(req.body) : {};
    const config = await getGlobalConfig();
    if (!config?.teamId || !config?.channelId) {
      return { statusCode: 400, body: JSON.stringify({ error: 'No channel configured' }) };
    }

    const token = await getAppToken();
    const { fetch } = await import('@forge/api');

    const title    = body.title   || body.summary  || 'Automation Rule Triggered';
    const detail   = body.detail  || body.message  || '';
    const issueKey = body.issueKey || body.issue?.key || '';
    const jiraUrl  = issueKey ? `${process.env.JIRA_BASE_URL}/browse/${issueKey}` : process.env.JIRA_BASE_URL;

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

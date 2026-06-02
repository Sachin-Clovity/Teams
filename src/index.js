import * as ResolverPkg from '@forge/resolver';
import { storage } from '@forge/api';
// import api, { route, storage } from '@forge/api'; // providers.auth — commented

const Resolver = ResolverPkg.default ?? ResolverPkg;

// ── Client credentials — app-level token, no user OAuth needed ─────────────
async function getAppToken() {
  const { fetch } = await import('@forge/api');
  const body = [
    `client_id=${encodeURIComponent(process.env.MS_CLIENT_ID)}`,
    `client_secret=${encodeURIComponent(process.env.MS_CLIENT_SECRET)}`,
    `scope=${encodeURIComponent('https://graph.microsoft.com/.default')}`,
    `grant_type=client_credentials`,
  ].join('&');

  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`,
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error(`Token error: ${JSON.stringify(data)}`);
  return data.access_token;
}

async function graphGet(path, token) {
  const { fetch } = await import('@forge/api');
  const res = await fetch(`https://graph.microsoft.com${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  return res.json();
}

// ── providers.auth — COMMENTED ─────────────────────────────────────────────
// async function getMicrosoftClient() {
//   const client = api.asUser().withProvider('microsoft');
//   let account;
//   try { account = await client.getAccount(); } catch (e) { return null; }
//   if (!account) { await client.requestCredentials(); return null; }
//   return client;
// }

const resolver = new Resolver();

resolver.define('getTeams', async () => {
  console.log('[getTeams] called');
  const token = await getAppToken();
  const data = await graphGet(`/v1.0/teams`, token);
  const teams = (data.value || []).map(t => ({ id: t.id, displayName: t.displayName }));
  console.log('[getTeams] count:', teams.length);
  return { teams };
});

resolver.define('getChannels', async ({ payload }) => {
  console.log('[getChannels] teamId:', payload.teamId);
  const token = await getAppToken();
  const data = await graphGet(`/v1.0/teams/${payload.teamId}/channels`, token);
  const channels = (data.value || []).map(c => ({ id: c.id, displayName: c.displayName }));
  return { channels };
});

resolver.define('saveConfig', async ({ payload }) => {
  const { teamId, channelId, teamName, channelName } = payload;
  if (!teamId || !channelId) throw new Error('teamId and channelId are required');
  await storage.set('teams-channel-config', { teamId, channelId, teamName: teamName || '', channelName: channelName || '' });
  return { success: true };
});

resolver.define('getConfig', async () => {
  const config = await storage.get('teams-channel-config');
  return { config: config || null };
});

resolver.define('clearConfig', async () => {
  await storage.delete('teams-channel-config');
  return { success: true };
});

export const handler = resolver.getDefinitions();

// ── jiraSync trigger (active) ───────────────────────────────────────────────
export async function jiraSync(event) {
  console.log('[jiraSync] event received:', event.eventType);

  const config = await storage.get('teams-channel-config');
  if (!config?.teamId || !config?.channelId) {
    console.log('[jiraSync] no channel config saved — skipping');
    return;
  }

  let token;
  try {
    token = await getAppToken();
  } catch (e) {
    console.log('[jiraSync] token error:', e.message);
    return;
  }

  const issue = event.issue || {};
  const fields = issue.fields || {};
  const eventType = event.eventType || 'jira:issueUpdated';

  const actionMap = {
    'jira:issueCreated': '🆕 Issue Created',
    'jira:issueUpdated': '✏️ Issue Updated',
    'jira:issueDeleted': '🗑️ Issue Deleted',
  };
  const action    = actionMap[eventType] || '🔔 Issue Event';
  const issueKey  = issue.key || 'Unknown';
  const summary   = fields.summary || '(no summary)';
  const priority  = fields.priority?.name || 'None';
  const status    = fields.status?.name || 'Unknown';
  const reporter  = fields.reporter?.displayName || 'Unknown';
  const assignee  = fields.assignee?.displayName || 'Unassigned';
  const issueType = fields.issuetype?.name || 'Issue';
  const jiraUrl   = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;

  const { fetch } = await import('@forge/api');
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/teams/${config.teamId}/channels/${config.channelId}/messages`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        body: {
          contentType: 'html',
          content: `<b>${action}: <a href="${jiraUrl}">${issueKey}</a></b><br/>
${summary}<br/><br/>
<b>Type:</b> ${issueType} &nbsp; <b>Status:</b> ${status} &nbsp; <b>Priority:</b> ${priority}<br/>
<b>Reporter:</b> ${reporter} &nbsp; <b>Assignee:</b> ${assignee}`,
        },
      }),
    }
  );
  console.log('[jiraSync] post status:', res.status);
}

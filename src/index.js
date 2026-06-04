import * as ResolverPkg from '@forge/resolver';
import api, { route, storage } from '@forge/api';
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

// ── Bot Framework token — for sending replies back to Teams ───────────────
async function getBotFrameworkToken() {
  console.log('[getBotToken] fetching token for tenant:', process.env.MS_TENANT_ID);
  console.log('[getBotToken] client_id present:', !!process.env.MS_CLIENT_ID);
  console.log('[getBotToken] client_secret present:', !!process.env.MS_CLIENT_SECRET);
  const { fetch } = await import('@forge/api');
  const body = [
    `client_id=${encodeURIComponent(process.env.MS_CLIENT_ID)}`,
    `client_secret=${encodeURIComponent(process.env.MS_CLIENT_SECRET)}`,
    `scope=${encodeURIComponent('https://api.botframework.com/.default')}`,
    `grant_type=client_credentials`,
  ].join('&');
  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`,
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }
  );
  const data = await res.json();
  console.log('[getBotToken] token response status:', res.status, '| has token:', !!data.access_token);
  if (!data.access_token) {
    console.log('[getBotToken] ERROR:', JSON.stringify(data));
    throw new Error(`Bot token error: ${JSON.stringify(data)}`);
  }
  return data.access_token;
}

// ── Post reply to Teams via Azure Bot serviceUrl ───────────────────────────
async function postToTeams(activity, msgContent) {
  const serviceUrl     = activity.serviceUrl;
  const conversationId = activity.conversation?.id;
  const activityId     = activity.id;
  console.log('[postToTeams] serviceUrl:', serviceUrl);
  console.log('[postToTeams] conversationId:', conversationId);
  console.log('[postToTeams] activityId:', activityId);
  if (!serviceUrl || !conversationId) {
    console.log('[postToTeams] ERROR: missing serviceUrl or conversationId — cannot reply');
    return;
  }
  const token = await getBotFrameworkToken();
  const { fetch } = await import('@forge/api');
  const reply = {
    type: 'message',
    conversation: { id: conversationId },
    from:      { id: activity.recipient?.id, name: activity.recipient?.name },
    recipient: { id: activity.from?.id,      name: activity.from?.name },
    replyToId: activityId,
    ...msgContent,
  };
  const url = `${serviceUrl}v3/conversations/${encodeURIComponent(conversationId)}/activities/${activityId}`;
  console.log('[postToTeams] posting reply to:', url);
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(reply),
  });
  const resText = await res.text();
  console.log('[postToTeams] reply status:', res.status, '| body:', resText);
}

// ── sendBotReply — build message and POST via serviceUrl ──────────────────
async function sendBotReply(activity, cards, text) {
  const attachments = cards?.length
    ? cards.map(c => ({ contentType: 'application/vnd.microsoft.card.adaptive', content: c }))
    : undefined;
  const msgContent = {
    ...(text        ? { text }        : {}),
    ...(attachments ? { attachments } : {}),
  };
  await postToTeams(activity, msgContent);
  return { statusCode: 200, body: JSON.stringify({}) };
}

// ── providers.auth — COMMENTED ─────────────────────────────────────────────
// async function getMicrosoftClient() {
//   const client = api.asUser().withProvider('microsoft');
//   const account = await client.getAccount();
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
  const { teamId, channelId, teamName, channelName, webhookUrl } = payload;
  if (!teamId || !channelId) throw new Error('teamId and channelId are required');
  await storage.set('teams-channel-config', { teamId, channelId, teamName: teamName || '', channelName: channelName || '', webhookUrl: webhookUrl || '' });
  return { success: true };
});

resolver.define('getConfig', async () => {
  const config = await storage.get('teams-channel-config');
  return { config: config || null };
});

resolver.define('testConnection', async () => {
  const config = await storage.get('teams-channel-config');
  if (!config?.webhookUrl) return { success: false, error: 'No webhook URL configured.' };
  const { fetch } = await import('@forge/api');
  const res = await fetch(config.webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ "@type": "MessageCard", "@context": "http://schema.org/extensions", "themeColor": "0078D4", "summary": "Test", "sections": [{ "activityTitle": "✅ **Jira–Teams connector is active!**", "markdown": true }] }),
  });
  if (!res.ok) return { success: false, error: `Webhook ${res.status}: ${await res.text()}` };
  return { success: true };
});

resolver.define('getWebhookUrl', async ({ context }) => {
  const appId = context?.appId || '37fbf02b-1217-4b50-97f9-3565740801bb';
  const envId = context?.environmentId || 'development';
  return { url: `https://webhook.forge.atlassian.com/automation-teams-webhook/${appId}/${envId}` };
});

resolver.define('clearConfig', async () => {
  await storage.delete('teams-channel-config');
  return { success: true };
});

resolver.define('sendCustomMessage', async ({ payload }) => {
  const config = await storage.get('teams-channel-config');
  if (!config?.webhookUrl) return { success: false, error: 'No webhook URL configured.' };
  const { message } = payload;
  if (!message?.trim()) return { success: false, error: 'Message cannot be empty.' };
  const { fetch } = await import('@forge/api');
  const res = await fetch(config.webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ text: message.trim() }),
  });
  if (!res.ok) return { success: false, error: `Webhook ${res.status}: ${await res.text()}` };
  return { success: true };
});

// ── Debug resolver — read last bot activity from storage ──────────────────
resolver.define('getBotDebug', async () => {
  const debug = await storage.get('bot-debug-log');
  return { debug: debug || 'No bot activity yet' };
});

export const handler = resolver.getDefinitions();

// ── Feature 1: Jira link previews in Teams (bot endpoint) ──────────────────
export async function teamsBotHandler(req) {
  try {
    console.log('[teamsBotHandler] ─── INCOMING REQUEST ───');
    console.log('[teamsBotHandler] req.body exists:', !!req.body);
    const body = req.body ? JSON.parse(req.body) : {};
    console.log('[teamsBotHandler] type:', body.type, '| name:', body.name);
    console.log('[teamsBotHandler] from:', body.from?.name, '| channel:', body.channelId);
    console.log('[teamsBotHandler] serviceUrl:', body.serviceUrl);
    console.log('[teamsBotHandler] text:', body.text);

    // Save debug info to storage (visible via UI — works on AGC where logs are restricted)
    await storage.set('bot-debug-log', {
      receivedAt: new Date().toISOString(),
      type: body.type,
      name: body.name,
      text: body.text,
      from: body.from?.name,
      channelId: body.channelId,
      serviceUrl: body.serviceUrl,
      hasConversation: !!body.conversation?.id,
    });

    // Teams sends composeExtension/queryLink when a Jira URL is pasted
    if (body.type === 'invoke' && body.name === 'composeExtension/queryLink') {
      const url = body.value?.url || '';
      console.log('[teamsBotHandler] link unfurl URL:', url);

      // Extract issue key from Jira URL — e.g. /browse/PROJ-123
      const match = url.match(/\/browse\/([A-Z][A-Z0-9_]*-\d+)/i);
      if (!match) {
        return { statusCode: 200, body: JSON.stringify({ composeExtension: { type: 'result', attachmentLayout: 'list', attachments: [] } }) };
      }

      const issueKey = match[1].toUpperCase();
      console.log('[teamsBotHandler] fetching issue:', issueKey);

      // Fetch issue from Jira API
      let issue, f;
      try {
        const jiraRes = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}?fields=summary,status,priority,issuetype,assignee,reporter,description`);
        issue = await jiraRes.json();
        console.log('[teamsBotHandler] jira response status:', jiraRes.status);
        if (issue.errorMessages || issue.errors) {
          console.log('[teamsBotHandler] jira error:', JSON.stringify(issue));
          await storage.set('bot-debug-log', { error: 'Jira fetch failed', detail: JSON.stringify(issue), issueKey });
          return { statusCode: 200, body: JSON.stringify({ composeExtension: { type: 'result', attachmentLayout: 'list', attachments: [] } }) };
        }
      } catch (jiraErr) {
        console.log('[teamsBotHandler] jira fetch error:', jiraErr.message);
        await storage.set('bot-debug-log', { error: jiraErr.message, issueKey, step: 'jira-fetch' });
        return { statusCode: 200, body: JSON.stringify({ composeExtension: { type: 'result', attachmentLayout: 'list', attachments: [] } }) };
      }
      f = issue.fields || {};

      const summary  = f.summary || issueKey;
      const status   = f.status?.name || 'Unknown';
      const priority = f.priority?.name || 'None';
      const type     = f.issuetype?.name || 'Issue';
      const assignee = f.assignee?.displayName || 'Unassigned';
      const reporter = f.reporter?.displayName || 'Unknown';
      const jiraUrl  = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;

      // Adaptive Card for rich preview
      const card = {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
          {
            type: 'ColumnSet',
            columns: [
              {
                type: 'Column', width: 'auto',
                items: [{ type: 'Image', url: f.issuetype?.iconUrl || '', width: '24px', height: '24px' }]
              },
              {
                type: 'Column', width: 'stretch',
                items: [
                  { type: 'TextBlock', text: `**${issueKey}**`, wrap: true, size: 'Medium' },
                  { type: 'TextBlock', text: summary, wrap: true, spacing: 'None', color: 'Default' }
                ]
              }
            ]
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Status',   value: status },
              { title: 'Priority', value: priority },
              { title: 'Type',     value: type },
              { title: 'Assignee', value: assignee },
              { title: 'Reporter', value: reporter },
            ]
          }
        ],
        actions: [{ type: 'Action.OpenUrl', title: 'Open in Jira', url: jiraUrl }]
      };

      const response = {
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: card,
            preview: {
              contentType: 'application/vnd.microsoft.card.thumbnail',
              content: {
                title: `${issueKey}: ${summary}`,
                subtitle: `${type} · ${status} · ${priority}`,
                text: `Assignee: ${assignee}`,
                buttons: [{ type: 'openUrl', title: 'Open in Jira', value: jiraUrl }]
              }
            }
          }]
        }
      };

      return { statusCode: 200, body: JSON.stringify(response), headers: { 'Content-Type': 'application/json' } };
    }

    // ── Feature 2: Bot commands — create / search / update ──────────────────
    if (body.type === 'message') {
      // Strip @mention tags, trim
      const text = (body.text || '').replace(/<at>[^<]*<\/at>/gi, '').trim();
      console.log('[teamsBotHandler] message command:', text);

      // BUG / TASK / STORY / EPIC — typed issue creation
      const typedMatch = text.match(/^(bug|task|story|epic)\s+([A-Z][A-Z0-9_]*)\s+(.+)$/i);
      if (typedMatch) {
        const issueType  = typedMatch[1].charAt(0).toUpperCase() + typedMatch[1].slice(1).toLowerCase();
        const projectKey = typedMatch[2].toUpperCase();
        const summary    = typedMatch[3].trim();
        const emojiMap   = { Bug: '🐛', Task: '✅', Story: '📖', Epic: '⚡' };
        const createRes  = await api.asApp().requestJira(route`/rest/api/3/issue`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ fields: { project: { key: projectKey }, summary, issuetype: { name: issueType } } })
        });
        const created = await createRes.json();
        if (created.key) {
          const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${created.key}`;
          return sendBotReply(body, [issueCard(created.key, summary, issueType, 'To Do', 'None', 'Unassigned', issueUrl, `${emojiMap[issueType] || '🆕'} ${issueType} Created`)]);
        }
        return sendBotReply(body, [], `❌ Failed to create ${issueType}: ${JSON.stringify(created.errors || created)}`);
      }

      // CREATE: "create PROJ Summary of the issue"
      const createMatch = text.match(/^create\s+([A-Z][A-Z0-9_]*)\s+(.+)$/i);
      if (createMatch) {
        const projectKey = createMatch[1].toUpperCase();
        const summary    = createMatch[2].trim();

        // Fetch valid issue types for this project
        const metaRes   = await api.asApp().requestJira(route`/rest/api/3/issue/createmeta?projectKeys=${projectKey}&expand=projects.issuetypes`);
        const metaData  = await metaRes.json();
        const issueTypes = metaData.projects?.[0]?.issuetypes || [];
        const preferred  = ['Task', 'Story', 'Bug', 'Subtask'];
        const issueType  = issueTypes.find(t => preferred.includes(t.name))?.name || issueTypes[0]?.name || 'Task';

        const createRes  = await api.asApp().requestJira(route`/rest/api/3/issue`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fields: {
              project:   { key: projectKey },
              summary,
              issuetype: { name: issueType },
            }
          })
        });
        const created = await createRes.json();
        if (created.key) {
          const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${created.key}`;
          return sendBotReply(body, [issueCard(created.key, summary, issueType, 'To Do', 'None', 'Unassigned', issueUrl, '✅ Issue Created')]);
        }
        return sendBotReply(body, [], `❌ Failed to create issue: ${JSON.stringify(created.errors || created)}`);
      }

      // SEARCH: "search payment bug" or "search assignee = currentUser()"
      const searchMatch = text.match(/^search\s+(.+)$/i);
      if (searchMatch) {
        const query = searchMatch[1].trim();
        const jql   = query.includes('=') || query.includes('ORDER') ? query : `text ~ "${query}" ORDER BY updated DESC`;
        const searchRes = await api.asApp().requestJira(route`/rest/api/3/search?jql=${jql}&maxResults=5&fields=summary,status,priority,issuetype,assignee`);
        const data  = await searchRes.json();
        const issues = (data.issues || []);
        if (!issues.length) return sendBotReply(body, [], `No issues found for: **${query}**`);
        const cards = issues.map(i => {
          const f = i.fields || {};
          return issueCard(i.key, f.summary, f.issuetype?.name, f.status?.name, f.priority?.name, f.assignee?.displayName || 'Unassigned', `${process.env.JIRA_BASE_URL}/browse/${i.key}`, '🔍 Search Result');
        });
        return sendBotReply(body, cards, `Found ${issues.length} issue(s):`);
      }

      // PROJECTS: "projects" — list all Jira projects
      if (/^projects$/i.test(text)) {
        const projRes  = await api.asApp().requestJira(route`/rest/api/3/project/search?maxResults=20&orderBy=name`);
        const projData = await projRes.json();
        const projects = (projData.values || []);
        if (!projects.length) return sendBotReply(body, [], `No projects found.`);
        const list = projects.map(p => `• **${p.key}** — ${p.name}`).join('\n');
        return sendBotReply(body, [], `**Jira Projects (${projects.total || projects.length}):**\n\n${list}`);
      }

      // SHOW: "show AITEST-40" — fetch and display issue details
      const showMatch = text.match(/^show\s+([A-Z][A-Z0-9_]*-\d+)$/i);
      if (showMatch) {
        const issueKey = showMatch[1].toUpperCase();
        const jiraRes  = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}?fields=summary,status,priority,issuetype,assignee,reporter,description,created,updated`);
        const issue    = await jiraRes.json();
        if (issue.errorMessages?.length || issue.errors) {
          return sendBotReply(body, [], `❌ Issue **${issueKey}** not found or no permission.`);
        }
        const f        = issue.fields || {};
        const summary  = f.summary   || '(no summary)';
        const status   = f.status?.name     || 'Unknown';
        const priority = f.priority?.name   || 'None';
        const type     = f.issuetype?.name  || 'Issue';
        const assignee = f.assignee?.displayName || 'Unassigned';
        const reporter = f.reporter?.displayName || 'Unknown';
        const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
        return sendBotReply(body, [issueCard(issueKey, summary, type, status, priority, assignee, issueUrl, `📋 ${reporter} reported`)]);
      }

      // UPDATE: "update PROJ-123 Done" (transition by status name)
      const updateMatch = text.match(/^update\s+([A-Z][A-Z0-9_]*-\d+)\s+(.+)$/i);
      if (updateMatch) {
        const issueKey  = updateMatch[1].toUpperCase();
        const newStatus = updateMatch[2].trim();
        const transRes  = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}/transitions`);
        const transData = await transRes.json();
        const transition = (transData.transitions || []).find(t => t.name.toLowerCase() === newStatus.toLowerCase());
        if (!transition) {
          const available = (transData.transitions || []).map(t => t.name).join(', ');
          return sendBotReply(body, [], `❌ Status **"${newStatus}"** not found. Available: ${available}`);
        }
        await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}/transitions`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ transition: { id: transition.id } })
        });
        return sendBotReply(body, [], `✅ **${issueKey}** updated to **${transition.name}**`);
      }

      // HELP — default
      return sendBotReply(body, [], helpText());
    }

    // Default response for other activity types (ping, etc.)
    return { statusCode: 200, body: JSON.stringify({}) };
  } catch (e) {
    console.log('[teamsBotHandler] error:', e.message);
    return { statusCode: 500, body: JSON.stringify({ error: e.message }) };
  }
}

// ── Bot response helpers ────────────────────────────────────────────────────
function botReply(cards, text) {
  const msg = { type: 'message' };
  if (text) msg.text = text;
  if (cards?.length) {
    msg.attachments = cards.map(card => ({
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: card,
    }));
  }
  return { statusCode: 200, body: JSON.stringify(msg), headers: { 'Content-Type': 'application/json' } };
}

function issueCard(key, summary, type, status, priority, assignee, url, label) {
  return {
    type: 'AdaptiveCard', version: '1.4',
    body: [
      { type: 'TextBlock', text: label || '', size: 'Small', color: 'Accent', weight: 'Bolder' },
      { type: 'TextBlock', text: `**[${key}](${url})** — ${summary}`, wrap: true, size: 'Medium' },
      {
        type: 'FactSet', facts: [
          { title: 'Type',     value: type     || 'Issue' },
          { title: 'Status',   value: status   || 'Unknown' },
          { title: 'Priority', value: priority || 'None' },
          { title: 'Assignee', value: assignee || 'Unassigned' },
        ]
      }
    ],
    actions: [{ type: 'Action.OpenUrl', title: 'Open in Jira', url }]
  };
}

function helpText() {
  return `**Jira Bot Commands:**\n\n` +
    `📁 **projects** — List all Jira projects\n` +
    `🆕 **create** \`PROJECT\` \`summary\` — Create issue (auto type)\n` +
    `🐛 **bug** \`PROJECT\` \`summary\` — Create Bug\n` +
    `✅ **task** \`PROJECT\` \`summary\` — Create Task\n` +
    `📖 **story** \`PROJECT\` \`summary\` — Create Story\n` +
    `⚡ **epic** \`PROJECT\` \`summary\` — Create Epic\n` +
    `🔍 **search** \`keyword or JQL\` — Search issues (top 5)\n` +
    `📋 **show** \`PROJ-123\` — View issue details\n` +
    `✏️ **update** \`PROJ-123\` \`Status\` — Transition issue status\n\n` +
    `_Examples:_\n` +
    `• \`bug AITEST Login page is broken\`\n` +
    `• \`task MS Deploy to production\`\n` +
    `• \`story AITEST User can login with SSO\`\n` +
    `• \`search payment bug\`\n` +
    `• \`show AITEST-40\`\n` +
    `• \`update AITEST-40 In Progress\``;
}

// ── Feature 6: Automation webhook (webtrigger) ─────────────────────────────
// Jira Automation rule calls this HTTP endpoint → posts to Teams
export async function automationHandler(req) {
  try {
    const body = req.body ? JSON.parse(req.body) : {};
    const config = await storage.get('teams-channel-config');
    if (!config?.teamId || !config?.channelId) {
      return { statusCode: 400, body: JSON.stringify({ error: 'No channel configured' }) };
    }

    const token = await getAppToken();
    const { fetch } = await import('@forge/api');

    const title   = body.title   || body.summary  || 'Automation Rule Triggered';
    const detail  = body.detail  || body.message  || '';
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

// ── jiraSync trigger (active) ───────────────────────────────────────────────
export async function jiraSync(event) {
  console.log('[jiraSync] event received:', event.eventType);

  const config = await storage.get('teams-channel-config');
  if (!config?.teamId || !config?.channelId) {
    console.log('[jiraSync] no channel config saved — skipping');
    return;
  }

  if (!config.webhookUrl) {
    console.log('[jiraSync] no webhookUrl in config — skipping');
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
  const color     = eventType === 'jira:issueCreated' ? '0078D4' : eventType === 'jira:issueDeleted' ? 'D83B01' : 'FFB900';

  // Power Automate (Teams Workflows) expects plain text field
  const payload = {
    text: `${action}: **${issueKey}** — ${summary}\nType: ${issueType} | Status: ${status} | Priority: ${priority}\nAssignee: ${assignee} | Reporter: ${reporter}\n🔗 ${jiraUrl}`,
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

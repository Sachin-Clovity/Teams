import * as ResolverPkg from '@forge/resolver';
import api, { route, storage } from '@forge/api';

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

// All Graph calls use app-level token from client credentials — no user OAuth needed.
async function graphGet(path) {
  const token = await getAppToken();
  const { fetch } = await import('@forge/api');
  const res = await fetch(`https://graph.microsoft.com${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph GET ${path} failed ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphPost(path, bodyObj) {
  const token = await getAppToken();
  const { fetch } = await import('@forge/api');
  const res = await fetch(`https://graph.microsoft.com${path}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(bodyObj),
  });
  if (!res.ok) throw new Error(`Graph POST ${path} failed ${res.status}: ${await res.text()}`);
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


const resolver = new Resolver();

resolver.define('getTeams', async () => {
  console.log('[getTeams] called');
  const data = await graphGet('/v1.0/teams');
  const teams = (data.value || []).map(t => ({ id: t.id, displayName: t.displayName }));
  console.log('[getTeams] count:', teams.length);
  return { teams };
});

resolver.define('getChannels', async ({ payload }) => {
  console.log('[getChannels] teamId:', payload.teamId);
  const data = await graphGet(`/v1.0/teams/${payload.teamId}/channels`);
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

// ── Auth debug resolver — verify app token works ──────────────────────────
resolver.define('getAuthStatus', async () => {
  console.log('[authDebug] checking app-level token');
  try {
    const data = await graphGet('/v1.0/organization');
    return { authenticated: true, org: data.value?.[0]?.displayName || 'OK' };
  } catch (e) {
    console.log('[authDebug] ERROR:', e.message);
    return { authenticated: false, error: e.message };
  }
});

async function getAadUserId(email) {
  const data = await graphGet(`/v1.0/users/${encodeURIComponent(email)}`);
  if (!data.id) throw new Error(`User not found in Azure AD: ${email}`);
  return data.id;
}

async function postChatMessage(chatId, htmlContent) {
  await graphPost(`/v1.0/chats/${chatId}/messages`, {
    body: { contentType: 'html', content: htmlContent },
  });
}

// ── Issue Context Panel resolvers ──────────────────────────────────────────

resolver.define('getNotificationSettings', async () => {
  const settings = await storage.get('notification-settings');
  return { settings: settings || { created: true, updated: true, deleted: false, commented: false } };
});

resolver.define('saveNotificationSettings', async ({ payload }) => {
  await storage.set('notification-settings', payload.settings);
  return { success: true };
});

resolver.define('getIssueDetails', async ({ payload }) => {
  const { issueKey } = payload;
  const res = await api.asApp().requestJira(
    route`/rest/api/3/issue/${issueKey}?fields=summary,status,assignee,reporter,priority,issuetype`
  );
  const data = await res.json();
  const f = data.fields || {};
  return {
    key: data.key,
    summary: f.summary || '',
    status: f.status?.name || '',
    priority: f.priority?.name || '',
    issueType: f.issuetype?.name || '',
    assignee: f.assignee ? { name: f.assignee.displayName, email: f.assignee.emailAddress || '' } : null,
    reporter: f.reporter ? { name: f.reporter.displayName, email: f.reporter.emailAddress || '' } : null,
  };
});

resolver.define('getCurrentUser', async () => {
  const res = await api.asUser().requestJira(route`/rest/api/3/myself`);
  const data = await res.json();
  return { email: data.emailAddress || '', name: data.displayName || '' };
});

resolver.define('getIssueTransitions', async ({ payload }) => {
  const res  = await api.asApp().requestJira(route`/rest/api/3/issue/${payload.issueKey}/transitions`);
  const data = await res.json();
  return { transitions: (data.transitions || []).map(t => ({ id: t.id, name: t.name })) };
});

resolver.define('transitionIssue', async ({ payload }) => {
  await api.asApp().requestJira(route`/rest/api/3/issue/${payload.issueKey}/transitions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ transition: { id: payload.transitionId } }),
  });
  return { success: true };
});

resolver.define('addIssueComment', async ({ payload }) => {
  const res = await api.asApp().requestJira(route`/rest/api/3/issue/${payload.issueKey}/comment`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      body: { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: payload.commentText }] }] },
    }),
  });
  const data = await res.json();
  return { success: !!data.id };
});

resolver.define('logIssueWork', async ({ payload }) => {
  const seconds = parseTimeStrToSeconds(payload.timeSpent);
  if (!seconds) throw new Error('Invalid time format. Use: 2h, 30m, 1h 30m');
  const body = { timeSpentSeconds: seconds };
  if (payload.description?.trim()) {
    body.comment = { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: payload.description.trim() }] }] };
  }
  const res = await api.asApp().requestJira(route`/rest/api/3/issue/${payload.issueKey}/worklog`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  return { success: !!data.id };
});

resolver.define('assignIssue', async ({ payload }) => {
  const userRes = await api.asApp().requestJira(route`/rest/api/3/user/search?query=${payload.email}&maxResults=1`);
  const users   = await userRes.json();
  const user    = Array.isArray(users) ? users[0] : null;
  if (!user) throw new Error(`User not found in Jira: ${payload.email}`);
  await api.asApp().requestJira(route`/rest/api/3/issue/${payload.issueKey}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields: { assignee: { accountId: user.accountId } } }),
  });
  return { success: true, displayName: user.displayName };
});

resolver.define('getIssueActivity', async ({ payload }) => {
  const res  = await api.asApp().requestJira(route`/rest/api/3/issue/${payload.issueKey}/comment?maxResults=5&orderBy=-created`);
  const data = await res.json();
  const comments = (data.comments || []).map(c => ({
    id:      c.id,
    author:  c.author?.displayName || 'Unknown',
    created: c.created,
    text:    extractTextFromADF(c.body),
  }));
  return { comments };
});

function parseTimeStrToSeconds(str) {
  let s = 0;
  const h = str?.match(/(\d+)\s*h/i);
  const m = str?.match(/(\d+)\s*m/i);
  if (h) s += parseInt(h[1], 10) * 3600;
  if (m) s += parseInt(m[1], 10) * 60;
  return s;
}

function extractTextFromADF(node) {
  if (!node) return '';
  if (node.type === 'text') return node.text || '';
  if (node.content) return node.content.map(extractTextFromADF).join('');
  return '';
}

// Creates a 1:1 Teams DM between fromEmail and toEmail, posts an issue link
resolver.define('startDM', async ({ payload }) => {
  const { fromEmail, toEmail, issueKey, issueSummary } = payload;
  const [fromId, toId] = await Promise.all([
    getAadUserId(fromEmail),
    getAadUserId(toEmail),
  ]);
  const chat = await graphPost('/v1.0/chats', {
    chatType: 'oneOnOne',
    members: [
      { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${fromId}` },
      { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${toId}` },
    ],
  });
  const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
  await postChatMessage(chat.id, `<b>Jira: <a href="${issueUrl}">${issueKey}</a></b> — ${issueSummary}`);
  return { success: true };
});

// Creates a named group Teams chat with all provided emails, posts an issue link
resolver.define('startGroupChat', async ({ payload }) => {
  const { emails, issueKey, issueSummary } = payload;
  const userIds = await Promise.all(emails.map(e => getAadUserId(e.trim())));
  const chat = await graphPost('/v1.0/chats', {
    chatType: 'group',
    topic: `${issueKey}: ${issueSummary.slice(0, 60)}`,
    members: userIds.map(id => ({
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${id}`,
    })),
  });
  const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
  await postChatMessage(chat.id, `<b>Jira: <a href="${issueUrl}">${issueKey}</a></b> — ${issueSummary}`);
  return { success: true };
});

// Manually posts this issue to the saved Teams channel configuration
resolver.define('postToChannelManual', async ({ payload }) => {
  const { issueKey, issueSummary } = payload;
  const config = await storage.get('teams-channel-config');
  if (!config?.teamId || !config?.channelId) throw new Error('No channel configured. Go to Teams Connector settings first.');
  const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
  await graphPost(`/v1.0/teams/${config.teamId}/channels/${config.channelId}/messages`, {
    body: { contentType: 'html', content: `<b>Jira: <a href="${issueUrl}">${issueKey}</a></b> — ${issueSummary}` },
  });
  return { success: true };
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
        version: '1.2',
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

    // ── Message actions: fetchTask — route by commandId ──────────────────
    if (body.type === 'invoke' && body.name === 'composeExtension/fetchTask') {
      const commandId  = body.value?.commandId || 'createJiraIssue';
      const rawContent = body.value?.messagePayload?.body?.content || '';
      const messageText = rawContent.replace(/<[^>]+>/g, '').trim().slice(0, 500);
      console.log('[teamsBotHandler] fetchTask commandId:', commandId, '| text preview:', messageText.slice(0, 60));

      let taskTitle, cardContent;
      if (commandId === 'commentInJira') {
        taskTitle   = 'Comment in Jira';
        cardContent = buildCommentCard(messageText);
      } else if (commandId === 'logTimeInJira') {
        taskTitle   = 'Log Time in Jira';
        cardContent = buildLogTimeCard(messageText);
      } else {
        taskTitle   = 'Create Jira Issue';
        cardContent = buildCreateIssueCard(messageText);
      }

      return {
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          task: {
            type: 'continue',
            value: {
              title: taskTitle,
              height: 'medium',
              width: 'medium',
              card: {
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: cardContent,
              },
            },
          },
        }),
      };
    }

    // ── Message actions: submitAction — route by commandId ────────────────
    if (body.type === 'invoke' && body.name === 'composeExtension/submitAction') {
      const commandId = body.value?.commandId || body.data?.commandId || 'createJiraIssue';
      const data = body.data || body.value?.data || {};
      console.log('[teamsBotHandler] submitAction commandId:', commandId, '| data:', JSON.stringify(data));

      // ── Comment in Jira ──────────────────────────────────────────────────
      if (commandId === 'commentInJira') {
        const { issueKey, commentText } = data;
        if (!issueKey || !commentText) {
          return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: '❌ Issue Key and Comment are required.' } }) };
        }
        const commentRes = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey.trim().toUpperCase()}/comment`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            body: { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: commentText.trim() }] }] },
          }),
        });
        const commented = await commentRes.json();
        if (commented.id) {
          const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey.trim().toUpperCase()}`;
          return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: `✅ Comment added to **${issueKey.trim().toUpperCase()}** — [View in Jira](${issueUrl})` } }) };
        }
        return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: `❌ Failed to add comment: ${JSON.stringify(commented.errors || commented)}` } }) };
      }

      // ── Log time in Jira ─────────────────────────────────────────────────
      if (commandId === 'logTimeInJira') {
        const { issueKey, timeSpent, workDescription } = data;
        if (!issueKey || !timeSpent) {
          return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: '❌ Issue Key and Time Spent are required. e.g. 2h, 30m, 1h 30m' } }) };
        }
        const seconds = parseTimeToSeconds(timeSpent);
        if (!seconds) {
          return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: '❌ Invalid time format. Use: 2h, 30m, 1h 30m, 90m' } }) };
        }
        const worklogBody = { timeSpentSeconds: seconds };
        if (workDescription?.trim()) {
          worklogBody.comment = { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: workDescription.trim() }] }] };
        }
        const logRes = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey.trim().toUpperCase()}/worklog`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(worklogBody),
        });
        const logged = await logRes.json();
        if (logged.id) {
          const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey.trim().toUpperCase()}`;
          return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: `✅ Logged **${timeSpent}** on **${issueKey.trim().toUpperCase()}** — [View in Jira](${issueUrl})` } }) };
        }
        return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: `❌ Failed to log time: ${JSON.stringify(logged.errors || logged)}` } }) };
      }

      // ── Create Jira Issue (default) ──────────────────────────────────────
      const { projectKey, summary, issueType, description } = data;
      console.log('[teamsBotHandler] submitAction — project:', projectKey, '| summary:', summary);

      if (!projectKey || !summary) {
        return {
          statusCode: 200,
          body: JSON.stringify({ task: { type: 'message', value: '❌ Project Key and Summary are required.' } }),
        };
      }

      const fields = {
        project:   { key: projectKey.trim().toUpperCase() },
        summary:   summary.trim(),
        issuetype: { name: issueType || 'Task' },
      };
      if (description?.trim()) {
        fields.description = {
          type: 'doc', version: 1,
          content: [{ type: 'paragraph', content: [{ type: 'text', text: description.trim() }] }],
        };
      }

      const createRes = await api.asApp().requestJira(route`/rest/api/3/issue`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields }),
      });
      const created = await createRes.json();
      console.log('[teamsBotHandler] submitAction — created key:', created.key);

      if (created.key) {
        const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${created.key}`;
        return {
          statusCode: 200,
          body: JSON.stringify({
            task: {
              type: 'message',
              value: `✅ **${created.key}** created — [Open in Jira](${issueUrl})`,
            },
          }),
        };
      }
      const errDetail = JSON.stringify(created.errors || created.errorMessages || created);
      return {
        statusCode: 200,
        body: JSON.stringify({ task: { type: 'message', value: `❌ Failed to create issue: ${errDetail}` } }),
      };
    }

    // ── Compose extension search: "Search Jira Issues" from message bar ──
    if (body.type === 'invoke' && body.name === 'composeExtension/query') {
      const commandId   = body.value?.commandId || '';
      const searchQuery = (body.value?.parameters?.[0]?.value || '').trim();
      console.log('[teamsBotHandler] composeExtension/query commandId:', commandId, '| query:', searchQuery);

      const jql = searchQuery
        ? `text ~ "${searchQuery.replace(/"/g, '\\"')}" ORDER BY updated DESC`
        : 'ORDER BY updated DESC';

      const searchRes = await api.asApp().requestJira(route`/rest/api/3/search?jql=${jql}&maxResults=8&fields=summary,status,priority,issuetype,assignee`);
      const searchData = await searchRes.json();
      const issues = searchData.issues || [];

      const attachments = issues.map(i => {
        const f   = i.fields || {};
        const url = `${process.env.JIRA_BASE_URL}/browse/${i.key}`;
        return {
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: issueCard(i.key, f.summary, f.issuetype?.name, f.status?.name, f.priority?.name, f.assignee?.displayName || 'Unassigned', url, `${f.issuetype?.name || 'Issue'} · ${f.status?.name || ''}`),
          preview: {
            contentType: 'application/vnd.microsoft.card.thumbnail',
            content: {
              title: `${i.key}: ${f.summary || ''}`,
              subtitle: `${f.issuetype?.name || 'Issue'} · ${f.status?.name || 'Unknown'} · ${f.priority?.name || 'None'}`,
              text: `Assignee: ${f.assignee?.displayName || 'Unassigned'}`,
              buttons: [{ type: 'openUrl', title: 'Open in Jira', value: url }],
            },
          },
        };
      });

      return {
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          composeExtension: {
            type: 'result',
            attachmentLayout: 'list',
            attachments,
          },
        }),
      };
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

      // PROJECTS: "projects" / "project" / "project list" — list all Jira projects
      if (/^projects?(\s+list)?$/i.test(text)) {
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

      // COMMENT: "comment PROJ-123 This is a comment"
      const commentMatch = text.match(/^comment\s+([A-Z][A-Z0-9_]*-\d+)\s+(.+)$/i);
      if (commentMatch) {
        const issueKey   = commentMatch[1].toUpperCase();
        const commentTxt = commentMatch[2].trim();
        const commentRes = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}/comment`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            body: { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: commentTxt }] }] },
          }),
        });
        const commented = await commentRes.json();
        if (commented.id) {
          const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
          return sendBotReply(body, [], `✅ Comment added to **[${issueKey}](${issueUrl})**`);
        }
        return sendBotReply(body, [], `❌ Failed to comment: ${JSON.stringify(commented.errors || commented)}`);
      }

      // LOG: "log PROJ-123 2h Worked on the bug fix"
      const logMatch = text.match(/^log\s+([A-Z][A-Z0-9_]*-\d+)\s+(\d+[hm](?:\s+\d+[hm])?)\s*(.*)$/i);
      if (logMatch) {
        const issueKey   = logMatch[1].toUpperCase();
        const timeStr    = logMatch[2].trim();
        const logNote    = logMatch[3].trim();
        const seconds    = parseTimeToSeconds(timeStr);
        if (!seconds) return sendBotReply(body, [], `❌ Invalid time format. Use: 2h, 30m, 1h 30m`);
        const worklogBody = { timeSpentSeconds: seconds };
        if (logNote) {
          worklogBody.comment = { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: logNote }] }] };
        }
        const logRes = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}/worklog`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(worklogBody),
        });
        const logged = await logRes.json();
        if (logged.id) {
          const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
          return sendBotReply(body, [], `✅ Logged **${timeStr}** on **[${issueKey}](${issueUrl})**`);
        }
        return sendBotReply(body, [], `❌ Failed to log time: ${JSON.stringify(logged.errors || logged)}`);
      }

      // ASSIGN: "assign PROJ-123 user@domain.com"
      const assignMatch = text.match(/^assign\s+([A-Z][A-Z0-9_]*-\d+)\s+(\S+@\S+)$/i);
      if (assignMatch) {
        const issueKey = assignMatch[1].toUpperCase();
        const email    = assignMatch[2].trim();
        // Find accountId from email
        const userRes  = await api.asApp().requestJira(route`/rest/api/3/user/search?query=${email}&maxResults=1`);
        const users    = await userRes.json();
        const user     = Array.isArray(users) ? users[0] : null;
        if (!user) return sendBotReply(body, [], `❌ User **${email}** not found in Jira.`);
        await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ fields: { assignee: { accountId: user.accountId } } }),
        });
        const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
        return sendBotReply(body, [], `✅ **${issueKey}** assigned to **${user.displayName}** — [View](${issueUrl})`);
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

// Adaptive Card shown in the task module when user clicks three-dot → Create Jira Issue
function buildCreateIssueCard(prefillText) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Create Jira Issue', weight: 'Bolder', size: 'Medium' },
      { type: 'Input.Text', id: 'projectKey', label: 'Project Key', placeholder: 'e.g. PROJ', isRequired: true },
      { type: 'Input.Text', id: 'summary',    label: 'Summary',     placeholder: 'Short description', value: prefillText.slice(0, 255), isRequired: true },
      {
        type: 'Input.ChoiceSet', id: 'issueType', label: 'Issue Type', value: 'Task',
        choices: [
          { title: 'Task',  value: 'Task' },
          { title: 'Bug',   value: 'Bug' },
          { title: 'Story', value: 'Story' },
          { title: 'Epic',  value: 'Epic' },
        ],
      },
      { type: 'Input.Text', id: 'description', label: 'Description', placeholder: 'Additional context', value: prefillText, isMultiline: true },
    ],
    actions: [{ type: 'Action.Submit', title: 'Create Issue' }],
  };
}

function buildCommentCard(prefillText) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Add Comment to Jira Issue', weight: 'Bolder', size: 'Medium' },
      { type: 'Input.Text', id: 'issueKey',    label: 'Issue Key',  placeholder: 'e.g. PROJ-123', isRequired: true },
      { type: 'Input.Text', id: 'commentText', label: 'Comment',    placeholder: 'Your comment...', value: prefillText, isMultiline: true, isRequired: true },
    ],
    actions: [{ type: 'Action.Submit', title: 'Add Comment', data: { commandId: 'commentInJira' } }],
  };
}

function buildLogTimeCard(prefillText) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Log Work Time in Jira', weight: 'Bolder', size: 'Medium' },
      { type: 'Input.Text', id: 'issueKey',      label: 'Issue Key',       placeholder: 'e.g. PROJ-123', isRequired: true },
      { type: 'Input.Text', id: 'timeSpent',     label: 'Time Spent',      placeholder: 'e.g. 2h, 30m, 1h 30m', isRequired: true },
      { type: 'Input.Text', id: 'workDescription', label: 'Work Description', placeholder: 'What did you work on?', value: prefillText, isMultiline: true },
    ],
    actions: [{ type: 'Action.Submit', title: 'Log Time', data: { commandId: 'logTimeInJira' } }],
  };
}

function parseTimeToSeconds(timeStr) {
  let seconds = 0;
  const hours = timeStr.match(/(\d+)\s*h/i);
  const mins  = timeStr.match(/(\d+)\s*m/i);
  if (hours) seconds += parseInt(hours[1], 10) * 3600;
  if (mins)  seconds += parseInt(mins[1],  10) * 60;
  return seconds;
}

function issueCard(key, summary, type, status, priority, assignee, url, label) {
  return {
    type: 'AdaptiveCard', version: '1.2',
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
    `🔍 **search** \`keyword or JQL\` — Search issues (top 8)\n` +
    `📋 **show** \`PROJ-123\` — View issue details\n` +
    `✏️ **update** \`PROJ-123\` \`Status\` — Transition issue status\n` +
    `💬 **comment** \`PROJ-123\` \`text\` — Add a comment\n` +
    `⏱️ **log** \`PROJ-123\` \`2h\` \`note\` — Log work time\n` +
    `👤 **assign** \`PROJ-123\` \`user@domain.com\` — Assign issue\n\n` +
    `_Message actions (right-click any message):_\n` +
    `• **Create issue in Jira** — turns message into an issue\n` +
    `• **Comment in Jira** — adds message as a comment\n` +
    `• **Log time in Jira** — logs work against an issue\n\n` +
    `_Search (compose bar):_\n` +
    `• Click **+** → **Jira Connector** → type to search issues\n\n` +
    `_Examples:_\n` +
    `• \`bug AITEST Login page is broken\`\n` +
    `• \`search payment bug\`\n` +
    `• \`show AITEST-40\`\n` +
    `• \`update AITEST-40 In Progress\`\n` +
    `• \`comment AITEST-40 Fixed in latest build\`\n` +
    `• \`log AITEST-40 2h Investigated root cause\`\n` +
    `• \`assign AITEST-40 dev@clovity.com\``;
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

  const notifSettings = await storage.get('notification-settings') || { created: true, updated: true, deleted: false, commented: false };
  const issue = event.issue || {};
  const fields = issue.fields || {};
  const eventType = event.eventType || 'jira:issueUpdated';

  if (eventType === 'jira:issueCreated' && !notifSettings.created) { console.log('[jiraSync] created notifications disabled'); return; }
  if (eventType === 'jira:issueUpdated' && !notifSettings.updated) { console.log('[jiraSync] updated notifications disabled'); return; }
  if (eventType === 'jira:issueDeleted' && !notifSettings.deleted) { console.log('[jiraSync] deleted notifications disabled'); return; }

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

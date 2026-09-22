import api, { route } from '@forge/api';
import { getGlobalConfig, getMsTenantId } from '../storage/kvsStore.js';
import { getAadUserId, postChatMessage } from '../graph/chat.js';
import { graphPost } from '../graph/client.js';
import { parseTimeToSeconds, extractTextFromADF } from '../jira/utils.js';

export function registerIssuePanelResolvers(resolver) {
  resolver.define('getIssueDetails', async ({ payload }) => {
    const { issueKey } = payload;
    const res = await api.asUser().requestJira(
      route`/rest/api/3/issue/${issueKey}?fields=summary,status,assignee,reporter,priority,issuetype,labels,duedate`
    );
    const data = await res.json();
    const f = data.fields || {};
    return {
      key: data.key,
      summary: f.summary || '',
      status: f.status?.name || '',
      priority: f.priority?.name || '',
      issueType: f.issuetype?.name || '',
      labels: f.labels || [],
      dueDate: f.duedate || '',
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
    const res  = await api.asUser().requestJira(route`/rest/api/3/issue/${payload.issueKey}/transitions`);
    const data = await res.json();
    return { transitions: (data.transitions || []).map(t => ({ id: t.id, name: t.name })) };
  });

  resolver.define('transitionIssue', async ({ payload }) => {
    await api.asUser().requestJira(route`/rest/api/3/issue/${payload.issueKey}/transitions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ transition: { id: payload.transitionId } }),
    });
    return { success: true };
  });

  resolver.define('addIssueComment', async ({ payload }) => {
    const res = await api.asUser().requestJira(route`/rest/api/3/issue/${payload.issueKey}/comment`, {
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
    const seconds = parseTimeToSeconds(payload.timeSpent);
    if (!seconds) throw new Error('Invalid time format. Use: 2h, 30m, 1h 30m');
    const body = { timeSpentSeconds: seconds };
    if (payload.description?.trim()) {
      body.comment = { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: payload.description.trim() }] }] };
    }
    const res = await api.asUser().requestJira(route`/rest/api/3/issue/${payload.issueKey}/worklog`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    return { success: !!data.id };
  });

  resolver.define('assignIssue', async ({ payload }) => {
    const userRes = await api.asUser().requestJira(route`/rest/api/3/user/search?query=${payload.email}&maxResults=1`);
    const users   = await userRes.json();
    const user    = Array.isArray(users) ? users[0] : null;
    if (!user) throw new Error(`User not found in Jira: ${payload.email}`);
    await api.asUser().requestJira(route`/rest/api/3/issue/${payload.issueKey}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields: { assignee: { accountId: user.accountId } } }),
    });
    return { success: true, displayName: user.displayName };
  });

  resolver.define('getIssueActivity', async ({ payload }) => {
    const res  = await api.asUser().requestJira(route`/rest/api/3/issue/${payload.issueKey}/comment?maxResults=5&orderBy=-created`);
    const data = await res.json();
    const comments = (data.comments || []).map(c => ({
      id:      c.id,
      author:  c.author?.displayName || 'Unknown',
      created: c.created,
      text:    extractTextFromADF(c.body),
    }));
    return { comments };
  });

  // Edit issue fields (priority / labels / due date) from the issue panel
  resolver.define('updateIssueFields', async ({ payload }) => {
    const { issueKey, priority, labels, dueDate } = payload;
    const fields = {};
    if (priority?.id) fields.priority = { id: priority.id };
    if (labels !== undefined) fields.labels = labels;
    if (dueDate !== undefined) fields.duedate = dueDate || null;
    if (!Object.keys(fields).length) return { success: false, error: 'No fields to update.' };
    await api.asUser().requestJira(route`/rest/api/3/issue/${issueKey}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields }),
    });
    return { success: true };
  });

  // Creates a 1:1 Teams DM between fromEmail and toEmail, posts an issue link
  resolver.define('startDM', async ({ payload }) => {
    const { fromEmail, toEmail, issueKey, issueSummary } = payload;
    const tenantId = await getMsTenantId();
    const [fromId, toId] = await Promise.all([getAadUserId(fromEmail, tenantId), getAadUserId(toEmail, tenantId)]);
    const chat = await graphPost('/v1.0/chats', {
      chatType: 'oneOnOne',
      members: [
        { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${fromId}` },
        { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${toId}` },
      ],
    }, tenantId);
    const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
    await postChatMessage(chat.id, `<b>Jira: <a href="${issueUrl}">${issueKey}</a></b> — ${issueSummary}`, tenantId);
    return { success: true };
  });

  // Creates a named group Teams chat with all provided emails, posts an issue link
  resolver.define('startGroupChat', async ({ payload }) => {
    const { emails, issueKey, issueSummary } = payload;
    const tenantId = await getMsTenantId();
    const userIds = await Promise.all(emails.map(e => getAadUserId(e.trim(), tenantId)));
    const chat = await graphPost('/v1.0/chats', {
      chatType: 'group',
      topic: `${issueKey}: ${issueSummary.slice(0, 60)}`,
      members: userIds.map(id => ({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${id}`,
      })),
    }, tenantId);
    const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
    await postChatMessage(chat.id, `<b>Jira: <a href="${issueUrl}">${issueKey}</a></b> — ${issueSummary}`, tenantId);
    return { success: true };
  });

  // Manually posts this issue to the saved Teams channel configuration
  resolver.define('postToChannelManual', async ({ payload }) => {
    const { issueKey, issueSummary } = payload;
    const config = await getGlobalConfig();
    if (!config?.teamId || !config?.channelId) throw new Error('No channel configured. Go to Teams Connector settings first.');
    const tenantId = await getMsTenantId();
    const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey}`;
    await graphPost(`/v1.0/teams/${config.teamId}/channels/${config.channelId}/messages`, {
      body: { contentType: 'html', content: `<b>Jira: <a href="${issueUrl}">${issueKey}</a></b> — ${issueSummary}` },
    }, tenantId);
    return { success: true };
  });
}

import { getValidAtlassianAuth, buildAuthorizationUrl, ATLASSIAN_REDIRECT_URI, atlassianFetch } from '../jira/atlassianAuth.js';
import { getIssueNotifySub, saveIssueNotifySub } from '../storage/kvsStore.js';
import { buildConnectCard } from './cards.js';
import { extractTextFromADF, commentBody } from '../jira/utils.js';

function taskContinue(title, card, height = 'medium', width = 'medium') {
  return {
    statusCode: 200,
    headers: { 'Content-Type': ['application/json'] },
    body: JSON.stringify({ task: { type: 'continue', value: { title, height, width, card: { contentType: 'application/vnd.microsoft.card.adaptive', content: card } } } }),
  };
}

function taskMessage(value) {
  return { statusCode: 200, headers: { 'Content-Type': ['application/json'] }, body: JSON.stringify({ task: { type: 'message', value } }) };
}

// Unlike a proactive bot message (which silently fails in any conversation the bot isn't a
// member of — see composeExtension.js), this replies with the connect link directly in the
// task-dialog response itself, so it always reaches the user regardless of bot conversation
// membership. Returns the auth record if connected, otherwise sends back a task/message
// response the caller should return immediately.
async function checkConnection(teamsUserId) {
  const auth = await getValidAtlassianAuth(teamsUserId);
  if (auth?.cloudId) return { auth };
  const url = buildAuthorizationUrl(ATLASSIAN_REDIRECT_URI, teamsUserId);
  return { response: taskContinue('Connect your Jira account', buildConnectCard(url, 'Connect your Jira account', 'Connect first, then try this action again.'), 'small') };
}

// ── task/fetch — builds the dialog shown when Comment / Edit / Notify is clicked on an issueCard() ──
export async function handleTaskFetch(body) {
  const data = body.value?.data || {};
  const { cardAction, issueKey } = data;
  // aadObjectId, not from.id — from.id is scoped per-conversation in Teams, aadObjectId is stable
  // across personal chat / channel / group chat for the same person.
  const teamsUserId = body.from?.aadObjectId || body.from?.id;
  console.log('[taskFetch]', cardAction, issueKey, '| from:', teamsUserId);

  if (cardAction === 'comment') {
    return taskContinue(`Comment on ${issueKey}`, {
      type: 'AdaptiveCard', version: '1.2',
      body: [
        { type: 'TextBlock', text: `Add a comment to ${issueKey}`, weight: 'Bolder', size: 'Medium' },
        { type: 'Input.Text', id: 'commentText', placeholder: 'Your comment…', isMultiline: true, isRequired: true },
      ],
      actions: [{ type: 'Action.Submit', title: 'Post Comment', data: { cardAction: 'comment', issueKey } }],
    });
  }

  if (cardAction === 'edit') {
    const conn = await checkConnection(teamsUserId);
    if (conn.response) return conn.response;

    // Pulled live from this Jira site — same idea as the official Jira app's edit dialog,
    // rather than a fixed guess at what priorities/assignees might exist. Labels don't get
    // the same "pick from every label on the site" treatment — sites with a lot of labels in
    // use turned that into an unreadably long checkbox wall, so it's a plain field again.
    const [issueRes, transRes, priorityRes, assignableRes] = await Promise.all([
      atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}?fields=summary,description,status,priority,assignee,duedate,labels,created,updated,reporter`),
      atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/transitions`),
      atlassianFetch(teamsUserId, `/rest/api/3/priority`),
      atlassianFetch(teamsUserId, `/rest/api/3/user/assignable/search?issueKey=${issueKey}&maxResults=50`),
    ]);
    const issue        = await issueRes.json();
    const transitions   = await transRes.json();
    const priorities    = await priorityRes.json().catch(() => []);
    const assignable     = await assignableRes.json().catch(() => []);
    const f = issue.fields || {};

    const statusChoices = (transitions.transitions || []).map(t => ({ title: t.name, value: t.name }));
    if (!statusChoices.some(c => c.value === f.status?.name)) {
      statusChoices.unshift({ title: f.status?.name || 'Unknown', value: f.status?.name || '' });
    }

    const priorityChoices = (Array.isArray(priorities) ? priorities : []).map(p => ({ title: p.name, value: p.name }));
    if (!priorityChoices.length) {
      priorityChoices.push(...['Highest', 'High', 'Medium', 'Low', 'Lowest'].map(v => ({ title: v, value: v })));
    }

    const assigneeChoices = [
      { title: '— Unassigned —', value: '' },
      ...(Array.isArray(assignable) ? assignable : []).map(u => ({ title: u.displayName, value: u.accountId })),
    ];
    if (f.assignee?.accountId && !assigneeChoices.some(c => c.value === f.assignee.accountId)) {
      assigneeChoices.push({ title: f.assignee.displayName || f.assignee.accountId, value: f.assignee.accountId });
    }

    // Own OAuth-connected siteUrl, not the JIRA_BASE_URL env var — guaranteed to match
    // whichever site this specific person is actually connected to.
    const jiraUrl = `${conn.auth.siteUrl}/browse/${issueKey}`;
    const fmt = iso => iso ? new Date(iso).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' }) : '—';

    // Two-column layout — left is what people actually edit (summary, description), right is
    // the field rail, loosely mirroring the official Jira app's edit dialog. Kept deliberately
    // short: no worklog section here (still available via the log command / issue panel) and
    // no full label picker (see above) — both made the dialog long and cluttered for no real gain.
    return taskContinue(`Edit ${issueKey}`, {
      type: 'AdaptiveCard', version: '1.2',
      body: [
        { type: 'TextBlock', text: issueKey, weight: 'Bolder', size: 'Medium', spacing: 'None' },
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column', width: 3,
              items: [
                { type: 'Input.Text', id: 'summary', label: 'Summary', isRequired: true, value: f.summary || '' },
                { type: 'Input.Text', id: 'description', label: 'Description', isMultiline: true, value: extractTextFromADF(f.description) },
                { type: 'Input.Text', id: 'labels', label: 'Labels (comma-separated)', value: (f.labels || []).join(', ') },
              ],
            },
            {
              type: 'Column', width: 2,
              items: [
                { type: 'Input.ChoiceSet', id: 'status', label: 'Status', value: f.status?.name || '', choices: statusChoices },
                { type: 'Input.ChoiceSet', id: 'assignee', label: 'Assignee', value: f.assignee?.accountId || '', choices: assigneeChoices },
                { type: 'Input.ChoiceSet', id: 'priority', label: 'Priority', value: f.priority?.name || '', choices: priorityChoices },
                { type: 'Input.Date', id: 'dueDate', label: 'Due date', value: f.duedate || '' },
                { type: 'TextBlock', text: `Created ${fmt(f.created)}`, isSubtle: true, size: 'Small', wrap: true, spacing: 'Medium', separator: true },
                { type: 'TextBlock', text: `Updated ${fmt(f.updated)}`, isSubtle: true, size: 'Small', wrap: true, spacing: 'None' },
                { type: 'TextBlock', text: `Reporter ${f.reporter?.displayName || 'Unknown'}`, isSubtle: true, size: 'Small', wrap: true, spacing: 'None' },
              ],
            },
          ],
        },
      ],
      actions: [
        { type: 'Action.Submit', title: 'Save Changes', data: { cardAction: 'edit', issueKey, originalStatus: f.status?.name || '' } },
        { type: 'Action.OpenUrl', title: 'Open in Jira', url: jiraUrl },
      ],
    }, 'large', 'medium');
  }

  if (cardAction === 'notify') {
    const sub = (await getIssueNotifySub(issueKey)) || { notifyOnUpdate: true, notifyOnComment: true };
    return taskContinue('Channel Notifications', {
      type: 'AdaptiveCard', version: '1.2',
      body: [
        { type: 'TextBlock', text: 'Channel Notifications', weight: 'Bolder', size: 'Medium' },
        { type: 'TextBlock', text: 'Customize the type of notifications this issue sends to your configured Teams channel.', isSubtle: true, wrap: true },
        { type: 'TextBlock', text: `For work items that match: issueKey = ${issueKey}`, wrap: true, spacing: 'Medium' },
        { type: 'TextBlock', text: 'Send a message to the channel when:', weight: 'Bolder', spacing: 'Medium' },
        { type: 'Input.Toggle', id: 'notifyOnUpdate', title: 'Work item is updated', value: sub.notifyOnUpdate ? 'true' : 'false' },
        { type: 'Input.Toggle', id: 'notifyOnComment', title: 'Comment is added', value: sub.notifyOnComment ? 'true' : 'false' },
      ],
      actions: [{ type: 'Action.Submit', title: 'Save notifications', data: { cardAction: 'notify', issueKey } }],
    });
  }

  return taskMessage(`Unknown action: ${cardAction}`);
}

// ── task/submit — handle whichever dialog's submit button was clicked ──
export async function handleTaskSubmit(body) {
  const data = body.value?.data || {};
  const { cardAction, issueKey } = data;
  const teamsUserId = body.from?.aadObjectId || body.from?.id;
  console.log('[taskSubmit]', cardAction, issueKey, '| from:', teamsUserId);

  if (cardAction === 'comment') {
    const conn = await checkConnection(teamsUserId);
    if (conn.response) return conn.response;
    const commentText = (data.commentText || '').trim();
    if (!commentText) return taskMessage('❌ Comment cannot be empty.');
    const res = await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/comment`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(commentBody(commentText)),
    });
    const commented = await res.json();
    return taskMessage(commented.id ? `✅ Comment added to ${issueKey}` : `❌ Failed to comment: ${JSON.stringify(commented.errors || commented)}`);
  }

  if (cardAction === 'edit') {
    const conn = await checkConnection(teamsUserId);
    if (conn.response) return conn.response;

    const { status, priority, assignee, dueDate, labels, summary, description, originalStatus } = data;

    // Status change goes through the transitions endpoint, not a plain field PUT
    if (status && status !== originalStatus) {
      const transRes  = await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/transitions`);
      const transData = await transRes.json();
      const match = (transData.transitions || []).find(t => t.name === status);
      if (match) {
        await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/transitions`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ transition: { id: match.id } }),
        });
      }
    }

    const fields = {};
    if (summary?.trim()) fields.summary = summary.trim();
    if (description !== undefined) {
      const text = description.trim();
      fields.description = { type: 'doc', version: 1, content: [{ type: 'paragraph', content: text ? [{ type: 'text', text }] : [] }] };
    }
    if (priority) fields.priority = { name: priority };
    if (dueDate !== undefined) fields.duedate = dueDate || null;
    if (labels !== undefined) fields.labels = labels.split(',').map(s => s.trim()).filter(Boolean);
    // The picker's accountId choice replaces the old email-search lookup entirely, so a
    // typo'd email can no longer silently fail to assign anyone.
    if (assignee !== undefined) fields.assignee = assignee ? { accountId: assignee } : null;

    if (Object.keys(fields).length) {
      await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields }),
      });
    }

    return taskMessage(`✅ ${issueKey} updated`);
  }

  if (cardAction === 'notify') {
    await saveIssueNotifySub(issueKey, {
      notifyOnUpdate: data.notifyOnUpdate === 'true',
      notifyOnComment: data.notifyOnComment === 'true',
    });
    return taskMessage(`🔔 Notification settings saved for ${issueKey}`);
  }

  return taskMessage(`Unknown action: ${cardAction}`);
}

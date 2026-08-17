import { getValidAtlassianAuth, buildAuthorizationUrl, ATLASSIAN_REDIRECT_URI, atlassianFetch } from '../jira/atlassianAuth.js';
import { getIssueNotifySub, saveIssueNotifySub } from '../storage/kvsStore.js';

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
  return { response: taskMessage(`Connect your Jira account first, then try again:\n\n${url}`) };
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

    const [issueRes, transRes] = await Promise.all([
      atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}?fields=summary,status,priority,assignee,duedate,labels`),
      atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/transitions`),
    ]);
    const issue = await issueRes.json();
    const transitions = await transRes.json();
    const f = issue.fields || {};

    const statusChoices = (transitions.transitions || []).map(t => ({ title: t.name, value: t.name }));
    if (!statusChoices.some(c => c.value === f.status?.name)) {
      statusChoices.unshift({ title: f.status?.name || 'Unknown', value: f.status?.name || '' });
    }

    return taskContinue(`Edit ${issueKey}`, {
      type: 'AdaptiveCard', version: '1.2',
      body: [
        { type: 'TextBlock', text: `${issueKey}: ${f.summary || ''}`, weight: 'Bolder', wrap: true },
        { type: 'Input.ChoiceSet', id: 'status', label: 'Status', value: f.status?.name || '', choices: statusChoices },
        {
          type: 'Input.ChoiceSet', id: 'priority', label: 'Priority', value: f.priority?.name || 'Medium',
          choices: [
            { title: 'Highest', value: 'Highest' }, { title: 'High', value: 'High' },
            { title: 'Medium', value: 'Medium' }, { title: 'Low', value: 'Low' }, { title: 'Lowest', value: 'Lowest' },
          ],
        },
        { type: 'Input.Text', id: 'assigneeEmail', label: 'Assignee email (leave blank to keep unchanged)', placeholder: f.assignee?.emailAddress || 'unassigned' },
        { type: 'Input.Date', id: 'dueDate', label: 'Due date', value: f.duedate || '' },
        { type: 'Input.Text', id: 'labels', label: 'Labels (comma-separated)', value: (f.labels || []).join(', ') },
      ],
      actions: [{ type: 'Action.Submit', title: 'Save Changes', data: { cardAction: 'edit', issueKey, originalStatus: f.status?.name || '' } }],
    });
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
      body: JSON.stringify({
        body: { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: commentText }] }] },
      }),
    });
    const commented = await res.json();
    return taskMessage(commented.id ? `✅ Comment added to ${issueKey}` : `❌ Failed to comment: ${JSON.stringify(commented.errors || commented)}`);
  }

  if (cardAction === 'edit') {
    const conn = await checkConnection(teamsUserId);
    if (conn.response) return conn.response;

    const { status, priority, assigneeEmail, dueDate, labels, originalStatus } = data;

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
    if (priority) fields.priority = { name: priority };
    if (dueDate !== undefined) fields.duedate = dueDate || null;
    if (labels !== undefined) fields.labels = labels.split(',').map(s => s.trim()).filter(Boolean);

    if (assigneeEmail?.trim()) {
      const userRes = await atlassianFetch(teamsUserId, `/rest/api/3/user/search?query=${encodeURIComponent(assigneeEmail.trim())}&maxResults=1`);
      const users   = await userRes.json();
      const user    = Array.isArray(users) ? users[0] : null;
      if (user) fields.assignee = { accountId: user.accountId };
    }

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

import api, { route } from '@forge/api';
import { setBotDebugLog } from '../storage/kvsStore.js';
import { issueCard, buildCreateIssueCard, buildCommentCard, buildLogTimeCard } from './cards.js';
import { parseTimeToSeconds } from '../jira/utils.js';

// Teams sends composeExtension/queryLink when a Jira URL is pasted — link unfurling.
export async function handleQueryLink(body) {
  const url = body.value?.url || '';
  console.log('[composeExtension] link unfurl URL:', url);

  // Extract issue key from Jira URL — e.g. /browse/PROJ-123
  const match = url.match(/\/browse\/([A-Z][A-Z0-9_]*-\d+)/i);
  if (!match) {
    return { statusCode: 200, body: JSON.stringify({ composeExtension: { type: 'result', attachmentLayout: 'list', attachments: [] } }) };
  }

  const issueKey = match[1].toUpperCase();
  console.log('[composeExtension] fetching issue:', issueKey);

  let issue, f;
  try {
    const jiraRes = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}?fields=summary,status,priority,issuetype,assignee,reporter,description`);
    issue = await jiraRes.json();
    console.log('[composeExtension] jira response status:', jiraRes.status);
    if (issue.errorMessages || issue.errors) {
      console.log('[composeExtension] jira error:', JSON.stringify(issue));
      await setBotDebugLog({ error: 'Jira fetch failed', detail: JSON.stringify(issue), issueKey });
      return { statusCode: 200, body: JSON.stringify({ composeExtension: { type: 'result', attachmentLayout: 'list', attachments: [] } }) };
    }
  } catch (jiraErr) {
    console.log('[composeExtension] jira fetch error:', jiraErr.message);
    await setBotDebugLog({ error: jiraErr.message, issueKey, step: 'jira-fetch' });
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

// Message actions: fetchTask — route by commandId, return the adaptive-card form.
export async function handleFetchTask(body) {
  const commandId  = body.value?.commandId || 'createJiraIssue';
  const rawContent = body.value?.messagePayload?.body?.content || '';
  const messageText = rawContent.replace(/<[^>]+>/g, '').trim().slice(0, 500);
  console.log('[composeExtension] fetchTask commandId:', commandId, '| text preview:', messageText.slice(0, 60));

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

// Message actions: submitAction — route by commandId, perform the Jira write.
export async function handleSubmitAction(body) {
  const commandId = body.value?.commandId || body.data?.commandId || 'createJiraIssue';
  const data = body.data || body.value?.data || {};
  console.log('[composeExtension] submitAction commandId:', commandId, '| data:', JSON.stringify(data));

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
  console.log('[composeExtension] submitAction — project:', projectKey, '| summary:', summary);

  if (!projectKey || !summary) {
    return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: '❌ Project Key and Summary are required.' } }) };
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
  console.log('[composeExtension] submitAction — created key:', created.key);

  if (created.key) {
    const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${created.key}`;
    return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: `✅ **${created.key}** created — [Open in Jira](${issueUrl})` } }) };
  }
  const errDetail = JSON.stringify(created.errors || created.errorMessages || created);
  return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value: `❌ Failed to create issue: ${errDetail}` } }) };
}

// Compose extension search: "Search Jira Issues" from message bar.
export async function handleQuery(body) {
  const commandId   = body.value?.commandId || '';
  const searchQuery = (body.value?.parameters?.[0]?.value || '').trim();
  console.log('[composeExtension] query commandId:', commandId, '| query:', searchQuery);

  const jql = searchQuery
    ? `text ~ "${searchQuery.replace(/"/g, '\\"')}" ORDER BY updated DESC`
    : 'ORDER BY updated DESC';

  const searchRes  = await api.asApp().requestJira(route`/rest/api/3/search?jql=${jql}&maxResults=8&fields=summary,status,priority,issuetype,assignee`);
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
    body: JSON.stringify({ composeExtension: { type: 'result', attachmentLayout: 'list', attachments } }),
  };
}

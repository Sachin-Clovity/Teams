import api, { route } from '@forge/api';
import { setBotDebugLog } from '../../storage/kvsStore.js';

const NO_RESULT = { statusCode: 200, body: JSON.stringify({ composeExtension: { type: 'result', attachmentLayout: 'list', attachments: [] } }) };

// Teams sends composeExtension/queryLink when a Jira URL is pasted — link unfurling.
export async function handleQueryLink(body) {
  const url = body.value?.url || '';
  console.log('[composeExtension] link unfurl URL:', url);

  // Extract issue key from Jira URL — e.g. /browse/PROJ-123
  const match = url.match(/\/browse\/([A-Z][A-Z0-9_]*-\d+)/i);
  if (!match) return NO_RESULT;

  const issueKey = match[1].toUpperCase();
  console.log('[composeExtension] fetching issue:', issueKey);

  let issue;
  try {
    const jiraRes = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}?fields=summary,status,priority,issuetype,assignee,reporter,description`);
    issue = await jiraRes.json();
    console.log('[composeExtension] jira response status:', jiraRes.status);
    if (issue.errorMessages || issue.errors) {
      console.log('[composeExtension] jira error:', JSON.stringify(issue));
      await setBotDebugLog({ error: 'Jira fetch failed', detail: JSON.stringify(issue), issueKey });
      return NO_RESULT;
    }
  } catch (jiraErr) {
    console.log('[composeExtension] jira fetch error:', jiraErr.message);
    await setBotDebugLog({ error: jiraErr.message, issueKey, step: 'jira-fetch' });
    return NO_RESULT;
  }
  const f = issue.fields || {};

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

  return { statusCode: 200, body: JSON.stringify(response), headers: { 'Content-Type': ['application/json'] } };
}

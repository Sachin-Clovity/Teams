import api, { route } from '@forge/api';
import { issueCard } from '../cards.js';

// Compose extension search: "Search Jira Issues" from message bar.
export async function handleQuery(body) {
  const commandId   = body.value?.commandId || '';
  const searchQuery = (body.value?.parameters?.[0]?.value || '').trim();
  console.log('[composeExtension] query commandId:', commandId, '| query:', searchQuery);

  const jql = searchQuery
    ? `text ~ "${searchQuery.replace(/"/g, '\\"')}" ORDER BY updated DESC`
    : 'updated >= "1970/01/01" ORDER BY updated DESC';

  const searchRes  = await api.asApp().requestJira(route`/rest/api/3/search/jql?jql=${jql}&maxResults=8&fields=summary,status,priority,issuetype,assignee`);
  const searchData = await searchRes.json();
  const issues = searchData.issues || [];

  const attachments = issues.map(i => {
    const f   = i.fields || {};
    const url = `${process.env.JIRA_BASE_URL}/browse/${i.key}`;
    return {
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: issueCard(i.key, f.summary, f.issuetype?.name, f.status?.name, f.priority?.name, f.assignee?.displayName || 'Unassigned', url, `${f.issuetype?.name || 'Issue'} · ${f.status?.name || ''}`, f.issuetype?.iconUrl),
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
    headers: { 'Content-Type': ['application/json'] },
    body: JSON.stringify({ composeExtension: { type: 'result', attachmentLayout: 'list', attachments } }),
  };
}

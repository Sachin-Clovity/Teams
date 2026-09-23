import api, { route } from '@forge/api';
import { sendBotReply } from '../../graph/botReply.js';
import { issueCard } from '../cards.js';

// SEARCH: "search payment bug" or a raw JQL fragment like "search assignee = currentUser()"
export async function handleSearch(body, match) {
  const query = match[1].trim();
  const jql   = query.includes('=') || query.includes('ORDER') ? query : `text ~ "${query.replace(/"/g, '\\"')}" ORDER BY updated DESC`;
  const searchRes = await api.asApp().requestJira(route`/rest/api/3/search/jql?jql=${jql}&maxResults=5&fields=summary,status,priority,issuetype,assignee`);
  const data   = await searchRes.json();
  const issues = data.issues || [];
  if (!issues.length) return sendBotReply(body, [], `No issues found for: **${query}**`);
  const cards = issues.map(i => {
    const f = i.fields || {};
    return issueCard(i.key, f.summary, f.issuetype?.name, f.status?.name, f.priority?.name, f.assignee?.displayName || 'Unassigned', `${process.env.JIRA_BASE_URL}/browse/${i.key}`, '🔍 Search Result', f.issuetype?.iconUrl);
  });
  return sendBotReply(body, cards, `Found ${issues.length} issue(s):`);
}

// PROJECTS: "projects" / "project" / "project list" — list all Jira projects
export async function handleProjects(body) {
  const projRes  = await api.asApp().requestJira(route`/rest/api/3/project/search?maxResults=20&orderBy=name`);
  const projData = await projRes.json();
  const projects = projData.values || [];
  if (!projects.length) return sendBotReply(body, [], `No projects found.`);
  const list = projects.map(p => `• **${p.key}** — ${p.name}`).join('\n');
  return sendBotReply(body, [], `**Jira Projects (${projData.total || projects.length}):**\n\n${list}`);
}

// SHOW: "show AITEST-40" — fetch and display issue details
export async function handleShow(body, match) {
  const issueKey = match[1].toUpperCase();
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
  return sendBotReply(body, [issueCard(issueKey, summary, type, status, priority, assignee, issueUrl, `📋 ${reporter} reported`, f.issuetype?.iconUrl)]);
}

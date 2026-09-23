import { sendBotReply } from '../../graph/botReply.js';
import { issueCard } from '../cards.js';
import { atlassianFetch } from '../../jira/atlassianAuth.js';

const TYPE_EMOJI = { Bug: '🐛', Task: '✅', Story: '📖', Epic: '⚡' };

// BUG / TASK / STORY / EPIC — typed issue creation: "bug PROJ summary text"
export async function handleTypedCreate(body, teamsUserId, match) {
  const issueType  = match[1].charAt(0).toUpperCase() + match[1].slice(1).toLowerCase();
  const projectKey = match[2].toUpperCase();
  const summary    = match[3].trim();
  const createRes  = await atlassianFetch(teamsUserId, '/rest/api/3/issue', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields: { project: { key: projectKey }, summary, issuetype: { name: issueType } } })
  });
  const created = await createRes.json();
  if (created.key) {
    const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${created.key}`;
    return sendBotReply(body, [issueCard(created.key, summary, issueType, 'To Do', 'None', 'Unassigned', issueUrl, `${TYPE_EMOJI[issueType] || '🆕'} ${issueType} Created`)]);
  }
  return sendBotReply(body, [], `❌ Failed to create ${issueType}: ${JSON.stringify(created.errors || created)}`);
}

// CREATE: "create PROJ Summary of the issue" — auto-picks a sensible issue type for the project
export async function handleCreate(body, teamsUserId, match) {
  const projectKey = match[1].toUpperCase();
  const summary    = match[2].trim();

  const metaRes   = await atlassianFetch(teamsUserId, `/rest/api/3/issue/createmeta?projectKeys=${projectKey}&expand=projects.issuetypes`);
  const metaData  = await metaRes.json();
  const issueTypes = metaData.projects?.[0]?.issuetypes || [];
  const preferred  = ['Task', 'Story', 'Bug', 'Subtask'];
  const issueType  = issueTypes.find(t => preferred.includes(t.name))?.name || issueTypes[0]?.name || 'Task';

  const createRes = await atlassianFetch(teamsUserId, '/rest/api/3/issue', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields: { project: { key: projectKey }, summary, issuetype: { name: issueType } } })
  });
  const created = await createRes.json();
  if (created.key) {
    const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${created.key}`;
    return sendBotReply(body, [issueCard(created.key, summary, issueType, 'To Do', 'None', 'Unassigned', issueUrl, '✅ Issue Created')]);
  }
  return sendBotReply(body, [], `❌ Failed to create issue: ${JSON.stringify(created.errors || created)}`);
}

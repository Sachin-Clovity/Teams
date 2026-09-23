import api, { route } from '@forge/api';
import { sendBotReply } from '../graph/botReply.js';
import { issueCard, helpText, buildConnectCard } from './cards.js';
import { parseTimeToSeconds, commentBody } from '../jira/utils.js';
import { buildAuthorizationUrl, ATLASSIAN_REDIRECT_URI, getValidAtlassianAuth, atlassianFetch, requireConnection } from '../jira/atlassianAuth.js';
import { clearAtlassianAuth } from '../storage/kvsStore.js';

// Text commands typed directly to the bot: connect / create / bug / task / story / epic /
// search / projects / show / update / comment / log / assign / help.
export async function handleBotMessage(body) {
  const text = (body.text || '').replace(/<at>[^<]*<\/at>/gi, '').trim();
  // from.id is scoped to the current conversation (personal chat vs channel vs group chat
  // each get a different value for the same person) — aadObjectId is the stable identity
  // that's the same everywhere, which is what lets `connect` (personal chat) and the
  // message-action / card buttons (channel or group chat) find the same saved auth record.
  const teamsUserId = body.from?.aadObjectId || body.from?.id;
  console.log('[botCommand] message command:', text, '| from:', teamsUserId, '| body.from:', JSON.stringify(body.from));

  try {
    return await dispatchCommand(body, text, teamsUserId);
  } catch (e) {
    // Without this, a Jira outage/rate-limit (a non-JSON error body making res.json() throw,
    // for example) propagates uncaught to teamsBotHandler's outer catch, which returns a
    // raw 500 — the user who typed the command gets no reply at all, not even an error message.
    console.log('[botCommand] unhandled error:', e.message);
    return sendBotReply(body, [], `❌ Something went wrong talking to Jira. Please try again in a moment.`);
  }
}

async function dispatchCommand(body, text, teamsUserId) {
  // CONNECT: link this Teams user to their real Jira account
  if (/^connect$/i.test(text)) {
    console.log('[botCommand] connect — teamsUserId used for lookup/save:', teamsUserId);
    const existing = await getValidAtlassianAuth(teamsUserId);
    if (existing?.cloudId) return sendBotReply(body, [], `✅ Already connected to **${existing.siteName}** (${existing.siteUrl}). Type \`disconnect\` to unlink.`);
    const url = buildAuthorizationUrl(ATLASSIAN_REDIRECT_URI, teamsUserId);
    return sendBotReply(body, [buildConnectCard(url)], undefined);
  }

  // DISCONNECT: unlink this Teams user's Jira account
  if (/^disconnect$/i.test(text)) {
    await clearAtlassianAuth(teamsUserId);
    return sendBotReply(body, [], `Disconnected. Type \`connect\` to link a Jira account again.`);
  }

  // CONSENT: hand a Teams/Azure admin the one-time admin-consent link their organization needs
  // to approve before anyone there can use this bot — only relevant once installed outside
  // Clovity's own tenant, but harmless (and useful) to have ready regardless.
  if (/^consent$/i.test(text)) {
    const url = `https://login.microsoftonline.com/organizations/adminconsent?client_id=${process.env.MS_CLIENT_ID}`;
    return sendBotReply(body, [buildConnectCard(
      url,
      'Admin approval needed',
      'An admin for your organization needs to approve this app once before everyone here can use it.',
      '✅ Approve for organization'
    )], undefined);
  }

  // BUG / TASK / STORY / EPIC — typed issue creation
  const typedMatch = text.match(/^(bug|task|story|epic)\s+([A-Z][A-Z0-9_]*)\s+(.+)$/i);
  if (typedMatch) {
    if (!(await requireConnection(body, teamsUserId))) return { statusCode: 200, body: JSON.stringify({}) };
    const issueType  = typedMatch[1].charAt(0).toUpperCase() + typedMatch[1].slice(1).toLowerCase();
    const projectKey = typedMatch[2].toUpperCase();
    const summary    = typedMatch[3].trim();
    const emojiMap   = { Bug: '🐛', Task: '✅', Story: '📖', Epic: '⚡' };
    const createRes  = await atlassianFetch(teamsUserId, '/rest/api/3/issue', {
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
    if (!(await requireConnection(body, teamsUserId))) return { statusCode: 200, body: JSON.stringify({}) };
    const projectKey = createMatch[1].toUpperCase();
    const summary    = createMatch[2].trim();

    // Fetch valid issue types for this project
    const metaRes   = await atlassianFetch(teamsUserId, `/rest/api/3/issue/createmeta?projectKeys=${projectKey}&expand=projects.issuetypes`);
    const metaData  = await metaRes.json();
    const issueTypes = metaData.projects?.[0]?.issuetypes || [];
    const preferred  = ['Task', 'Story', 'Bug', 'Subtask'];
    const issueType  = issueTypes.find(t => preferred.includes(t.name))?.name || issueTypes[0]?.name || 'Task';

    const createRes  = await atlassianFetch(teamsUserId, '/rest/api/3/issue', {
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
    const jql   = query.includes('=') || query.includes('ORDER') ? query : `text ~ "${query.replace(/"/g, '\\"')}" ORDER BY updated DESC`;
    const searchRes = await api.asApp().requestJira(route`/rest/api/3/search/jql?jql=${jql}&maxResults=5&fields=summary,status,priority,issuetype,assignee`);
    const data  = await searchRes.json();
    const issues = (data.issues || []);
    if (!issues.length) return sendBotReply(body, [], `No issues found for: **${query}**`);
    const cards = issues.map(i => {
      const f = i.fields || {};
      return issueCard(i.key, f.summary, f.issuetype?.name, f.status?.name, f.priority?.name, f.assignee?.displayName || 'Unassigned', `${process.env.JIRA_BASE_URL}/browse/${i.key}`, '🔍 Search Result', f.issuetype?.iconUrl);
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
    return sendBotReply(body, [issueCard(issueKey, summary, type, status, priority, assignee, issueUrl, `📋 ${reporter} reported`, f.issuetype?.iconUrl)]);
  }

  // UPDATE: "update PROJ-123 Done" (transition by status name)
  const updateMatch = text.match(/^update\s+([A-Z][A-Z0-9_]*-\d+)\s+(.+)$/i);
  if (updateMatch) {
    if (!(await requireConnection(body, teamsUserId))) return { statusCode: 200, body: JSON.stringify({}) };
    const issueKey  = updateMatch[1].toUpperCase();
    const newStatus = updateMatch[2].trim();
    const transRes  = await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/transitions`);
    const transData = await transRes.json();
    const transition = (transData.transitions || []).find(t => t.name.toLowerCase() === newStatus.toLowerCase());
    if (!transition) {
      const available = (transData.transitions || []).map(t => t.name).join(', ');
      return sendBotReply(body, [], `❌ Status **"${newStatus}"** not found. Available: ${available}`);
    }
    await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/transitions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ transition: { id: transition.id } })
    });
    return sendBotReply(body, [], `✅ **${issueKey}** updated to **${transition.name}**`);
  }

  // COMMENT: "comment PROJ-123 This is a comment"
  const commentMatch = text.match(/^comment\s+([A-Z][A-Z0-9_]*-\d+)\s+(.+)$/i);
  if (commentMatch) {
    if (!(await requireConnection(body, teamsUserId))) return { statusCode: 200, body: JSON.stringify({}) };
    const issueKey   = commentMatch[1].toUpperCase();
    const commentTxt = commentMatch[2].trim();
    const commentRes = await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/comment`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(commentBody(commentTxt)),
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
    if (!(await requireConnection(body, teamsUserId))) return { statusCode: 200, body: JSON.stringify({}) };
    const issueKey   = logMatch[1].toUpperCase();
    const timeStr    = logMatch[2].trim();
    const logNote    = logMatch[3].trim();
    const seconds    = parseTimeToSeconds(timeStr);
    if (!seconds) return sendBotReply(body, [], `❌ Invalid time format. Use: 2h, 30m, 1h 30m`);
    const worklogBody = { timeSpentSeconds: seconds };
    if (logNote) {
      worklogBody.comment = { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: logNote }] }] };
    }
    const logRes = await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/worklog`, {
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
    if (!(await requireConnection(body, teamsUserId))) return { statusCode: 200, body: JSON.stringify({}) };
    const issueKey = assignMatch[1].toUpperCase();
    const email    = assignMatch[2].trim();
    const userRes  = await atlassianFetch(teamsUserId, `/rest/api/3/user/search?query=${email}&maxResults=1`);
    const users    = await userRes.json();
    const user     = Array.isArray(users) ? users[0] : null;
    if (!user) return sendBotReply(body, [], `❌ User **${email}** not found in Jira.`);
    await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}`, {
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

import { sendBotReply } from '../../graph/botReply.js';
import { parseTimeToSeconds, commentBody } from '../../jira/utils.js';
import { atlassianFetch } from '../../jira/atlassianAuth.js';

// UPDATE: "update PROJ-123 Done" (transition by status name)
export async function handleUpdate(body, teamsUserId, match) {
  const issueKey  = match[1].toUpperCase();
  const newStatus = match[2].trim();
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
export async function handleComment(body, teamsUserId, match) {
  const issueKey   = match[1].toUpperCase();
  const commentTxt = match[2].trim();
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
export async function handleLog(body, teamsUserId, match) {
  const issueKey = match[1].toUpperCase();
  const timeStr   = match[2].trim();
  const logNote   = match[3].trim();
  const seconds   = parseTimeToSeconds(timeStr);
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
export async function handleAssign(body, teamsUserId, match) {
  const issueKey = match[1].toUpperCase();
  const email    = match[2].trim();
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

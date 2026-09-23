import api, { route } from '@forge/api';
import { parseTimeToSeconds, commentBody } from '../../jira/utils.js';
import { handleWizardStep } from './wizard.js';
import { taskMessage } from './helpers.js';

async function commentInJira(data) {
  const { issueKey, commentText } = data;
  if (!issueKey || !commentText) return taskMessage('❌ Issue Key and Comment are required.');
  const commentRes = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey.trim().toUpperCase()}/comment`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(commentBody(commentText.trim())),
  });
  const commented = await commentRes.json();
  if (commented.id) {
    const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${issueKey.trim().toUpperCase()}`;
    return taskMessage(`✅ Comment added to ${issueKey.trim().toUpperCase()}\n\n${issueUrl}`);
  }
  return taskMessage(`❌ Failed to add comment: ${JSON.stringify(commented.errors || commented)}`);
}

async function logTimeInJira(data) {
  const { issueKey, timeSpent, workDescription } = data;
  if (!issueKey || !timeSpent) return taskMessage('❌ Issue Key and Time Spent are required. e.g. 2h, 30m, 1h 30m');
  const seconds = parseTimeToSeconds(timeSpent);
  if (!seconds) return taskMessage('❌ Invalid time format. Use: 2h, 30m, 1h 30m, 90m');

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
    return taskMessage(`✅ Logged ${timeSpent} on ${issueKey.trim().toUpperCase()}\n\n${issueUrl}`);
  }
  return taskMessage(`❌ Failed to log time: ${JSON.stringify(logged.errors || logged)}`);
}

// Message actions: submitAction — route by commandId, perform the Jira write.
export async function handleSubmitAction(body) {
  const commandId = body.value?.commandId || body.data?.commandId || 'createJiraIssue';
  const data = body.data || body.value?.data || {};
  console.log('[composeExtension] submitAction commandId:', commandId, '| data:', JSON.stringify(data));

  if (commandId === 'commentInJira') return commentInJira(data);
  if (commandId === 'logTimeInJira') return logTimeInJira(data);

  // Create a work item — 3/4-step wizard: Site → Project → Type → Details.
  // Use aadObjectId, not from.id — from.id is conversation-scoped in Teams (personal chat vs
  // channel/group chat give the same person different ids), which is why `connect` and this
  // wizard could disagree on "am I connected". aadObjectId is stable across all of them.
  const teamsUserId = body.from?.aadObjectId || body.from?.id;
  return handleWizardStep(body, data, teamsUserId);
}

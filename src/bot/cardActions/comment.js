import { atlassianFetch } from '../../jira/atlassianAuth.js';
import { commentBody } from '../../jira/utils.js';
import { taskContinue, taskMessage, checkConnection } from './helpers.js';

export function fetchComment(issueKey) {
  return taskContinue(`Comment on ${issueKey}`, {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: `Add a comment to ${issueKey}`, weight: 'Bolder', size: 'Medium' },
      { type: 'Input.Text', id: 'commentText', placeholder: 'Your comment…', isMultiline: true, isRequired: true },
    ],
    actions: [{ type: 'Action.Submit', title: 'Post Comment', data: { cardAction: 'comment', issueKey } }],
  });
}

export async function submitComment(teamsUserId, issueKey, data) {
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

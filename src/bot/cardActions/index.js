import { taskMessage } from './helpers.js';
import { fetchComment, submitComment } from './comment.js';
import { fetchEdit, submitEdit } from './edit.js';
import { fetchNotify, submitNotify } from './notify.js';

// ── task/fetch — builds the dialog shown when Comment / Edit / Notify is clicked on an issueCard() ──
export async function handleTaskFetch(body) {
  const data = body.value?.data || {};
  const { cardAction, issueKey } = data;
  // aadObjectId, not from.id — from.id is scoped per-conversation in Teams, aadObjectId is stable
  // across personal chat / channel / group chat for the same person.
  const teamsUserId = body.from?.aadObjectId || body.from?.id;
  console.log('[taskFetch]', cardAction, issueKey, '| from:', teamsUserId);

  if (cardAction === 'comment') return fetchComment(issueKey);
  if (cardAction === 'edit')    return fetchEdit(teamsUserId, issueKey);
  if (cardAction === 'notify')  return fetchNotify(issueKey);
  return taskMessage(`Unknown action: ${cardAction}`);
}

// ── task/submit — handle whichever dialog's submit button was clicked ──
export async function handleTaskSubmit(body) {
  const data = body.value?.data || {};
  const { cardAction, issueKey } = data;
  const teamsUserId = body.from?.aadObjectId || body.from?.id;
  console.log('[taskSubmit]', cardAction, issueKey, '| from:', teamsUserId);

  if (cardAction === 'comment') return submitComment(teamsUserId, issueKey, data);
  if (cardAction === 'edit')    return submitEdit(teamsUserId, issueKey, data);
  if (cardAction === 'notify')  return submitNotify(issueKey, data);
  return taskMessage(`Unknown action: ${cardAction}`);
}

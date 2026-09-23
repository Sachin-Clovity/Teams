import { buildAuthorizationUrl, ATLASSIAN_REDIRECT_URI, getValidAtlassianAuth } from '../../jira/atlassianAuth.js';
import { buildConnectCard } from '../cards.js';

// Every wizard step (site/project/type/details) replies the same shape — a task/continue
// carrying the next card, always titled "Create a work item".
export function continueWith(card) {
  return {
    statusCode: 200,
    headers: { 'Content-Type': ['application/json'] },
    body: JSON.stringify({ task: { type: 'continue', value: { title: 'Create a work item', height: 'medium', width: 'medium', card: { contentType: 'application/vnd.microsoft.card.adaptive', content: card } } } }),
  };
}

// Shared by every wizard step — if the connection lapses (token expired, user disconnected)
// partway through, this stops the step with a clear reconnect prompt instead of atlassianFetch()
// throwing NOT_CONNECTED uncaught into a raw 500.
export async function requireConnectionOrPrompt(teamsUserId) {
  const auth = await getValidAtlassianAuth(teamsUserId);
  if (auth?.cloudId) return auth;
  const url = buildAuthorizationUrl(ATLASSIAN_REDIRECT_URI, teamsUserId);
  // task/message dialogs render plain text, not markdown — a [text](url) link here would show
  // up as literal unparsed characters. task/continue with a card gives a real clickable
  // "Sign in" button instead of a bare URL the user has to copy-paste.
  const card = buildConnectCard(url, 'Connect your Jira account', 'Connect first, then try creating the issue again.');
  return { errorResponse: { statusCode: 200, body: JSON.stringify({ task: { type: 'continue', value: { title: 'Connect your Jira account', height: 'small', width: 'medium', card: { contentType: 'application/vnd.microsoft.card.adaptive', content: card } } } }) } };
}

export function taskMessage(value) {
  return { statusCode: 200, body: JSON.stringify({ task: { type: 'message', value } }) };
}

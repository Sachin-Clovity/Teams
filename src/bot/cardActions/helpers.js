import { getValidAtlassianAuth, buildAuthorizationUrl, ATLASSIAN_REDIRECT_URI } from '../../jira/atlassianAuth.js';
import { buildConnectCard } from '../cards.js';

export function taskContinue(title, card, height = 'medium', width = 'medium') {
  return {
    statusCode: 200,
    headers: { 'Content-Type': ['application/json'] },
    body: JSON.stringify({ task: { type: 'continue', value: { title, height, width, card: { contentType: 'application/vnd.microsoft.card.adaptive', content: card } } } }),
  };
}

export function taskMessage(value) {
  return { statusCode: 200, headers: { 'Content-Type': ['application/json'] }, body: JSON.stringify({ task: { type: 'message', value } }) };
}

// Unlike a proactive bot message (which silently fails in any conversation the bot isn't a
// member of), this replies with the connect link directly in the task-dialog response itself,
// so it always reaches the user regardless of bot conversation membership. Returns the auth
// record if connected, otherwise sends back a task/message response the caller should return
// immediately.
export async function checkConnection(teamsUserId) {
  const auth = await getValidAtlassianAuth(teamsUserId);
  if (auth?.cloudId) return { auth };
  const url = buildAuthorizationUrl(ATLASSIAN_REDIRECT_URI, teamsUserId);
  return { response: taskContinue('Connect your Jira account', buildConnectCard(url, 'Connect your Jira account', 'Connect first, then try this action again.'), 'small') };
}

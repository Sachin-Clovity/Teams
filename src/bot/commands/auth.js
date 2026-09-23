import { sendBotReply } from '../../graph/botReply.js';
import { buildConnectCard } from '../cards.js';
import { buildAuthorizationUrl, ATLASSIAN_REDIRECT_URI, getValidAtlassianAuth } from '../../jira/atlassianAuth.js';
import { clearAtlassianAuth } from '../../storage/kvsStore.js';

// CONNECT: link this Teams user to their real Jira account
export async function handleConnect(body, teamsUserId) {
  console.log('[botCommand] connect — teamsUserId used for lookup/save:', teamsUserId);
  const existing = await getValidAtlassianAuth(teamsUserId);
  if (existing?.cloudId) return sendBotReply(body, [], `✅ Already connected to **${existing.siteName}** (${existing.siteUrl}). Type \`disconnect\` to unlink.`);
  const url = buildAuthorizationUrl(ATLASSIAN_REDIRECT_URI, teamsUserId);
  return sendBotReply(body, [buildConnectCard(url)], undefined);
}

// DISCONNECT: unlink this Teams user's Jira account
export async function handleDisconnect(body, teamsUserId) {
  await clearAtlassianAuth(teamsUserId);
  return sendBotReply(body, [], `Disconnected. Type \`connect\` to link a Jira account again.`);
}

// CONSENT: hand a Teams/Azure admin the one-time admin-consent link their organization needs
// to approve before anyone there can use this bot — only relevant once installed outside
// Clovity's own tenant, but harmless (and useful) to have ready regardless.
export async function handleConsent(body) {
  const url = `https://login.microsoftonline.com/organizations/adminconsent?client_id=${process.env.MS_CLIENT_ID}`;
  return sendBotReply(body, [buildConnectCard(
    url,
    'Admin approval needed',
    'An admin for your organization needs to approve this app once before everyone here can use it.',
    '✅ Approve for organization'
  )], undefined);
}

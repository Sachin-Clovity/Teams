import { graphGet, graphPost } from './client.js';

// tenantId (optional, threaded through every call below): which company's directory to look
// the user up in / send the chat in — see getAppToken() in auth.js for why this matters.

export async function getAadUserId(email, tenantId) {
  const data = await graphGet(`/v1.0/users/${encodeURIComponent(email)}`, tenantId);
  if (!data.id) throw new Error(`User not found in Azure AD: ${email}`);
  return data.id;
}

export async function postChatMessage(chatId, htmlContent, tenantId) {
  await graphPost(`/v1.0/chats/${chatId}/messages`, {
    body: { contentType: 'html', content: htmlContent },
  }, tenantId);
}

// Sends a 1:1 Teams DM using a fixed service-account sender (MS_NOTIFIER_EMAIL).
// Graph app-only chat creation needs two real member accounts, so automated
// notifications need a "from" identity too — same mechanism the manual startDM action uses.
// MS_NOTIFIER_EMAIL must be a real mailbox in that same tenantId, since it plays the "from" role.
export async function sendPersonalDM(toEmail, htmlContent, tenantId) {
  const notifierEmail = process.env.MS_NOTIFIER_EMAIL;
  if (!notifierEmail) { console.log('[sendPersonalDM] MS_NOTIFIER_EMAIL not set — skipping personal notification'); return; }
  try {
    const [fromId, toId] = await Promise.all([getAadUserId(notifierEmail, tenantId), getAadUserId(toEmail, tenantId)]);
    const chat = await graphPost('/v1.0/chats', {
      chatType: 'oneOnOne',
      members: [
        { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${fromId}` },
        { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${toId}` },
      ],
    }, tenantId);
    await postChatMessage(chat.id, htmlContent, tenantId);
  } catch (e) {
    console.log('[sendPersonalDM] error:', e.message);
  }
}

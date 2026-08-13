import { getBotFrameworkToken } from './auth.js';

// Posts a reply to Teams via the Azure Bot serviceUrl (proactive/conversational reply).
export async function postToTeams(activity, msgContent) {
  const serviceUrl     = activity.serviceUrl;
  const conversationId = activity.conversation?.id;
  const activityId     = activity.id;
  console.log('[postToTeams] serviceUrl:', serviceUrl);
  console.log('[postToTeams] conversationId:', conversationId);
  console.log('[postToTeams] activityId:', activityId);
  if (!serviceUrl || !conversationId) {
    console.log('[postToTeams] ERROR: missing serviceUrl or conversationId — cannot reply');
    return;
  }
  const token = await getBotFrameworkToken();
  const { fetch } = await import('@forge/api');
  const reply = {
    type: 'message',
    conversation: { id: conversationId },
    from:      { id: activity.recipient?.id, name: activity.recipient?.name },
    recipient: { id: activity.from?.id,      name: activity.from?.name },
    replyToId: activityId,
    ...msgContent,
  };
  const url = `${serviceUrl}v3/conversations/${encodeURIComponent(conversationId)}/activities/${activityId}`;
  console.log('[postToTeams] posting reply to:', url);
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(reply),
  });
  const resText = await res.text();
  console.log('[postToTeams] reply status:', res.status, '| body:', resText);
}

// Builds an adaptive-card message and posts it via serviceUrl.
export async function sendBotReply(activity, cards, text) {
  const attachments = cards?.length
    ? cards.map(c => ({ contentType: 'application/vnd.microsoft.card.adaptive', content: c }))
    : undefined;
  const msgContent = {
    ...(text        ? { text }        : {}),
    ...(attachments ? { attachments } : {}),
  };
  await postToTeams(activity, msgContent);
  return { statusCode: 200, body: JSON.stringify({}) };
}

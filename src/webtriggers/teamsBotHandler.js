import { setBotDebugLog } from '../storage/kvsStore.js';
import { handleQueryLink, handleFetchTask, handleSubmitAction, handleQuery } from '../bot/composeExtension.js';
import { handleBotMessage } from '../bot/commands.js';

// Jira link previews, message actions, compose extension search, and text commands — the Teams bot endpoint.
export async function teamsBotHandler(req) {
  try {
    console.log('[teamsBotHandler] ─── INCOMING REQUEST ───');
    console.log('[teamsBotHandler] req.body exists:', !!req.body);
    const body = req.body ? JSON.parse(req.body) : {};
    console.log('[teamsBotHandler] type:', body.type, '| name:', body.name);
    console.log('[teamsBotHandler] from:', body.from?.name, '| channel:', body.channelId);
    console.log('[teamsBotHandler] serviceUrl:', body.serviceUrl);
    console.log('[teamsBotHandler] text:', body.text);

    // Save debug info to storage (visible via UI — works on AGC where logs are restricted)
    await setBotDebugLog({
      receivedAt: new Date().toISOString(),
      type: body.type,
      name: body.name,
      text: body.text,
      from: body.from?.name,
      channelId: body.channelId,
      serviceUrl: body.serviceUrl,
      hasConversation: !!body.conversation?.id,
    });

    if (body.type === 'invoke' && body.name === 'composeExtension/queryLink')    return handleQueryLink(body);
    if (body.type === 'invoke' && body.name === 'composeExtension/fetchTask')    return handleFetchTask(body);
    if (body.type === 'invoke' && body.name === 'composeExtension/submitAction') return handleSubmitAction(body);
    if (body.type === 'invoke' && body.name === 'composeExtension/query')       return handleQuery(body);
    if (body.type === 'message') return handleBotMessage(body);

    // Default response for other activity types (ping, etc.)
    return { statusCode: 200, body: JSON.stringify({}) };
  } catch (e) {
    console.log('[teamsBotHandler] error:', e.message);
    return { statusCode: 500, body: JSON.stringify({ error: e.message }) };
  }
}

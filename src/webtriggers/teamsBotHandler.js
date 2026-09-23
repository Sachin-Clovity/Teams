import { setBotDebugLog, getMsTenantId, saveMsTenantId } from '../storage/kvsStore.js';
import { handleQueryLink, handleFetchTask, handleSubmitAction, handleQuery } from '../bot/composeExtension/index.js';
import { handleBotMessage } from '../bot/commands/index.js';
import { handleTaskFetch, handleTaskSubmit } from '../bot/cardActions/index.js';
import { verifyBotFrameworkRequest } from '../graph/botAuth.js';

// Jira link previews, message actions, compose extension search, and text commands — the Teams bot endpoint.
export async function teamsBotHandler(req) {
  try {
    console.log('[teamsBotHandler] ─── INCOMING REQUEST ───');
    console.log('[teamsBotHandler] req.body exists:', !!req.body);

    // Now enforcing — real Teams traffic (messages, wizard steps, task fetch/submit) logged
    // VALID consistently with no false negatives, so this endpoint no longer trusts
    // body.from.aadObjectId on its own; a request that doesn't carry a genuine Bot Framework
    // JWT is rejected before any command/resolver runs.
    const authCheck = await verifyBotFrameworkRequest(req.headers);
    console.log('[teamsBotHandler] Bot Framework auth check:', authCheck.valid ? 'VALID' : `INVALID (${authCheck.reason})`);
    if (!authCheck.valid) {
      console.log('[teamsBotHandler] REJECTING unverified request:', authCheck.reason);
      return { statusCode: 401, body: JSON.stringify({ error: 'Unauthorized' }) };
    }

    const body = req.body ? JSON.parse(req.body) : {};
    console.log('[teamsBotHandler] type:', body.type, '| name:', body.name);
    console.log('[teamsBotHandler] from:', body.from?.name, '| channel:', body.channelId);
    console.log('[teamsBotHandler] serviceUrl:', body.serviceUrl);
    console.log('[teamsBotHandler] text:', body.text);

    // Learn this installation's Microsoft tenant from real traffic, once — every activity
    // carries conversation.tenantId, so whichever company's Teams sends the very first message
    // becomes the tenant all our Graph calls (team/channel picker, personal DMs) use, instead
    // of the single tenant hardcoded in the MS_TENANT_ID environment variable.
    const incomingTenantId = body.conversation?.tenantId;
    if (incomingTenantId) {
      const known = await getMsTenantId();
      if (known !== incomingTenantId) {
        console.log('[teamsBotHandler] learning ms tenant id:', incomingTenantId, known ? `(was ${known})` : '(first time)');
        await saveMsTenantId(incomingTenantId);
      }
    }

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
    // Comment/Edit/Notify buttons on issueCard() open a task-module dialog — same mechanism
    // as the message-action three-dot menu, just triggered from a bot-sent card instead.
    if (body.type === 'invoke' && body.name === 'task/fetch')  return handleTaskFetch(body);
    if (body.type === 'invoke' && body.name === 'task/submit') return handleTaskSubmit(body);
    if (body.type === 'message') return handleBotMessage(body);

    // Default response for other activity types (ping, etc.)
    return { statusCode: 200, body: JSON.stringify({}) };
  } catch (e) {
    console.log('[teamsBotHandler] error:', e.message);
    return { statusCode: 500, body: JSON.stringify({ error: e.message }) };
  }
}

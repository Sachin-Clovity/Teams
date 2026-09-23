import { sendBotReply } from '../../graph/botReply.js';
import { helpText } from '../cards.js';
import { requireConnection } from '../../jira/atlassianAuth.js';
import { handleConnect, handleDisconnect, handleConsent } from './auth.js';
import { handleTypedCreate, handleCreate } from './create.js';
import { handleSearch, handleProjects, handleShow } from './read.js';
import { handleUpdate, handleComment, handleLog, handleAssign } from './write.js';

const NO_REPLY = { statusCode: 200, body: JSON.stringify({}) };

// Text commands typed directly to the bot: connect / create / bug / task / story / epic /
// search / projects / show / update / comment / log / assign / help. Each command's actual
// logic lives in a sibling file grouped by concern (auth/create/read/write) — this file is
// just the regex table that routes to them.
export async function handleBotMessage(body) {
  const text = (body.text || '').replace(/<at>[^<]*<\/at>/gi, '').trim();
  // from.id is scoped to the current conversation (personal chat vs channel vs group chat
  // each get a different value for the same person) — aadObjectId is the stable identity
  // that's the same everywhere, which is what lets `connect` (personal chat) and the
  // message-action / card buttons (channel or group chat) find the same saved auth record.
  const teamsUserId = body.from?.aadObjectId || body.from?.id;
  console.log('[botCommand] message command:', text, '| from:', teamsUserId, '| body.from:', JSON.stringify(body.from));

  try {
    return await dispatchCommand(body, text, teamsUserId);
  } catch (e) {
    // Without this, a Jira outage/rate-limit (a non-JSON error body making res.json() throw,
    // for example) propagates uncaught to teamsBotHandler's outer catch, which returns a
    // raw 500 — the user who typed the command gets no reply at all, not even an error message.
    console.log('[botCommand] unhandled error:', e.message);
    return sendBotReply(body, [], `❌ Something went wrong talking to Jira. Please try again in a moment.`);
  }
}

async function dispatchCommand(body, text, teamsUserId) {
  if (/^connect$/i.test(text))    return handleConnect(body, teamsUserId);
  if (/^disconnect$/i.test(text)) return handleDisconnect(body, teamsUserId);
  if (/^consent$/i.test(text))    return handleConsent(body);

  const typedMatch = text.match(/^(bug|task|story|epic)\s+([A-Z][A-Z0-9_]*)\s+(.+)$/i);
  if (typedMatch) return withConnection(body, teamsUserId, () => handleTypedCreate(body, teamsUserId, typedMatch));

  const createMatch = text.match(/^create\s+([A-Z][A-Z0-9_]*)\s+(.+)$/i);
  if (createMatch) return withConnection(body, teamsUserId, () => handleCreate(body, teamsUserId, createMatch));

  const searchMatch = text.match(/^search\s+(.+)$/i);
  if (searchMatch) return handleSearch(body, searchMatch);

  if (/^projects?(\s+list)?$/i.test(text)) return handleProjects(body);

  const showMatch = text.match(/^show\s+([A-Z][A-Z0-9_]*-\d+)$/i);
  if (showMatch) return handleShow(body, showMatch);

  const updateMatch = text.match(/^update\s+([A-Z][A-Z0-9_]*-\d+)\s+(.+)$/i);
  if (updateMatch) return withConnection(body, teamsUserId, () => handleUpdate(body, teamsUserId, updateMatch));

  const commentMatch = text.match(/^comment\s+([A-Z][A-Z0-9_]*-\d+)\s+(.+)$/i);
  if (commentMatch) return withConnection(body, teamsUserId, () => handleComment(body, teamsUserId, commentMatch));

  const logMatch = text.match(/^log\s+([A-Z][A-Z0-9_]*-\d+)\s+(\d+[hm](?:\s+\d+[hm])?)\s*(.*)$/i);
  if (logMatch) return withConnection(body, teamsUserId, () => handleLog(body, teamsUserId, logMatch));

  const assignMatch = text.match(/^assign\s+([A-Z][A-Z0-9_]*-\d+)\s+(\S+@\S+)$/i);
  if (assignMatch) return withConnection(body, teamsUserId, () => handleAssign(body, teamsUserId, assignMatch));

  return sendBotReply(body, [], helpText()); // HELP — default
}

// requireConnection() already sends the "connect first" card reply itself when not connected —
// callers just need to skip the actual command in that case, which is what this wraps.
async function withConnection(body, teamsUserId, run) {
  if (!(await requireConnection(body, teamsUserId))) return NO_REPLY;
  return run();
}

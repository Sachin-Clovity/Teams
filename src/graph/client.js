import { getAppToken } from './auth.js';
import { fetchWithRetry } from '../utils/http.js';

// All Graph calls use the app-level token from client credentials — no user OAuth needed.
// tenantId (optional): scopes the call to a specific company's tenant — see getAppToken().
// Every call goes through fetchWithRetry — this is the single place all Graph traffic funnels
// through (personal DMs, team/channel lookups, channel posts), so a burst of events (a bulk
// issue import firing many trigger invocations back to back) that trips Graph's rate limit
// gets absorbed here for everyone, not re-implemented per caller.

export async function graphGet(path, tenantId) {
  const token = await getAppToken(tenantId);
  const { fetch } = await import('@forge/api');
  const res = await fetchWithRetry(fetch, `https://graph.microsoft.com${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  }, `Graph GET ${path}`);
  if (!res.ok) throw new Error(`Graph GET ${path} failed ${res.status}: ${await res.text()}`);
  return res.json();
}

export async function graphPost(path, bodyObj, tenantId) {
  const token = await getAppToken(tenantId);
  const { fetch } = await import('@forge/api');
  const res = await fetchWithRetry(fetch, `https://graph.microsoft.com${path}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(bodyObj),
  }, `Graph POST ${path}`);
  if (!res.ok) throw new Error(`Graph POST ${path} failed ${res.status}: ${await res.text()}`);
  return res.json();
}

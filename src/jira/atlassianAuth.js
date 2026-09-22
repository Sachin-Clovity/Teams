import { createHmac, timingSafeEqual } from 'crypto';
import { getAtlassianAuth, saveAtlassianAuth } from '../storage/kvsStore.js';

const AUTH_BASE = 'https://auth.atlassian.com';
const API_BASE  = 'https://api.atlassian.com';
const SCOPES = ['read:jira-work', 'write:jira-work', 'read:jira-user', 'offline_access'].join(' ');

// Shared with atlassianOAuthCallback.js — the one webtrigger URL registered as this
// OAuth app's callback in the Atlassian Developer Console.
export const ATLASSIAN_REDIRECT_URI = 'https://1e22dbd0-9cce-44ac-b3b7-03d8530fb34e.hello.atlassian-dev.net/x1/UUV5ZkTMpm3qpqhcFjoATLIy3YU';

// Signs the teamsUserId into the OAuth "state" param so the callback can prove it wasn't
// tampered with — without this, anyone could hand-craft a callback URL with someone else's
// teamsUserId as state and have their own Atlassian tokens saved under that person's identity.
// Reuses the OAuth app's own client secret as the HMAC key rather than adding a new Forge
// variable — its HMAC output doesn't expose the underlying secret.
export function signState(teamsUserId) {
  const sig = createHmac('sha256', process.env.ATLASSIAN_OAUTH_CLIENT_SECRET).update(teamsUserId).digest('hex');
  return `${teamsUserId}.${sig}`;
}

// Returns the original teamsUserId if the signature checks out, otherwise null.
export function verifyState(signedState) {
  if (!signedState) return null;
  const dot = signedState.lastIndexOf('.');
  if (dot === -1) return null;
  const teamsUserId = signedState.slice(0, dot);
  const providedSig = signedState.slice(dot + 1);
  const expectedSig = createHmac('sha256', process.env.ATLASSIAN_OAUTH_CLIENT_SECRET).update(teamsUserId).digest('hex');
  const a = Buffer.from(providedSig, 'hex');
  const b = Buffer.from(expectedSig, 'hex');
  if (a.length !== b.length || !timingSafeEqual(a, b)) return null;
  return teamsUserId;
}

// Where we send the user to sign in and pick which of their Jira sites to authorize.
export function buildAuthorizationUrl(redirectUri, teamsUserId) {
  const params = new URLSearchParams({
    audience: 'api.atlassian.com',
    client_id: process.env.ATLASSIAN_OAUTH_CLIENT_ID,
    scope: SCOPES,
    redirect_uri: redirectUri,
    response_type: 'code',
    prompt: 'consent',
    state: teamsUserId ? signState(teamsUserId) : '',
  });
  return `${AUTH_BASE}/authorize?${params.toString()}`;
}

export async function exchangeCodeForToken(code, redirectUri) {
  const { fetch } = await import('@forge/api');
  const res = await fetch(`${AUTH_BASE}/oauth/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      grant_type: 'authorization_code',
      client_id: process.env.ATLASSIAN_OAUTH_CLIENT_ID,
      client_secret: process.env.ATLASSIAN_OAUTH_CLIENT_SECRET,
      code,
      redirect_uri: redirectUri,
    }),
  });
  const data = await res.json();
  if (!data.access_token) throw new Error(`Atlassian token exchange failed: ${JSON.stringify(data)}`);
  return data;
}

async function refreshAccessToken(refreshToken) {
  const { fetch } = await import('@forge/api');
  const res = await fetch(`${AUTH_BASE}/oauth/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      grant_type: 'refresh_token',
      client_id: process.env.ATLASSIAN_OAUTH_CLIENT_ID,
      client_secret: process.env.ATLASSIAN_OAUTH_CLIENT_SECRET,
      refresh_token: refreshToken,
    }),
  });
  const data = await res.json();
  if (!data.access_token) throw new Error(`Atlassian token refresh failed: ${JSON.stringify(data)}`);
  return data;
}

// Every Jira (and Confluence) site the signed-in Atlassian account can reach — this is the
// list the "choose your instance" screen is built from.
export async function getAccessibleResources(accessToken) {
  const { fetch } = await import('@forge/api');
  const res = await fetch(`${API_BASE}/oauth/token/accessible-resources`, {
    headers: { Authorization: `Bearer ${accessToken}`, Accept: 'application/json' },
  });
  return res.json();
}

// Valid (auto-refreshed if needed) auth record for this Teams user, or null if never connected
// or refresh failed (e.g. they revoked access) — callers should treat null as "not connected."
export async function getValidAtlassianAuth(teamsUserId) {
  const auth = await getAtlassianAuth(teamsUserId);
  if (!auth) return null;

  const now = Date.now();
  if (auth.expiresAt > now + 30_000) return auth;

  try {
    const refreshed = await refreshAccessToken(auth.refreshToken);
    const updated = {
      ...auth,
      accessToken: refreshed.access_token,
      refreshToken: refreshed.refresh_token || auth.refreshToken,
      expiresAt: now + (refreshed.expires_in ? refreshed.expires_in * 1000 : 3_600_000),
    };
    await saveAtlassianAuth(teamsUserId, updated);
    return updated;
  } catch (e) {
    console.log('[atlassianAuth] refresh failed:', e.message);
    return null;
  }
}

// Shared gate for any bot action that needs a real Jira identity — sends a "connect" prompt
// and returns null if not connected, otherwise returns the auth record.
export async function requireConnection(body, teamsUserId) {
  const { sendBotReply } = await import('../graph/botReply.js');
  const { buildConnectCard } = await import('../bot/cards.js');
  const auth = await getValidAtlassianAuth(teamsUserId);
  if (auth?.cloudId) return auth;
  const url = buildAuthorizationUrl(ATLASSIAN_REDIRECT_URI, teamsUserId);
  await sendBotReply(body, [buildConnectCard(url, 'Connect your Jira account', 'Connect first so this action runs as you, not the bot.')], undefined);
  return null;
}

// Authenticated call to this Teams user's connected Jira site — the reporter/actor on any
// write this makes is the real person, since it runs on their own OAuth token, not the app's.
export async function atlassianFetch(teamsUserId, path, options = {}) {
  const auth = await getValidAtlassianAuth(teamsUserId);
  if (!auth) throw new Error('NOT_CONNECTED');
  const { fetch } = await import('@forge/api');
  const res = await fetch(`${API_BASE}/ex/jira/${auth.cloudId}${path}`, {
    ...options,
    headers: {
      ...(options.headers || {}),
      Authorization: `Bearer ${auth.accessToken}`,
      Accept: 'application/json',
    },
  });
  return res;
}

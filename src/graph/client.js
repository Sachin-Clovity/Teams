import { getAppToken } from './auth.js';

// All Graph calls use the app-level token from client credentials — no user OAuth needed.
// tenantId (optional): scopes the call to a specific company's tenant — see getAppToken().

export async function graphGet(path, tenantId) {
  const token = await getAppToken(tenantId);
  const { fetch } = await import('@forge/api');
  const res = await fetch(`https://graph.microsoft.com${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph GET ${path} failed ${res.status}: ${await res.text()}`);
  return res.json();
}

export async function graphPost(path, bodyObj, tenantId) {
  const token = await getAppToken(tenantId);
  const { fetch } = await import('@forge/api');
  const res = await fetch(`https://graph.microsoft.com${path}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(bodyObj),
  });
  if (!res.ok) throw new Error(`Graph POST ${path} failed ${res.status}: ${await res.text()}`);
  return res.json();
}

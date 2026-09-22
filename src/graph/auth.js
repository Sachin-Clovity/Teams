// Client-credentials tokens — app-level, no user OAuth needed.
// Cached in module scope (best-effort across warm invocations) so a single
// request that makes several Graph/Bot calls doesn't re-authenticate every time.

const TOKEN_REFRESH_BUFFER_MS = 30_000;

// Keyed by tenant id ('' for the default/home tenant) — a second company's installation
// requests Graph tokens from its own tenant authority, so its cached token must never be
// handed back to a request meant for a different tenant.
let appTokenCache = {};
let botTokenCache = { token: null, expiresAt: 0 };

async function fetchClientCredentialsToken(scope, tenantId) {
  const { fetch } = await import('@forge/api');
  const body = [
    `client_id=${encodeURIComponent(process.env.MS_CLIENT_ID)}`,
    `client_secret=${encodeURIComponent(process.env.MS_CLIENT_SECRET)}`,
    `scope=${encodeURIComponent(scope)}`,
    `grant_type=client_credentials`,
  ].join('&');

  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId || process.env.MS_TENANT_ID}/oauth2/v2.0/token`,
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }
  );
  return { status: res.status, data: await res.json() };
}

// tenantId: the Microsoft 365 tenant to call Graph as. Application permissions are consented
// per-tenant, so a token from our home tenant's authority can only see our home tenant's
// directory — pass the calling installation's own learned tenant id (see kvsStore.getMsTenantId)
// to reach that company's Teams/users instead. Omit to use the MS_TENANT_ID env var (Clovity).
export async function getAppToken(tenantId) {
  const cacheKey = tenantId || '';
  const now = Date.now();
  const cached = appTokenCache[cacheKey];
  if (cached?.token && cached.expiresAt > now + TOKEN_REFRESH_BUFFER_MS) {
    return cached.token;
  }
  const { data } = await fetchClientCredentialsToken('https://graph.microsoft.com/.default', tenantId);
  if (!data.access_token) throw new Error(`Token error: ${JSON.stringify(data)}`);
  appTokenCache[cacheKey] = { token: data.access_token, expiresAt: now + (data.expires_in ? data.expires_in * 1000 : 3_600_000) };
  return data.access_token;
}

// Bot Framework token — for sending replies back to Teams via the Azure Bot serviceUrl.
export async function getBotFrameworkToken() {
  const now = Date.now();
  if (botTokenCache.token && botTokenCache.expiresAt > now + TOKEN_REFRESH_BUFFER_MS) {
    return botTokenCache.token;
  }
  console.log('[getBotToken] fetching token for tenant:', process.env.MS_TENANT_ID);
  console.log('[getBotToken] client_id present:', !!process.env.MS_CLIENT_ID);
  console.log('[getBotToken] client_secret present:', !!process.env.MS_CLIENT_SECRET);
  const { status, data } = await fetchClientCredentialsToken('https://api.botframework.com/.default');
  console.log('[getBotToken] token response status:', status, '| has token:', !!data.access_token);
  if (!data.access_token) {
    console.log('[getBotToken] ERROR:', JSON.stringify(data));
    throw new Error(`Bot token error: ${JSON.stringify(data)}`);
  }
  botTokenCache = { token: data.access_token, expiresAt: now + (data.expires_in ? data.expires_in * 1000 : 3_600_000) };
  return data.access_token;
}

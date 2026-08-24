import { createPublicKey, verify as cryptoVerify } from 'crypto';

// Verifies that an incoming Teams bot request actually carries a genuine token signed by
// Microsoft's Bot Framework, not just a POST body someone crafted by hand. Every real
// activity from Azure Bot Service includes an `Authorization: Bearer <JWT>` header signed
// with a key from this well-known endpoint — Bot Framework SDKs do this validation
// automatically; we're not using the SDK, so it has to happen here instead.
const OPENID_CONFIG_URL = 'https://login.botframework.com/v1/.well-known/openidconfiguration';
const JWKS_CACHE_MS = 24 * 60 * 60 * 1000; // signing keys rotate rarely

let jwksCache = { keys: null, fetchedAt: 0 };

function base64UrlDecode(str) {
  return Buffer.from(str.replace(/-/g, '+').replace(/_/g, '/'), 'base64');
}

function getHeader(headers, name) {
  if (!headers) return null;
  const key = Object.keys(headers).find(k => k.toLowerCase() === name.toLowerCase());
  if (!key) return null;
  const val = headers[key];
  return Array.isArray(val) ? val[0] : val;
}

async function getSigningKeys() {
  const now = Date.now();
  if (jwksCache.keys && (now - jwksCache.fetchedAt) < JWKS_CACHE_MS) return jwksCache.keys;
  const { fetch } = await import('@forge/api');
  const configRes = await fetch(OPENID_CONFIG_URL);
  const config    = await configRes.json();
  const jwksRes   = await fetch(config.jwks_uri);
  const jwks      = await jwksRes.json();
  jwksCache = { keys: jwks.keys || [], fetchedAt: now };
  return jwksCache.keys;
}

// Returns { valid: boolean, reason: string } — never throws, so callers can log-and-continue
// during the rollout period before switching this to actually block requests.
export async function verifyBotFrameworkRequest(headers) {
  const authHeader = getHeader(headers, 'authorization');
  if (!authHeader?.startsWith('Bearer ')) return { valid: false, reason: 'missing_or_malformed_auth_header' };

  const token = authHeader.slice(7);
  const parts = token.split('.');
  if (parts.length !== 3) return { valid: false, reason: 'malformed_jwt' };
  const [headerB64, payloadB64, sigB64] = parts;

  let header, payload;
  try {
    header  = JSON.parse(base64UrlDecode(headerB64).toString('utf8'));
    payload = JSON.parse(base64UrlDecode(payloadB64).toString('utf8'));
  } catch {
    return { valid: false, reason: 'unparseable_jwt' };
  }

  const now = Math.floor(Date.now() / 1000);
  if (payload.exp && payload.exp < now) return { valid: false, reason: 'expired' };
  if (payload.aud !== process.env.MS_CLIENT_ID) return { valid: false, reason: `audience_mismatch:${payload.aud}` };

  try {
    const keys = await getSigningKeys();
    const jwk  = keys.find(k => k.kid === header.kid);
    if (!jwk) return { valid: false, reason: `no_matching_key:${header.kid}` };

    const publicKey  = createPublicKey({ key: jwk, format: 'jwk' });
    const signature  = base64UrlDecode(sigB64);
    const signedData = Buffer.from(`${headerB64}.${payloadB64}`);
    const valid = cryptoVerify('RSA-SHA256', signedData, publicKey, signature);
    return valid ? { valid: true, reason: 'ok' } : { valid: false, reason: 'bad_signature' };
  } catch (e) {
    return { valid: false, reason: `verify_error:${e.message}` };
  }
}

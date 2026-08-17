import { exchangeCodeForToken, getAccessibleResources } from '../jira/atlassianAuth.js';
import { getAtlassianAuth, saveAtlassianAuth } from '../storage/kvsStore.js';

// This function's own deployed webtrigger URL — Atlassian requires an exact match against
// the callback URL registered in the OAuth app, so this gets filled in after first deploy.
const REDIRECT_URI = 'https://1e22dbd0-9cce-44ac-b3b7-03d8530fb34e.hello.atlassian-dev.net/x1/UUV5ZkTMpm3qpqhcFjoATLIy3YU';

function page(title, bodyHtml) {
  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>${title}</title>
<style>
  body { font-family: -apple-system, "Segoe UI", sans-serif; padding: 40px 24px; max-width: 460px; margin: 0 auto; color: #172B4D; text-align: center; }
  h1 { font-size: 18px; margin-bottom: 8px; }
  p { color: #6B778C; font-size: 14px; line-height: 1.5; }
  .site { display: block; padding: 12px 16px; border: 1px solid #DFE1E6; border-radius: 8px; margin: 10px 0; text-decoration: none; color: #172B4D; font-weight: 600; text-align: left; }
  .site:hover { border-color: #0052CC; color: #0052CC; }
  .success { color: #216E4E; }
  .error { color: #AE2E24; }
</style>
</head>
<body>${bodyHtml}</body>
</html>`;
  return { statusCode: 200, headers: { 'Content-Type': ['text/html; charset=utf-8'] }, body: html };
}

export async function atlassianOAuthCallback(req) {
  const qp   = req.queryParameters || {};
  const step = qp.step?.[0];

  try {
    // ── Step 2: user picked which Jira site to connect (only shown if they have more than one) ──
    if (step === 'select') {
      const teamsUserId = qp.teamsUserId?.[0];
      const cloudId     = qp.cloudId?.[0];
      const siteName    = qp.siteName?.[0] || '';
      const siteUrl     = qp.siteUrl?.[0] || '';

      const existing = await getAtlassianAuth(teamsUserId);
      if (!existing) {
        return page('Connect Jira', `<h1 class="error">Session expired</h1><p>Please go back to Teams and start connecting again.</p>`);
      }
      await saveAtlassianAuth(teamsUserId, { ...existing, cloudId, siteName, siteUrl });
      return page('Connected', `<h1 class="success">✅ Connected to ${siteName}</h1><p>You can close this window and go back to Teams.</p>`);
    }

    // ── Step 1: OAuth code exchange, right after the user approved access ──
    const code        = qp.code?.[0];
    const teamsUserId = qp.state?.[0];
    console.log('[atlassianOAuthCallback] saving auth for teamsUserId (from state):', teamsUserId);
    if (!code || !teamsUserId) {
      return page('Connect Jira', `<h1 class="error">Missing information</h1><p>Please start the connection again from Teams.</p>`);
    }

    const tokenData = await exchangeCodeForToken(code, REDIRECT_URI);
    const expiresAt = Date.now() + (tokenData.expires_in ? tokenData.expires_in * 1000 : 3_600_000);

    // Save the token immediately, site selection filled in below (or by step "select")
    await saveAtlassianAuth(teamsUserId, {
      accessToken: tokenData.access_token,
      refreshToken: tokenData.refresh_token,
      expiresAt,
      cloudId: null,
      siteName: '',
      siteUrl: '',
    });

    const resources = await getAccessibleResources(tokenData.access_token);
    const sites = (resources || []).filter(r => (r.scopes || []).length > 0);

    if (!sites.length) {
      return page('Connect Jira', `<h1 class="error">No accessible sites</h1><p>Your Atlassian account doesn't have access to any Jira site yet. Ask your admin for access, then try connecting again.</p>`);
    }

    if (sites.length === 1) {
      const s = sites[0];
      await saveAtlassianAuth(teamsUserId, {
        accessToken: tokenData.access_token,
        refreshToken: tokenData.refresh_token,
        expiresAt,
        cloudId: s.id,
        siteName: s.name,
        siteUrl: s.url,
      });
      return page('Connected', `<h1 class="success">✅ Connected to ${s.name}</h1><p>You can close this window and go back to Teams.</p>`);
    }

    const links = sites.map(s =>
      `<a class="site" href="?step=select&teamsUserId=${encodeURIComponent(teamsUserId)}&cloudId=${encodeURIComponent(s.id)}&siteName=${encodeURIComponent(s.name)}&siteUrl=${encodeURIComponent(s.url)}">${s.name}</a>`
    ).join('');
    return page('Choose your Jira site', `<h1>Choose your Jira site</h1><p>Your Atlassian account has access to more than one site — pick which one to connect:</p>${links}`);
  } catch (e) {
    console.log('[atlassianOAuthCallback] error:', e.message);
    return page('Connect Jira', `<h1 class="error">Something went wrong</h1><p>${e.message}</p>`);
  }
}

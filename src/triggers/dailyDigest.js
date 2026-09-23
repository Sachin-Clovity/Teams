import api, { route } from '@forge/api';
import { getProjectConfigIndex, getProjectConfig, getGlobalConfig } from '../storage/kvsStore.js';
import { fetchWithRetry } from '../utils/http.js';

// Scheduled trigger — daily summary of open issues, per configured project
// (or site-wide if no project-level configs exist).
export async function dailyDigest() {
  console.log('[dailyDigest] running');
  const projectKeys  = await getProjectConfigIndex();
  const globalConfig = await getGlobalConfig();

  const routes = projectKeys.length
    ? await Promise.all(projectKeys.map(async k => ({ projectKey: k, config: await getProjectConfig(k) })))
    : (globalConfig?.webhookUrl ? [{ projectKey: null, config: globalConfig }] : []);

  const { fetch } = await import('@forge/api');

  for (const { projectKey, config } of routes) {
    if (!config?.webhookUrl) continue;
    const jql = projectKey
      ? `project = ${projectKey} AND statusCategory != Done ORDER BY updated DESC`
      : `statusCategory != Done ORDER BY updated DESC`;
    try {
      const res    = await api.asApp().requestJira(route`/rest/api/3/search/jql?jql=${jql}&maxResults=5&fields=summary,status`);
      const data   = await res.json();
      const issues = data.issues || [];
      const total  = data.total ?? issues.length;
      if (!total) { console.log('[dailyDigest] no open issues for', projectKey || 'site'); continue; }
      const list = issues.map(i => `• **${i.key}** — ${i.fields?.summary || ''} (${i.fields?.status?.name || 'Unknown'})`).join('\n');
      const payload = { text: `📊 **Daily Digest${projectKey ? ` — ${projectKey}` : ''}**: ${total} open issue(s)\n\n${list}` };
      const res2 = await fetchWithRetry(fetch, config.webhookUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) }, `daily digest for ${projectKey || 'site'}`);
      console.log('[dailyDigest] posted for', projectKey || 'site', '| status:', res2.status);
    } catch (e) {
      console.log('[dailyDigest] error for', projectKey || 'site', ':', e.message);
    }
  }
}

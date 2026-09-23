import { graphGet } from '../graph/client.js';
import { isSiteAdmin } from '../jira/utils.js';
import {
  getGlobalConfig, saveGlobalConfig, clearGlobalConfig,
  getNotificationSettings, saveNotificationSettings,
  getBotDebugLog, getSyncDebugLog, getMsTenantId,
} from '../storage/kvsStore.js';

export function registerAdminResolvers(resolver) {
  resolver.define('getTeams', async () => {
    // Use whichever Microsoft tenant this installation has actually talked to (learned from
    // real bot traffic) so a second company sees their own Teams, not Clovity's — falls back
    // to the MS_TENANT_ID env var until their bot has said anything to us yet.
    const tenantId = await getMsTenantId();
    console.log('[getTeams] called | tenantId:', tenantId || '(default)');
    const data = await graphGet('/v1.0/teams', tenantId);
    const teams = (data.value || []).map(t => ({ id: t.id, displayName: t.displayName }));
    console.log('[getTeams] count:', teams.length);
    return { teams };
  });

  resolver.define('getChannels', async ({ payload }) => {
    const tenantId = await getMsTenantId();
    console.log('[getChannels] teamId:', payload.teamId, '| tenantId:', tenantId || '(default)');
    const data = await graphGet(`/v1.0/teams/${payload.teamId}/channels`, tenantId);
    const channels = (data.value || []).map(c => ({ id: c.id, displayName: c.displayName }));
    return { channels };
  });

  resolver.define('saveConfig', async ({ payload }) => {
    if (!await isSiteAdmin()) throw new Error('Only Jira site administrators can change the global Teams channel.');
    if (!payload.teamId || !payload.channelId) throw new Error('teamId and channelId are required');
    await saveGlobalConfig(payload);
    return { success: true };
  });

  resolver.define('getConfig', async () => {
    const config = await getGlobalConfig();
    return { config: config || null };
  });

  resolver.define('clearConfig', async () => {
    if (!await isSiteAdmin()) throw new Error('Only Jira site administrators can change the global Teams channel.');
    await clearGlobalConfig();
    return { success: true };
  });

  resolver.define('testConnection', async () => {
    const config = await getGlobalConfig();
    if (!config?.webhookUrl) return { success: false, error: 'No webhook URL configured.' };
    const { fetch } = await import('@forge/api');
    const res = await fetch(config.webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ "@type": "MessageCard", "@context": "http://schema.org/extensions", "themeColor": "0078D4", "summary": "Test", "sections": [{ "activityTitle": "✅ **Jira–Teams connector is active!**", "markdown": true }] }),
    });
    if (!res.ok) return { success: false, error: `Webhook ${res.status}: ${await res.text()}` };
    return { success: true };
  });

  resolver.define('getWebhookUrl', async ({ context }) => {
    const appId = context?.appId || '37fbf02b-1217-4b50-97f9-3565740801bb';
    const envId = context?.environmentId || 'development';
    return { url: `https://webhook.forge.atlassian.com/automation-teams-webhook/${appId}/${envId}` };
  });

  resolver.define('sendCustomMessage', async ({ payload }) => {
    if (!await isSiteAdmin()) throw new Error('Only Jira site administrators can broadcast to the Teams channel.');
    const config = await getGlobalConfig();
    if (!config?.webhookUrl) return { success: false, error: 'No webhook URL configured.' };
    const { message } = payload;
    if (!message?.trim()) return { success: false, error: 'Message cannot be empty.' };
    const { fetch } = await import('@forge/api');
    const res = await fetch(config.webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: message.trim() }),
    });
    if (!res.ok) return { success: false, error: `Webhook ${res.status}: ${await res.text()}` };
    return { success: true };
  });

  resolver.define('getNotificationSettings', async () => {
    const settings = await getNotificationSettings();
    return { settings };
  });

  resolver.define('saveNotificationSettings', async ({ payload }) => {
    if (!await isSiteAdmin()) throw new Error('Only Jira site administrators can change notification settings.');
    await saveNotificationSettings(payload.settings);
    return { success: true };
  });

  resolver.define('getBotDebug', async () => {
    const debug = await getBotDebugLog();
    return { debug: debug || 'No bot activity yet' };
  });

  resolver.define('getSyncDebug', async () => {
    const debug = await getSyncDebugLog();
    return { debug: debug || 'No Jira trigger activity yet' };
  });

  resolver.define('getAuthStatus', async () => {
    console.log('[authDebug] checking app-level token');
    try {
      const tenantId = await getMsTenantId();
      const data = await graphGet('/v1.0/organization', tenantId);
      return { authenticated: true, org: data.value?.[0]?.displayName || 'OK', tenantId: tenantId || null };
    } catch (e) {
      console.log('[authDebug] ERROR:', e.message);
      return { authenticated: false, error: e.message };
    }
  });
}

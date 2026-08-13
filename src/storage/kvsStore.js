import { kvs } from '@forge/kvs';
import { DEFAULT_FIELDS } from '../config/constants.js';

const GLOBAL_CONFIG_KEY         = 'teams-channel-config';
const NOTIFICATION_SETTINGS_KEY = 'notification-settings';
const PROJECT_CONFIG_INDEX_KEY  = 'project-config-index';
const BOT_DEBUG_LOG_KEY         = 'bot-debug-log';

const projectConfigKey  = projectKey => `project-config:${projectKey}`;
const personalConfigKey = accountId => `personal-config:${accountId}`;

// ── Global channel config (fallback destination when no project override exists) ──
export async function getGlobalConfig() {
  return kvs.get(GLOBAL_CONFIG_KEY);
}

export async function saveGlobalConfig({ teamId, channelId, teamName, channelName, webhookUrl, fields }) {
  await kvs.set(GLOBAL_CONFIG_KEY, {
    teamId, channelId,
    teamName: teamName || '', channelName: channelName || '',
    webhookUrl: webhookUrl || '',
    fields: fields?.length ? fields : DEFAULT_FIELDS,
  });
}

export async function clearGlobalConfig() {
  await kvs.delete(GLOBAL_CONFIG_KEY);
}

// ── Global event-type notification toggles (created/updated/deleted/commented) ──
export async function getNotificationSettings() {
  return (await kvs.get(NOTIFICATION_SETTINGS_KEY)) || { created: true, updated: true, deleted: false, commented: false };
}

export async function saveNotificationSettings(settings) {
  await kvs.set(NOTIFICATION_SETTINGS_KEY, settings);
}

// ── Per-project config: channel routing + notification filters + fields ────
export async function getProjectConfigIndex() {
  return (await kvs.get(PROJECT_CONFIG_INDEX_KEY)) || [];
}

export async function getProjectConfig(projectKey) {
  return kvs.get(projectConfigKey(projectKey));
}

export async function saveProjectConfig(projectKey, { teamId, channelId, teamName, channelName, webhookUrl, filters, fields }) {
  await kvs.set(projectConfigKey(projectKey), {
    teamId: teamId || '', channelId: channelId || '',
    teamName: teamName || '', channelName: channelName || '',
    webhookUrl: webhookUrl || '',
    filters: {
      issueTypes: filters?.issueTypes || [],
      statuses:   filters?.statuses   || [],
      priorities: filters?.priorities || [],
    },
    fields: fields?.length ? fields : DEFAULT_FIELDS,
  });
  const index = await getProjectConfigIndex();
  if (!index.includes(projectKey)) await kvs.set(PROJECT_CONFIG_INDEX_KEY, [...index, projectKey]);
}

export async function clearProjectConfig(projectKey) {
  await kvs.delete(projectConfigKey(projectKey));
  const index = await getProjectConfigIndex();
  await kvs.set(PROJECT_CONFIG_INDEX_KEY, index.filter(k => k !== projectKey));
}

// ── Per-user personal notification preferences ──────────────────────────────
export async function getPersonalConfig(accountId) {
  return (await kvs.get(personalConfigKey(accountId))) || { dmOnAssigned: true, dmOnStatusChange: false };
}

export async function savePersonalConfig(accountId, settings) {
  await kvs.set(personalConfigKey(accountId), settings);
}

// ── Bot debug log — visible via admin UI (works on AGC where logs are restricted) ──
export async function getBotDebugLog() {
  return kvs.get(BOT_DEBUG_LOG_KEY);
}

export async function setBotDebugLog(data) {
  await kvs.set(BOT_DEBUG_LOG_KEY, data);
}

import api, { route } from '@forge/api';
import { getProjectConfig, saveProjectConfig, clearProjectConfig } from '../storage/kvsStore.js';

export function registerProjectResolvers(resolver) {
  resolver.define('getProjectConfig', async ({ payload }) => {
    const config = await getProjectConfig(payload.projectKey);
    return { config: config || null };
  });

  resolver.define('saveProjectConfig', async ({ payload }) => {
    if (!payload.projectKey) throw new Error('projectKey is required');
    await saveProjectConfig(payload.projectKey, payload);
    return { success: true };
  });

  resolver.define('clearProjectConfig', async ({ payload }) => {
    await clearProjectConfig(payload.projectKey);
    return { success: true };
  });

  // Issue types / statuses / priorities available for this project, for the filter UI
  resolver.define('getProjectMeta', async ({ payload }) => {
    const { projectKey } = payload;
    const metaRes  = await api.asApp().requestJira(route`/rest/api/3/issue/createmeta?projectKeys=${projectKey}&expand=projects.issuetypes`);
    const metaData = await metaRes.json();
    const issueTypes = (metaData.projects?.[0]?.issuetypes || []).map(t => t.name);

    const statusRes  = await api.asApp().requestJira(route`/rest/api/3/project/${projectKey}/statuses`);
    const statusData = await statusRes.json();
    const statusSet  = new Set();
    (statusData || []).forEach(it => (it.statuses || []).forEach(s => statusSet.add(s.name)));

    const priorityRes  = await api.asApp().requestJira(route`/rest/api/3/priority`);
    const priorityData = await priorityRes.json();
    const priorities = (priorityData || []).map(p => p.name);

    return { issueTypes, statuses: [...statusSet], priorities };
  });
}

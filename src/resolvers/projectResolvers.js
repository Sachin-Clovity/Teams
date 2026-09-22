import api, { route } from '@forge/api';
import { getProjectConfig, saveProjectConfig, clearProjectConfig } from '../storage/kvsStore.js';
import { isProjectAdmin } from '../jira/utils.js';

export function registerProjectResolvers(resolver) {
  resolver.define('getProjectConfig', async ({ payload }) => {
    const config = await getProjectConfig(payload.projectKey);
    return { config: config || null };
  });

  resolver.define('saveProjectConfig', async ({ payload }) => {
    if (!payload.projectKey) throw new Error('projectKey is required');
    if (!await isProjectAdmin(payload.projectKey)) throw new Error('Only administrators of this project can change its Teams settings.');
    await saveProjectConfig(payload.projectKey, payload);
    return { success: true };
  });

  resolver.define('clearProjectConfig', async ({ payload }) => {
    if (!await isProjectAdmin(payload.projectKey)) throw new Error('Only administrators of this project can change its Teams settings.');
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

    // Components are project-specific (unlike statuses/priorities) — some projects have none,
    // in which case the filter UI just shows an empty, harmless list.
    const componentRes  = await api.asApp().requestJira(route`/rest/api/3/project/${projectKey}/components`);
    const componentData = await componentRes.json();
    const components = (componentData || []).map(c => c.name);

    return { issueTypes, statuses: [...statusSet], priorities, components };
  });
}

import api, { route } from '@forge/api';
import { getProjectConfig, getProjectConfigIndex, saveProjectConfig, clearProjectConfig } from '../storage/kvsStore.js';
import { isProjectAdmin, isSiteAdmin } from '../jira/utils.js';

// A site admin manages routing for every project centrally from the global Admin page; a
// project admin can only manage their own project's routing from that project's own settings
// page. Either is enough — this is an OR, not a stack of two separate checks.
async function canManageProjectRouting(projectKey) {
  return (await isSiteAdmin()) || (await isProjectAdmin(projectKey));
}

export function registerProjectResolvers(resolver) {
  resolver.define('getProjectConfig', async ({ payload }) => {
    const config = await getProjectConfig(payload.projectKey);
    return { config: config || null };
  });

  resolver.define('saveProjectConfig', async ({ payload }) => {
    if (!payload.projectKey) throw new Error('projectKey is required');
    if (!await canManageProjectRouting(payload.projectKey)) throw new Error('Only a site admin or this project\'s admin can change its Teams settings.');
    await saveProjectConfig(payload.projectKey, payload);
    return { success: true };
  });

  resolver.define('clearProjectConfig', async ({ payload }) => {
    if (!await canManageProjectRouting(payload.projectKey)) throw new Error('Only a site admin or this project\'s admin can change its Teams settings.');
    await clearProjectConfig(payload.projectKey);
    return { success: true };
  });

  // Every project → channel mapping currently saved, for the Admin page's central mapping
  // list — reading this list back is harmless for anyone (same visibility as the rest of the
  // global Admin page), only the writes above are gated.
  resolver.define('getAllProjectConfigs', async () => {
    const keys = await getProjectConfigIndex();
    const configs = await Promise.all(keys.map(async key => ({ projectKey: key, config: await getProjectConfig(key) })));
    return { configs: configs.filter(c => c.config) };
  });

  // Project picker for the mapping form — every project on the site, name + key only.
  resolver.define('searchProjects', async () => {
    const res  = await api.asApp().requestJira(route`/rest/api/3/project/search?maxResults=200&orderBy=name`);
    const data = await res.json();
    return { projects: (data.values || []).map(p => ({ key: p.key, name: p.name })) };
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

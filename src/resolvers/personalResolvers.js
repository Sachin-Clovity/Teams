import api, { route } from '@forge/api';
import { getPersonalConfig, savePersonalConfig } from '../storage/kvsStore.js';

export function registerPersonalResolvers(resolver) {
  resolver.define('getPersonalConfig', async ({ context }) => {
    const settings = await getPersonalConfig(context.accountId);
    return { settings };
  });

  resolver.define('savePersonalConfig', async ({ payload, context }) => {
    await savePersonalConfig(context.accountId, payload.settings);
    return { success: true };
  });

  resolver.define('getPriorities', async () => {
    const res  = await api.asApp().requestJira(route`/rest/api/3/priority`);
    const data = await res.json();
    return { priorities: (data || []).map(p => ({ id: p.id, name: p.name })) };
  });
}

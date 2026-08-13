import resolver from './resolvers/index.js';

export const handler = resolver.getDefinitions();

export { teamsBotHandler } from './webtriggers/teamsBotHandler.js';
export { automationHandler } from './webtriggers/automationHandler.js';
export { jiraSync } from './triggers/jiraSync.js';
export { dailyDigest } from './triggers/dailyDigest.js';

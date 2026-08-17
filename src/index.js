import resolver from './resolvers/index.js';

export const handler = resolver.getDefinitions();

export { teamsBotHandler } from './webtriggers/teamsBotHandler.js';
export { automationHandler } from './webtriggers/automationHandler.js';
export { teamsTabData } from './webtriggers/teamsTabData.js';
export { teamsTabContent } from './webtriggers/teamsTabContent.js';
export { teamsTabConfig } from './webtriggers/teamsTabConfig.js';
export { atlassianOAuthCallback } from './webtriggers/atlassianOAuthCallback.js';
export { jiraSync } from './triggers/jiraSync.js';
export { dailyDigest } from './triggers/dailyDigest.js';

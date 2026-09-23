// Compose extension entrypoints — split by concern into sibling files (link unfurling, message
// action task forms, message action submits + the create-issue wizard, and search) instead of
// one file that mixed all four. teamsBotHandler.js imports from this barrel unchanged.
export { handleQueryLink } from './queryLink.js';
export { handleFetchTask } from './fetchTask.js';
export { handleSubmitAction } from './submitAction.js';
export { handleQuery } from './query.js';

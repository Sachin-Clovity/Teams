import { issueCard, buildCreateIssueProjectCard, buildCreateIssueTypeCard, buildCreateIssueDetailsCard } from '../cards.js';
import { sendBotReply } from '../../graph/botReply.js';
import { atlassianFetch } from '../../jira/atlassianAuth.js';
import { continueWith, requireConnectionOrPrompt, taskMessage } from './helpers.js';

// Fields every type always needs, or that we ask for separately (summary/description) or
// handle specially (parent/epic) — everything else required gets rendered as a generic field.
const HANDLED_FIELD_KEYS = new Set(['summary', 'description', 'project', 'issuetype', 'reporter', 'parent']);

async function siteStep(data, teamsUserId) {
  const auth = await requireConnectionOrPrompt(teamsUserId);
  if (auth.errorResponse) return auth.errorResponse;
  const projRes  = await atlassianFetch(teamsUserId, '/rest/api/3/project/search?maxResults=50&orderBy=name');
  const projData = await projRes.json();
  const projects = (projData.values || []).map(p => ({ key: p.key, name: p.name }));
  return continueWith(buildCreateIssueProjectCard(data.prefillText || '', data.siteId, projects));
}

async function projectStep(data, teamsUserId) {
  const auth = await requireConnectionOrPrompt(teamsUserId);
  if (auth.errorResponse) return auth.errorResponse;
  const { projectKey, siteId, prefillText } = data;
  if (!projectKey) return taskMessage('❌ Please select a project.');
  const metaRes   = await atlassianFetch(teamsUserId, `/rest/api/3/issue/createmeta?projectKeys=${projectKey}&expand=projects.issuetypes`);
  const metaData  = await metaRes.json();
  const issueTypes = (metaData.projects?.[0]?.issuetypes || []).map(t => ({ id: t.id, name: t.name }));
  return continueWith(buildCreateIssueTypeCard(prefillText || '', siteId, projectKey, issueTypes));
}

async function typeStep(data, teamsUserId) {
  const auth = await requireConnectionOrPrompt(teamsUserId);
  if (auth.errorResponse) return auth.errorResponse;
  const { projectKey, siteId, prefillText, issueTypeId } = data;
  if (!issueTypeId) return taskMessage('❌ Please select a work item type.');

  const metaRes  = await atlassianFetch(teamsUserId, `/rest/api/3/issue/createmeta?projectKeys=${projectKey}&issuetypeIds=${issueTypeId}&expand=projects.issuetypes.fields`);
  const metaData = await metaRes.json();
  const issuetype = metaData.projects?.[0]?.issuetypes?.[0];
  if (!issuetype) return taskMessage('❌ Could not load that work item type. Please try again.');

  const fieldList = Object.values(issuetype.fields || {});
  // Classic (company-managed) projects expose a distinct "Epic Link" custom field; team-managed
  // projects use the plain "parent" field for the same purpose (and require it for sub-tasks).
  const epicField = fieldList.find(f => f.schema?.custom?.includes('gh-epic-link')) || fieldList.find(f => f.key === 'parent');
  const requiredExtras = fieldList.filter(f => f.required && !HANDLED_FIELD_KEYS.has(f.key) && f.key !== epicField?.key);
  const extraFields = requiredExtras.slice(0, 5).map(f => ({ id: f.key, name: f.name })); // cap so the card stays a reasonable size
  if (requiredExtras.length > 5) {
    console.log('[composeExtension] more than 5 extra required fields for', issuetype.name, '— only rendering the first 5');
  }

  return continueWith(buildCreateIssueDetailsCard(prefillText || '', siteId, projectKey, issuetype.id, issuetype.name, !!issuetype.subtask, issuetype.subtask ? null : epicField, extraFields));
}

async function detailsStep(body, data, teamsUserId) {
  const auth = await requireConnectionOrPrompt(teamsUserId);
  if (auth.errorResponse) return auth.errorResponse;
  const { projectKey, summary, description, issueTypeId, issueTypeName, isSubtask, epicFieldId, extraFieldIds, parentKey, epicKey } = data;
  console.log('[composeExtension] submitAction — project:', projectKey, '| summary:', summary);

  if (!projectKey || !summary) return taskMessage('❌ Project and Summary are required.');
  if (isSubtask && !parentKey?.trim()) return taskMessage('❌ Sub-tasks need a parent issue key.');

  const fields = {
    project:   { key: projectKey.trim().toUpperCase() },
    summary:   summary.trim(),
    issuetype: issueTypeId ? { id: issueTypeId } : { name: issueTypeName || 'Task' },
  };
  if (description?.trim()) {
    fields.description = {
      type: 'doc', version: 1,
      content: [{ type: 'paragraph', content: [{ type: 'text', text: description.trim() }] }],
    };
  }
  if (isSubtask) {
    fields.parent = { key: parentKey.trim().toUpperCase() };
  } else if (epicFieldId && epicKey?.trim()) {
    // "parent" (team-managed) takes an issue reference; the classic Epic Link custom field
    // takes a plain issue key string — same user input, two different shapes on the wire.
    fields[epicFieldId] = epicFieldId === 'parent' ? { key: epicKey.trim().toUpperCase() } : epicKey.trim().toUpperCase();
  }
  (extraFieldIds ? extraFieldIds.split(',').filter(Boolean) : []).forEach(fieldId => {
    const value = data[`custom_${fieldId}`];
    if (value?.trim()) fields[fieldId] = value.trim();
  });

  const createRes = await atlassianFetch(teamsUserId, '/rest/api/3/issue', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields }),
  });
  const created = await createRes.json();
  console.log('[composeExtension] submitAction — created key:', created.key);

  if (created.key) {
    const issueUrl = `${process.env.JIRA_BASE_URL}/browse/${created.key}`;
    // Best-effort: post the real issue card into the conversation, matching the official app.
    // This only succeeds if the bot is already a member of that conversation — a message action
    // can be invoked on messages in chats the bot was never added to, which Bot Framework refuses
    // to post into (403 BotNotInConversationRoster). It fails silently, so the link below is a
    // guaranteed fallback the user can always reach the issue from either way.
    await sendBotReply(body, [issueCard(created.key, summary.trim(), issueTypeName || 'Task', 'To Do', 'None', 'Unassigned', issueUrl, '✅ Created')]);
    return taskMessage(`✅ ${created.key} created\n\n${issueUrl}`);
  }
  const errDetail = JSON.stringify(created.errors || created.errorMessages || created);
  return taskMessage(`❌ Failed to create issue: ${errDetail}`);
}

// Create-a-work-item wizard — Site → Project → Type → Details, one function per step.
export async function handleWizardStep(body, data, teamsUserId) {
  const { wizardStep } = data;
  if (wizardStep === 'site')    return siteStep(data, teamsUserId);
  if (wizardStep === 'project') return projectStep(data, teamsUserId);
  if (wizardStep === 'type')    return typeStep(data, teamsUserId);
  return detailsStep(body, data, teamsUserId); // wizardStep === 'details' — final step, creates the issue
}

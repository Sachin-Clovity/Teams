import api, { route } from '@forge/api';

export function parseTimeToSeconds(str) {
  let seconds = 0;
  const hours = str?.match(/(\d+)\s*h/i);
  const mins  = str?.match(/(\d+)\s*m/i);
  if (hours) seconds += parseInt(hours[1], 10) * 3600;
  if (mins)  seconds += parseInt(mins[1],  10) * 60;
  return seconds;
}

export function extractTextFromADF(node) {
  if (!node) return '';
  if (node.type === 'text') return node.text || '';
  if (node.content) return node.content.map(extractTextFromADF).join('');
  return '';
}

export async function getJiraUserEmail(accountId) {
  const res  = await api.asApp().requestJira(route`/rest/api/3/user?accountId=${accountId}`);
  const data = await res.json();
  return data.emailAddress || null;
}

// accountIds of everyone watching this issue (excludes nothing — caller decides who to skip)
export async function getIssueWatchers(issueKey) {
  const res  = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}/watchers`);
  const data = await res.json();
  return (data.watchers || []).map(w => w.accountId).filter(Boolean);
}

// accountIds of everyone @mentioned in an ADF document (comment body, etc.)
export function extractMentionedAccountIds(node, out = []) {
  if (!node) return out;
  if (node.type === 'mention' && node.attrs?.id) out.push(node.attrs.id);
  if (node.content) node.content.forEach(child => extractMentionedAccountIds(child, out));
  return out;
}

// Builds a comment body for POST /issue/{key}/comment. The `sd.public.comment` property only
// means anything on a Jira Service Management request (it marks the comment internal-only,
// hidden from the customer on the portal) — Jira ignores it on a plain Software/Business
// issue, so it's always safe to include. Without it, a comment posted through the bot or issue
// panel defaults to customer-visible on JSM, which risks leaking an internal note to the
// reporter; defaulting every comment we create to internal is the safer failure mode.
export function commentBody(text) {
  return {
    body: { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text }] }] },
    properties: [{ key: 'sd.public.comment', value: { internal: true } }],
  };
}

// Most recent comment on an issue, or null if it has none
export async function getLatestComment(issueKey) {
  const res  = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}/comment?maxResults=1&orderBy=-created`);
  const data = await res.json();
  return data.comments?.[0] || null;
}

// Forge's invoke() only proves WHO is calling (a real logged-in user on this site) — it does
// not check WHAT they're allowed to do. jira:globalPage and jira:projectSettingsPage are just
// reachable by any site user who knows the URL/resolver name, so writes that change site-wide
// or project-wide behavior (channel routing, webhooks) need an explicit permission check here,
// using the caller's own Jira session via api.asUser() — never api.asApp(), which would report
// the app's own (much broader) permissions instead of the actual caller's.
export async function isSiteAdmin() {
  const res  = await api.asUser().requestJira(route`/rest/api/3/mypermissions?permissions=ADMINISTER`);
  const data = await res.json();
  return !!data.permissions?.ADMINISTER?.havePermission;
}

export async function isProjectAdmin(projectKey) {
  const res  = await api.asUser().requestJira(route`/rest/api/3/mypermissions?projectKey=${projectKey}&permissions=ADMINISTER_PROJECTS`);
  const data = await res.json();
  return !!data.permissions?.ADMINISTER_PROJECTS?.havePermission;
}

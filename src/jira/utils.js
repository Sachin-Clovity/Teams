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

// Most recent comment on an issue, or null if it has none
export async function getLatestComment(issueKey) {
  const res  = await api.asApp().requestJira(route`/rest/api/3/issue/${issueKey}/comment?maxResults=1&orderBy=-created`);
  const data = await res.json();
  return data.comments?.[0] || null;
}

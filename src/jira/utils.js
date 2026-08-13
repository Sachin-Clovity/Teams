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

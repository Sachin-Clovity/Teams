import api, { route } from '@forge/api';

// JSON API the tab's client-side JS calls to populate its issue list.
// ?jql=<JQL> takes priority; otherwise ?email=<teams-user-email> resolves to a Jira
// accountId and defaults to "assigned to me"; with neither, falls back to recently updated.
export async function teamsTabData(req) {
  try {
    let jql   = req.queryParameters?.jql?.[0];
    const email = req.queryParameters?.email?.[0];

    if (!jql && email) {
      const userRes = await api.asApp().requestJira(route`/rest/api/3/user/search?query=${email}&maxResults=1`);
      const users   = await userRes.json();
      const user    = Array.isArray(users) ? users[0] : null;
      jql = user ? `assignee = "${user.accountId}" ORDER BY updated DESC` : null;
    }
    if (!jql) jql = 'updated >= "1970/01/01" ORDER BY updated DESC';

    const res  = await api.asApp().requestJira(route`/rest/api/3/search/jql?jql=${jql}&maxResults=20&fields=summary,status,priority,issuetype,assignee`);
    const data = await res.json();

    if (data.errorMessages?.length) {
      return { statusCode: 200, headers: { 'Content-Type': ['application/json'] }, body: JSON.stringify({ error: data.errorMessages.join(' '), issues: [] }) };
    }

    const issues = (data.issues || []).map(i => {
      const f = i.fields || {};
      return {
        key: i.key,
        summary: f.summary || '',
        status: f.status?.name || 'Unknown',
        priority: f.priority?.name || 'None',
        issueType: f.issuetype?.name || 'Issue',
        assignee: f.assignee?.displayName || 'Unassigned',
        url: `${process.env.JIRA_BASE_URL}/browse/${i.key}`,
      };
    });

    return { statusCode: 200, headers: { 'Content-Type': ['application/json'] }, body: JSON.stringify({ issues }) };
  } catch (e) {
    return { statusCode: 200, headers: { 'Content-Type': ['application/json'] }, body: JSON.stringify({ error: e.message, issues: [] }) };
  }
}

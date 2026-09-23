import { atlassianFetch } from '../../jira/atlassianAuth.js';
import { extractTextFromADF } from '../../jira/utils.js';
import { taskContinue, taskMessage, checkConnection } from './helpers.js';

export async function fetchEdit(teamsUserId, issueKey) {
  const conn = await checkConnection(teamsUserId);
  if (conn.response) return conn.response;

  // Pulled live from this Jira site — same idea as the official Jira app's edit dialog,
  // rather than a fixed guess at what priorities/assignees might exist. Labels don't get the
  // same "pick from every label on the site" treatment — sites with a lot of labels in use
  // turned that into an unreadably long checkbox wall, so it's a plain field again.
  const [issueRes, transRes, priorityRes, assignableRes] = await Promise.all([
    atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}?fields=summary,description,status,priority,assignee,duedate,labels,created,updated,reporter`),
    atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/transitions`),
    atlassianFetch(teamsUserId, `/rest/api/3/priority`),
    atlassianFetch(teamsUserId, `/rest/api/3/user/assignable/search?issueKey=${issueKey}&maxResults=50`),
  ]);
  const issue       = await issueRes.json();
  const transitions  = await transRes.json();
  const priorities   = await priorityRes.json().catch(() => []);
  const assignable   = await assignableRes.json().catch(() => []);
  const f = issue.fields || {};

  const statusChoices = (transitions.transitions || []).map(t => ({ title: t.name, value: t.name }));
  if (!statusChoices.some(c => c.value === f.status?.name)) {
    statusChoices.unshift({ title: f.status?.name || 'Unknown', value: f.status?.name || '' });
  }

  const priorityChoices = (Array.isArray(priorities) ? priorities : []).map(p => ({ title: p.name, value: p.name }));
  if (!priorityChoices.length) {
    priorityChoices.push(...['Highest', 'High', 'Medium', 'Low', 'Lowest'].map(v => ({ title: v, value: v })));
  }

  const assigneeChoices = [
    { title: '— Unassigned —', value: '' },
    ...(Array.isArray(assignable) ? assignable : []).map(u => ({ title: u.displayName, value: u.accountId })),
  ];
  if (f.assignee?.accountId && !assigneeChoices.some(c => c.value === f.assignee.accountId)) {
    assigneeChoices.push({ title: f.assignee.displayName || f.assignee.accountId, value: f.assignee.accountId });
  }

  // Own OAuth-connected siteUrl, not the JIRA_BASE_URL env var — guaranteed to match
  // whichever site this specific person is actually connected to.
  const jiraUrl = `${conn.auth.siteUrl}/browse/${issueKey}`;
  const fmt = iso => iso ? new Date(iso).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' }) : '—';

  // Two-column layout — left is what people actually edit (summary, description), right is the
  // field rail, loosely mirroring the official Jira app's edit dialog. Kept deliberately short:
  // no worklog section here (still available via the log command / issue panel) and no full
  // label picker (see above) — both made the dialog long and cluttered for no real gain.
  return taskContinue(`Edit ${issueKey}`, {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: issueKey, weight: 'Bolder', size: 'Medium', spacing: 'None' },
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column', width: 3,
            items: [
              { type: 'Input.Text', id: 'summary', label: 'Summary', isRequired: true, value: f.summary || '' },
              { type: 'Input.Text', id: 'description', label: 'Description', isMultiline: true, value: extractTextFromADF(f.description) },
              { type: 'Input.Text', id: 'labels', label: 'Labels (comma-separated)', value: (f.labels || []).join(', ') },
            ],
          },
          {
            type: 'Column', width: 2,
            items: [
              { type: 'Input.ChoiceSet', id: 'status', label: 'Status', value: f.status?.name || '', choices: statusChoices },
              { type: 'Input.ChoiceSet', id: 'assignee', label: 'Assignee', value: f.assignee?.accountId || '', choices: assigneeChoices },
              { type: 'Input.ChoiceSet', id: 'priority', label: 'Priority', value: f.priority?.name || '', choices: priorityChoices },
              { type: 'Input.Date', id: 'dueDate', label: 'Due date', value: f.duedate || '' },
              { type: 'TextBlock', text: `Created ${fmt(f.created)}`, isSubtle: true, size: 'Small', wrap: true, spacing: 'Medium', separator: true },
              { type: 'TextBlock', text: `Updated ${fmt(f.updated)}`, isSubtle: true, size: 'Small', wrap: true, spacing: 'None' },
              { type: 'TextBlock', text: `Reporter ${f.reporter?.displayName || 'Unknown'}`, isSubtle: true, size: 'Small', wrap: true, spacing: 'None' },
            ],
          },
        ],
      },
    ],
    actions: [
      { type: 'Action.Submit', title: 'Save Changes', data: { cardAction: 'edit', issueKey, originalStatus: f.status?.name || '' } },
      { type: 'Action.OpenUrl', title: 'Open in Jira', url: jiraUrl },
    ],
  }, 'large', 'medium');
}

export async function submitEdit(teamsUserId, issueKey, data) {
  const conn = await checkConnection(teamsUserId);
  if (conn.response) return conn.response;

  const { status, priority, assignee, dueDate, labels, summary, description, originalStatus } = data;

  // Status change goes through the transitions endpoint, not a plain field PUT
  if (status && status !== originalStatus) {
    const transRes  = await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/transitions`);
    const transData = await transRes.json();
    const match = (transData.transitions || []).find(t => t.name === status);
    if (match) {
      await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}/transitions`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ transition: { id: match.id } }),
      });
    }
  }

  const fields = {};
  if (summary?.trim()) fields.summary = summary.trim();
  if (description !== undefined) {
    const text = description.trim();
    fields.description = { type: 'doc', version: 1, content: [{ type: 'paragraph', content: text ? [{ type: 'text', text }] : [] }] };
  }
  if (priority) fields.priority = { name: priority };
  if (dueDate !== undefined) fields.duedate = dueDate || null;
  if (labels !== undefined) fields.labels = labels.split(',').map(s => s.trim()).filter(Boolean);
  // The picker's accountId choice replaces the old email-search lookup entirely, so a typo'd
  // email can no longer silently fail to assign anyone.
  if (assignee !== undefined) fields.assignee = assignee ? { accountId: assignee } : null;

  if (Object.keys(fields).length) {
    await atlassianFetch(teamsUserId, `/rest/api/3/issue/${issueKey}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields }),
    });
  }

  return taskMessage(`✅ ${issueKey} updated`);
}

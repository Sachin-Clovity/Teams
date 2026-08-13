export function issueCard(key, summary, type, status, priority, assignee, url, label) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: label || '', size: 'Small', color: 'Accent', weight: 'Bolder' },
      { type: 'TextBlock', text: `**[${key}](${url})** — ${summary}`, wrap: true, size: 'Medium' },
      {
        type: 'FactSet', facts: [
          { title: 'Type',     value: type     || 'Issue' },
          { title: 'Status',   value: status   || 'Unknown' },
          { title: 'Priority', value: priority || 'None' },
          { title: 'Assignee', value: assignee || 'Unassigned' },
        ]
      }
    ],
    actions: [{ type: 'Action.OpenUrl', title: 'Open in Jira', url }]
  };
}

// Adaptive Card shown in the task module when user clicks three-dot → Create Jira Issue
export function buildCreateIssueCard(prefillText) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Create Jira Issue', weight: 'Bolder', size: 'Medium' },
      { type: 'Input.Text', id: 'projectKey', label: 'Project Key', placeholder: 'e.g. PROJ', isRequired: true },
      { type: 'Input.Text', id: 'summary',    label: 'Summary',     placeholder: 'Short description', value: prefillText.slice(0, 255), isRequired: true },
      {
        type: 'Input.ChoiceSet', id: 'issueType', label: 'Issue Type', value: 'Task',
        choices: [
          { title: 'Task',  value: 'Task' },
          { title: 'Bug',   value: 'Bug' },
          { title: 'Story', value: 'Story' },
          { title: 'Epic',  value: 'Epic' },
        ],
      },
      { type: 'Input.Text', id: 'description', label: 'Description', placeholder: 'Additional context', value: prefillText, isMultiline: true },
    ],
    actions: [{ type: 'Action.Submit', title: 'Create Issue' }],
  };
}

export function buildCommentCard(prefillText) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Add Comment to Jira Issue', weight: 'Bolder', size: 'Medium' },
      { type: 'Input.Text', id: 'issueKey',    label: 'Issue Key',  placeholder: 'e.g. PROJ-123', isRequired: true },
      { type: 'Input.Text', id: 'commentText', label: 'Comment',    placeholder: 'Your comment...', value: prefillText, isMultiline: true, isRequired: true },
    ],
    actions: [{ type: 'Action.Submit', title: 'Add Comment', data: { commandId: 'commentInJira' } }],
  };
}

export function buildLogTimeCard(prefillText) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Log Work Time in Jira', weight: 'Bolder', size: 'Medium' },
      { type: 'Input.Text', id: 'issueKey',      label: 'Issue Key',       placeholder: 'e.g. PROJ-123', isRequired: true },
      { type: 'Input.Text', id: 'timeSpent',     label: 'Time Spent',      placeholder: 'e.g. 2h, 30m, 1h 30m', isRequired: true },
      { type: 'Input.Text', id: 'workDescription', label: 'Work Description', placeholder: 'What did you work on?', value: prefillText, isMultiline: true },
    ],
    actions: [{ type: 'Action.Submit', title: 'Log Time', data: { commandId: 'logTimeInJira' } }],
  };
}

export function helpText() {
  return `**Jira Bot Commands:**\n\n` +
    `📁 **projects** — List all Jira projects\n` +
    `🆕 **create** \`PROJECT\` \`summary\` — Create issue (auto type)\n` +
    `🐛 **bug** \`PROJECT\` \`summary\` — Create Bug\n` +
    `✅ **task** \`PROJECT\` \`summary\` — Create Task\n` +
    `📖 **story** \`PROJECT\` \`summary\` — Create Story\n` +
    `⚡ **epic** \`PROJECT\` \`summary\` — Create Epic\n` +
    `🔍 **search** \`keyword or JQL\` — Search issues (top 8)\n` +
    `📋 **show** \`PROJ-123\` — View issue details\n` +
    `✏️ **update** \`PROJ-123\` \`Status\` — Transition issue status\n` +
    `💬 **comment** \`PROJ-123\` \`text\` — Add a comment\n` +
    `⏱️ **log** \`PROJ-123\` \`2h\` \`note\` — Log work time\n` +
    `👤 **assign** \`PROJ-123\` \`user@domain.com\` — Assign issue\n\n` +
    `_Message actions (right-click any message):_\n` +
    `• **Create issue in Jira** — turns message into an issue\n` +
    `• **Comment in Jira** — adds message as a comment\n` +
    `• **Log time in Jira** — logs work against an issue\n\n` +
    `_Search (compose bar):_\n` +
    `• Click **+** → **Jira Connector** → type to search issues\n\n` +
    `_Examples:_\n` +
    `• \`bug AITEST Login page is broken\`\n` +
    `• \`search payment bug\`\n` +
    `• \`show AITEST-40\`\n` +
    `• \`update AITEST-40 In Progress\`\n` +
    `• \`comment AITEST-40 Fixed in latest build\`\n` +
    `• \`log AITEST-40 2h Investigated root cause\`\n` +
    `• \`assign AITEST-40 dev@clovity.com\``;
}

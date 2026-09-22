const TYPE_EMOJI = { Task: '✅', Bug: '🐛', Story: '📖', Epic: '⚡', Subtask: '✅' };
const PRIORITY_ICON = { Highest: '🔴', High: '🟠', Medium: '🟡', Low: '🟢', Lowest: '🔵' };

// Matches the official Jira app's card: type icon, linked title, "Assigned to X" subtitle,
// a compact priority/status line, and Comment/Edit/Notify as the action row (Comment styled
// as the primary button) — instead of a separate "Open in Jira" link and a full FactSet.
export function issueCard(key, summary, type, status, priority, assignee, url, label, iconUrl) {
  const body = [];
  if (label) body.push({ type: 'TextBlock', text: label, size: 'Small', color: 'Accent', weight: 'Bolder' });

  body.push({
    type: 'ColumnSet',
    columns: [
      {
        type: 'Column', width: 'auto',
        items: [
          iconUrl
            ? { type: 'Image', url: iconUrl, width: '24px', height: '24px' }
            : { type: 'TextBlock', text: TYPE_EMOJI[type] || '📌', size: 'Large' },
        ],
      },
      {
        type: 'Column', width: 'stretch',
        items: [
          { type: 'TextBlock', text: `**[${key}: ${summary}](${url})**`, wrap: true, size: 'Medium' },
          { type: 'TextBlock', text: `Assigned to ${assignee || 'Unassigned'}`, isSubtle: true, spacing: 'None', wrap: true },
        ],
      },
    ],
  });

  body.push({
    type: 'TextBlock', spacing: 'Small', wrap: true,
    text: `${PRIORITY_ICON[priority] || '⚪'} ${priority || 'None'}  |  **${(status || 'Unknown').toUpperCase()}**`,
  });

  return {
    type: 'AdaptiveCard', version: '1.2',
    body,
    actions: [
      // These three open a full task-module dialog (same mechanism as the message-action
      // three-dot menu) rather than expanding inline — matches how the official Jira app does it.
      { type: 'Action.Submit', title: 'Comment', style: 'positive', data: { msteams: { type: 'task/fetch' }, cardAction: 'comment', issueKey: key } },
      { type: 'Action.Submit', title: 'Edit',    data: { msteams: { type: 'task/fetch' }, cardAction: 'edit',    issueKey: key } },
      { type: 'Action.Submit', title: '🔔 Notify', data: { msteams: { type: 'task/fetch' }, cardAction: 'notify', issueKey: key } },
    ],
  };
}

// ── "Create a work item" wizard — Site → Project → Type → Details ────────────────────
// Four separate steps (chained via task/continue responses), matching the official Jira
// app's flow, rather than one flat form. Each step's Action.Submit carries forward
// everything picked so far in its `data`, since Adaptive Card submits don't persist state.

// Step 1: choose the Jira site (we only ever have one today, but the step exists so the
// flow is ready for multi-site without a redesign later).
export function buildCreateIssueSiteCard(prefillText, siteId, siteName) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Create a work item', weight: 'Bolder', size: 'Medium' },
      {
        type: 'Input.ChoiceSet', id: 'siteId', label: 'Jira site', value: siteId,
        choices: [{ title: siteName, value: siteId }],
      },
    ],
    actions: [{ type: 'Action.Submit', title: 'Next', data: { wizardStep: 'site', prefillText } }],
  };
}

// Step 2: choose the project — a real filterable dropdown built from live project data,
// not a text field where you have to already know the exact key.
export function buildCreateIssueProjectCard(prefillText, siteId, projects) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Create a work item', weight: 'Bolder', size: 'Medium' },
      {
        type: 'Input.ChoiceSet', id: 'projectKey', label: 'Select project', style: 'filtered', isRequired: true,
        placeholder: 'Start typing to search projects',
        choices: projects.map(p => ({ title: `${p.name} (${p.key})`, value: p.key })),
      },
    ],
    actions: [{ type: 'Action.Submit', title: 'Next', data: { wizardStep: 'project', prefillText, siteId } }],
  };
}

// Step 3: pick the work item type. Split out as its own step (rather than folded into
// details) because Adaptive Cards are static once rendered — there's no way to show/hide
// the parent-issue field or epic-link field reactively based on which type gets picked on
// the same screen, so the type has to be chosen first, then the next step is built knowing
// definitively whether it's a sub-task and what fields that specific type actually needs.
export function buildCreateIssueTypeCard(prefillText, siteId, projectKey, issueTypes) {
  const choices = (issueTypes?.length ? issueTypes : [{ id: '', name: 'Task' }])
    .map(t => ({ title: t.name, value: t.id }));
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Create a work item', weight: 'Bolder', size: 'Medium' },
      { type: 'TextBlock', text: `Project: ${projectKey}`, isSubtle: true },
      { type: 'Input.ChoiceSet', id: 'issueTypeId', label: 'Select work item type', value: choices[0]?.value || '', choices, isRequired: true },
    ],
    actions: [{ type: 'Action.Submit', title: 'Next', data: { wizardStep: 'type', prefillText, siteId, projectKey } }],
  };
}

// Step 4: summary/description, plus whatever that specific work item type actually needs —
// a required parent issue for sub-tasks, an optional epic link for everything else (when the
// project has one), and any other required custom fields the project's screen scheme demands.
// epicField / extraFields come from cross-referencing the type's live createmeta field list —
// without rendering the required ones, issue creation would fail with a Jira error the user
// has no way to act on from inside this card.
export function buildCreateIssueDetailsCard(prefillText, siteId, projectKey, issueTypeId, issueTypeName, isSubtask, epicField, extraFields) {
  const body = [
    { type: 'TextBlock', text: 'Create a work item', weight: 'Bolder', size: 'Medium' },
    { type: 'TextBlock', text: `Project: ${projectKey}  ·  Type: ${issueTypeName}`, isSubtle: true },
  ];

  if (isSubtask) {
    body.push({ type: 'Input.Text', id: 'parentKey', label: 'Parent issue key', placeholder: 'e.g. PROJ-100', isRequired: true });
  } else if (epicField) {
    body.push({ type: 'Input.Text', id: 'epicKey', label: 'Epic (optional)', placeholder: 'e.g. PROJ-1' });
  }

  body.push(
    { type: 'Input.Text', id: 'summary', label: 'Summary', placeholder: 'Short description', value: prefillText.slice(0, 255), isRequired: true },
    { type: 'Input.Text', id: 'description', label: 'Description', placeholder: 'Additional context', value: prefillText, isMultiline: true },
  );

  (extraFields || []).forEach(f => {
    body.push({ type: 'Input.Text', id: `custom_${f.id}`, label: f.name, isRequired: true });
  });

  return {
    type: 'AdaptiveCard', version: '1.2',
    body,
    actions: [{
      type: 'Action.Submit', title: 'Create',
      data: {
        wizardStep: 'details', siteId, projectKey, issueTypeId, issueTypeName, isSubtask,
        epicFieldId: epicField?.id || '',
        extraFieldIds: (extraFields || []).map(f => f.id).join(','),
      },
    }],
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

// A real "sign in" screen for the chat — one button that opens the browser, instead of a
// plain markdown link buried in a text message (which reads as "copy this URL yourself").
export function buildConnectCard(url, title, description, buttonTitle) {
  return {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: title || 'Connect your Jira account', weight: 'Bolder', size: 'Medium' },
      { type: 'TextBlock', text: description || 'Sign in once so actions you take here are reported as you, not the bot.', wrap: true, isSubtle: true },
    ],
    actions: [{ type: 'Action.OpenUrl', title: buttonTitle || '🔗 Sign in to Jira', url }],
  };
}

export function helpText() {
  return `**Jira Bot Commands:**\n\n` +
    `🔗 **connect** — Link your Jira account (needed before creating issues, so they're reported as you)\n` +
    `🔌 **disconnect** — Unlink your Jira account\n` +
    `✅ **consent** — Get the admin-approval link your organization needs (once, org-wide)\n` +
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
